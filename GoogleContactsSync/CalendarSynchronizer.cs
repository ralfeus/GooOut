
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.GData.Calendar;
using System.Windows.Forms;
using Google.GData.Client;
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml;

namespace R.GoogleOutlookSync
{
    internal partial class CalendarSynchronizer : ItemTypeSynchronizer
    {
        /// <summary>
        /// Represents calendar to synchronize with
        /// </summary>
        CalendarEntry _googleCalendar;
        string _googleCalendarFeedUrl;
        string _googleCalendarName;
        /// <summary>
        /// Represents folder in Outlook mailbox containing calendar items for synchronization
        /// </summary>
        string _outlookCalendarFolderToSyncID;
        /// <summary>
        /// Contains exceptions from recurrent events
        /// </summary>
        Dictionary<string, List<EventEntry>> _googleExceptions;
        new IEnumerable<FieldHandlers> _fieldHandlers;
        /// <summary>
        /// Contains 0:00 of next date after defined interval. Next date is not inclusive
        /// </summary>
        DateTime _rangeEnd;
        DateTime _rangeStart;

        internal CalendarSynchronizer(
            string googleAuthToken,
            //Outlook.NameSpace outlookNamespace,
            string calendarFolderToSyncID)
        {
            this._googleConnection = new CalendarService(Application.ProductName);
            this._googleConnection.SetAuthenticationToken(googleAuthToken);
            this._outlookCalendarFolderToSyncID = calendarFolderToSyncID;
            this._googleExceptions = new Dictionary<string, List<EventEntry>>();

            this.LoadSettings();
            var query = new CalendarQuery(Properties.Settings.Default.GoogleUrl_Calendar);
            var calendars = GoogleUtilities.TryDo<AtomFeed>(() => this._googleConnection.Query(query));
            // first load calendars synchronization settings to have Google calendar name to synchronize
            // in case no settings are loaded first found calendar will be used
            foreach (var cal in calendars.Entries)
                Logger.Log("Found calendar: " + cal.Title.Text, EventType.Information);
            this._googleCalendar = (CalendarEntry)calendars.Entries.First(
                cal => (cal.Title.Text == this._googleCalendarName) || String.IsNullOrEmpty(this._googleCalendarName));
            this._googleCalendarFeedUrl = this._googleCalendar.Links.First(
                link => (link.Rel == "alternate") && (link.Type == "application/atom+xml")).HRef.Content;
            /// Prepare list of all available field handlers
            this.InitFieldHandlers();
        }

        /// <summary>
        /// Get last modification time for Google event. Check also time of exceptions modification 
        /// </summary>
        /// <param name="googleItem"></param>
        /// <returns></returns>
        protected override DateTime GetLastModificationTime(AtomEntry googleItem)
        {
            if (this._googleExceptions.ContainsKey(((EventEntry)googleItem).EventId))
                return
                    new DateTime(Math.Max(
                        GoogleUtilities.GetLastModificationTime(googleItem).Ticks,
                        this._googleExceptions[((EventEntry)googleItem).EventId].Max(exception => exception.Updated.Ticks)));
            else
                return GoogleUtilities.GetLastModificationTime(googleItem);
        }

        /// <summary>
        /// Finds all field handler methods and prepares list grouped by field
        /// </summary>
        private new void InitFieldHandlers()
        {
            /// 1. Get methods, which current instance have. Include non-public methods too 
            ///   (handlers are private since they are not used anywhere else)
            var methods = typeof(CalendarSynchronizer).GetMethods(BindingFlags.Instance | BindingFlags.NonPublic);
            /// 2. Filter only those methods, for which following condition is true:
            ///   Method has at least one attribute FieldComparer for comparer and FieldSetter for setter accordingly
            ///   In case comparer has no correspondent setter the exception will be thrown
            
            /// Get all comparer methods
            var comparerMethods =
                from comparerMethod in methods
                where
                    comparerMethod.GetCustomAttributes(false).FirstOrDefault(attr => attr is FieldComparerAttribute) != null
                select comparerMethod;
            /// Get all getter methods
            var getterMethods =
                from getterMethod in methods
                where
                    getterMethod.GetCustomAttributes(false).FirstOrDefault(attr => attr is FieldGetterAttribute) != null
                select getterMethod;
            /// Get all setter methods
            var setterMethods =
                from setterMethod in methods
                where
                    setterMethod.GetCustomAttributes(false).FirstOrDefault(attr => attr is FieldSetterAttribute) != null
                select setterMethod;
            /// Combine all field handle methods
            this._fieldHandlers =
                from
                    comparerMethod in comparerMethods
                    join setterMethod in setterMethods on
                        ((FieldHandlerAttribute)comparerMethod.GetCustomAttributes(false).First(attr => attr is FieldComparerAttribute)).Field equals
                        ((FieldHandlerAttribute)setterMethod.GetCustomAttributes(false).First(attr => attr is FieldSetterAttribute)).Field
                    //join getterMethod in getterMethods on
                    //    ((FieldHandler)comparerMethod.GetCustomAttributes(false).First(attr => attr is FieldComparer)).Field equals
                    //    ((FieldHandler)getterMethod.GetCustomAttributes(false).First(attr => attr is FieldGetter)).Field
                select new FieldHandlers(comparerMethod, /*getterMethod,*/ setterMethod);

        }

        protected override ComparisonResult Compare(AtomEntry googleItem, object outlookItem)
        {
            /// If IDs are same items are same. Then we can check whether they are identical
            /// If IDs are different items are different
            /// Aggregate will invoke each method in _fieldComparers list. Boolean AND will ensure all comparers return true
            //try
            //{
                if (this.CompareIDs((EventEntry)googleItem, (Outlook.AppointmentItem)outlookItem))
                {
                    var fieldsAreSame =
                        this._fieldHandlers.Aggregate(
                            true,
                            (res, handler) => res && (bool)handler.Comparer.Invoke(this, new object[] { googleItem, outlookItem }));
                    //var fieldsAreSame = true;
                    //foreach (var fieldHandler in this._fieldHandlers)
                    //    fieldsAreSame &= (bool)fieldHandler.Comparer.Invoke(this, new object[] { googleItem, outlookItem });
                    if (fieldsAreSame)
                        return ComparisonResult.Identical;
                    else
                        return ComparisonResult.SameButChanged;
                }
                else
                    return ComparisonResult.Different;
            //}
            //catch (TargetInvocationException exc)
            //{
            //    throw exc.InnerException;
            //}
        }

        protected override bool IsItemValid(AtomEntry googleItem)
        {
            /// We try to create EventSchedule of item. If due to any reason (Exception) it's impossible
            /// we treat item as invalid
            try
            {
                new EventSchedule((EventEntry)googleItem);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        protected override bool IsItemValid(object outlookItem)
        {
            /// So far we treat all Outlook items as valid
            return true;
        }

        protected override void LoadGoogleItems()
        {
            Logger.Log("Loading Google calendar items...", EventType.Information);
            var succeeded = false;
            var attemptsLeft = 3;
            Exception lastError = null;
            do
            {
                try
                {
                    var query = new EventQuery(this._googleCalendarFeedUrl);
                    query.ExtraParameters = String.Format(
                        "showdeleted=false", //start-min={0}&start-max={1}&&singleevents=true
                        this._rangeStart.ToString("yyyy-MM-ddTHH:mm:ss"),
                        this._rangeEnd.ToString("yyyy-MM-ddTHH:mm:ss"));
                    // Set the maximum number of results to return for the query.
                    // Note: A GData server may choose to provide fewer results, but will never provide
                    // more than the requested maximum.
                    query.NumberToRetrieve = 100;
                    query.StartIndex = 1;
                    AtomFeed feed;
                    // perform retrieving until there is nothing to retrieve
                    do
                    {
                        // start retrieving from next item after last retrieved
                        feed = this._googleConnection.Query(query);
                        /// Add item either as primary event entry 
                        /// or as exception of recurrence
                        foreach (EventEntry googleItem in feed.Entries)
                            if (googleItem.OriginalEvent == null)
                                this.GoogleItems.Add(googleItem);
                            else
                            {
                                if (!this._googleExceptions.ContainsKey(googleItem.OriginalEvent.IdOriginal))
                                    this._googleExceptions.Add(googleItem.OriginalEvent.IdOriginal, new List<EventEntry>());
                                this._googleExceptions[googleItem.OriginalEvent.IdOriginal].Add(googleItem);
                            }
                        query.StartIndex += feed.Entries.Count;
                        // if last query returned no entries there is nothing to retrieve anymore
                    } while (feed.Entries.Count != 0);
                    /// Create a batch update feed for Google items
                    if (this._googleBatchUpdateFeed == null)
                        this._googleBatchUpdateFeed = feed.CreateBatchFeed(GDataBatchOperationType.update);
                    succeeded = true;
                }
                catch (Exception exc)
                {
                    ErrorHandler.Handle(exc);
                    lastError = exc;
                    --attemptsLeft;
                }
            } while (!succeeded && attemptsLeft > 0);
        }

        protected override void LoadOutlookItems()
        {
            Logger.Log("Loading Outlook calendar items...", EventType.Information);
            var calendarItems = this.GetOutlookItems(this._outlookCalendarFolderToSyncID);
            Logger.Log("Sorting Outlook calendar items", EventType.Debug);
            calendarItems.Sort("Start");
            //calendarItems.IncludeRecurrences = true;
            //var query = "NOT(\"urn:schemas:calendar:dtstart\" IS NULL)";
            var query = String.Format("@SQL=NOT(\"urn:schemas:calendar:dtstart\" IS NULL)");
            Logger.Log("Filtering Outlook calendar items", EventType.Debug);
            Outlook.AppointmentItem item = calendarItems.Find(query) as Outlook.AppointmentItem;

            while (item != null)
            {
                // any MeetingItem is actually Appointment with special status. So we get correspondent appointmen
                //if (item is Outlook.MeetingItem)
                //    tmpEvent = ((Outlook.MeetingItem)item).GetAssociatedAppointment(false);
                //else if (item is Outlook.AppointmentItem)
                //    tmpEvent = (Outlook.AppointmentItem)item;
                // it's also possible non-calendar item can be placed in the Calendar folder. We'll just ignore all such items
                //else
                //    continue;
                //if ((tmpEvent.Start >= this._rangeStart) && (tmpEvent.Start < this._rangeEnd))
                //{
                    this.OutlookItems.Add(item);
                //}
                // release appoinment item since it's COM object 
                //Marshal.ReleaseComObject(item);
                item = (Outlook.AppointmentItem)calendarItems.FindNext();
            }
            Marshal.ReleaseComObject(calendarItems);
        }

        private void LoadSettings()
        {
            // if start and end synchronization range can't be read from the registry default values will be used
            var daysBefore = 0;
            var daysAfter = 0;
            try
            {
                var regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Properties.Settings.Default.ApplicationRegistryKey);
                if (regKey == null)
                    throw new Exception();
                this._googleCalendarName = (string)regKey.GetValue(Properties.Settings.Default.CalendarSettings_CalendarName);
                daysBefore = (int)regKey.GetValue(Properties.Settings.Default.CalendarSettings_RangeBefore);
                daysAfter = (int)regKey.GetValue(Properties.Settings.Default.CalendarSettings_RangeAfter);
            }
            // in case no registry key exist or some registry values can't be loaded default values are used
            catch (Exception)
            { }

            // set start and end of the synchronization range
            this._rangeStart = DateTime.Now.AddDays(-daysBefore).Date;
            this._rangeEnd = DateTime.Now.AddDays(daysAfter + 1).Date;
        }

        protected override void UpdateGoogleItem(ItemMatcher pair)
        {

#if DEBUG
            var outlookItem = pair.Outlook;
            var googleItem = pair.Google;
#endif
            try
            {
                if (pair.SyncAction.Action == Action.Create)
                {
                    /// Here new Google event is created and saved into Google calendar
                    /// then it will be updated according to Outlook's original
                    /// If something wrong will happen during updating newly created Google event will be deleted
                    pair.Google = new EventEntry();
                    pair.Google.Service = this._googleConnection;
                    GoogleUtilities.SetOutlookID(pair.Google, OutlookUtilities.GetItemID(pair.Outlook));
                    try
                    {
                        pair.Google = this._googleConnection.Insert(new Uri(this._googleCalendarFeedUrl), pair.Google);

                        foreach (var fieldHandler in this._fieldHandlers)
                        {
                            fieldHandler.Setter.Invoke(this, new object[] { 
                            pair.Google,
                            pair.Outlook,
                            Target.Google});
                        }
                        /// Queue item to backup for update
                        //pair.Google.BatchData = new GDataBatchEntryData(GDataBatchOperationType.update);
                        //this._googleBatchUpdateFeed.Entries.Add(pair.Google);
                        /// The item must be updated before setting exceptions because batch operation order is set by Google itself
                        /// Therefore it can come exception is attempted to be set before original item, which will cause the error
                        /// May be later it will be revorked
                        pair.Google.Update();
                        /// As last point we set newly created Google event's ID to Outlook's item
                        OutlookUtilities.SetGoogleID(pair.Outlook, GoogleUtilities.GetItemID(pair.Google));
                        this.SaveOutlookItem(pair.Outlook);
                        this._syncResult.CreatedItems++;
                    }
                    catch (Exception exc)
                    {
                        pair.Google.BatchData = new GDataBatchEntryData(GDataBatchOperationType.delete);
                        this._googleBatchUpdateFeed.Entries.Add(pair.Google);
                    }
                }
                else if (pair.SyncAction.Action == Action.Delete)
                {
                    //pair.Google.Delete();
                    pair.Google.BatchData = new GDataBatchEntryData(GDataBatchOperationType.delete);
                    this._googleBatchUpdateFeed.Entries.Add(pair.Google);
                    this._syncResult.DeletedItems++;
                }
                else if (pair.SyncAction.Action == Action.Update)
                {

                    /// Get list of field setters for fields, which differ
                    var fieldSetters =
                        from fieldHandler in this._fieldHandlers
                        where !(bool)fieldHandler.Comparer.Invoke(this, new object[] { pair.Google, pair.Outlook })
                        select fieldHandler.Setter;
                    foreach (var fieldSetter in fieldSetters)
                    {
                        fieldSetter.Invoke(this, new object[] {
                        pair.Google,
                        pair.Outlook,
                        Target.Google});
                    }
                    //pair.Google.Update();
                    //pair.Google.BatchData = new GDataBatchEntryData(GDataBatchOperationType.update);
                    //this._googleBatchUpdateFeed.Entries.Add(pair.Google);
                    /// The item must be updated before setting exceptions because batch operation order is set by Google itself
                    /// Therefore it can come exception is attempted to be set before original item, which will cause the error
                    /// May be later it will be revorked
                    pair.Google.Update();
                    this._syncResult.UpdatedItems++;
                }
            }
            catch (Exception exc)
            {
                //var buffer = new byte[2097152];
                //var stream = new MemoryStream(buffer, true);
                //var writer = new XmlTextWriter(stream, Encoding.Unicode);
                //pair.Google.SaveToXml(writer);
                //var size = stream.Position;
                //writer.Close();
                //stream = new MemoryStream(buffer);
                //var reader = new StreamReader(stream, Encoding.Unicode, false, (int)size);
                //var atom = reader.ReadToEnd();
                //reader.Close();
                //stream.Close();
                Logger.Log(
                    String.Format("{0} {1} {2}. {3}\r\n{4}",
                        Properties.Resources.Error_ItemSynchronizationFailure,
                        pair.SyncAction,
                        ((Outlook.AppointmentItem)pair.Outlook).Subject,
                        "", //atom,
                        ErrorHandler.BuildExceptionDescription(exc)),
                    EventType.Error);
                this._syncResult.ErrorItems++;            
            }
        }

        protected override void UpdateOutlookItem(ItemMatcher pair)
        {
#if DEBUG
            var outlookItem = pair.Outlook;
            var googleItem = pair.Google;
#endif
            try
            {
                if (pair.SyncAction.Action == Action.Create)
                {
                    var outlookEvent = (Outlook.AppointmentItem)OutlookConnection.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                    foreach (var fieldHandler in this._fieldHandlers)
                    {
                        fieldHandler.Setter.Invoke(this, new object[] { 
                        pair.Google,
                        outlookEvent,
                        Target.Outlook
                    });
                    }
                    OutlookUtilities.SetGoogleID(outlookEvent, GoogleUtilities.GetItemID(pair.Google));
                    this.SaveOutlookItem(outlookEvent);
                    if (this._outlookCalendarFolderToSyncID != OutlookConnection.Namespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).EntryID)
                        OutlookUtilities.TryDo(() => outlookEvent.Move(OutlookConnection.Namespace.GetFolderFromID(this._outlookCalendarFolderToSyncID)));
                    GoogleUtilities.SetOutlookID(pair.Google, OutlookUtilities.GetItemID(outlookEvent));
                    var res = pair.Google.Update();
                    Marshal.ReleaseComObject(outlookEvent);
                    this._syncResult.CreatedItems++;
                }
                else if (pair.SyncAction.Action == Action.Delete)
                {
                    ((Outlook.AppointmentItem)pair.Outlook).Delete();
                    Marshal.ReleaseComObject(pair.Outlook);
                    this._syncResult.DeletedItems++;
                }
                else if (pair.SyncAction.Action == Action.Update)
                {
                    /// Get list of field setters for fields, which differ
                    var fieldSetters =
                        from fieldHandler in this._fieldHandlers
                        where !(bool)fieldHandler.Comparer.Invoke(this, new object[] { pair.Google, pair.Outlook })
                        select fieldHandler.Setter;
                    foreach (var fieldSetter in fieldSetters)
                    {
                        fieldSetter.Invoke(this, new object[] {
                        pair.Google,
                        pair.Outlook,
                        Target.Outlook
                    });
                    }
                    this.SaveOutlookItem(pair.Outlook);
                    Marshal.ReleaseComObject(pair.Outlook);
                    this._syncResult.UpdatedItems++;
                }
            }
            catch (Exception exc)
            {
                Logger.Log(String.Format("{0} {1} '{2}'. {3}", Properties.Resources.Error_ItemSynchronizationFailure, pair.SyncAction, pair.Google.Title.Text, ErrorHandler.BuildExceptionDescription(exc)), EventType.Error);
                this._syncResult.ErrorItems++;
            }
        }

        protected override void SaveOutlookItem(object outlookItem)
        {
            ((Outlook.AppointmentItem)outlookItem).Save();
        }
    }
}

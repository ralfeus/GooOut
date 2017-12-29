using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Calendar.v3;
using Google.Apis.Requests;
using System.Threading;
using System.Net.Http;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;

namespace R.GoogleOutlookSync
{
    internal partial class CalendarSynchronizer : ItemTypeSynchronizer
    {
        /// <summary>
        /// Wrapper around common Google service for CalendarService
        /// </summary>
        private CalendarService CalendarService { get => (CalendarService)this._googleService; }
        /// <summary>
        /// Represents calendar to synchronize with
        /// </summary>
        Calendar _googleCalendar;
        private IList<EventReminder> _defaultReminders;
        string _googleCalendarName;

        /// <summary>
        /// Represents folder in Outlook mailbox containing calendar items for synchronization
        /// </summary>
        string _outlookCalendarFolderToSyncID;
        /// <summary>
        /// Contains exceptions from recurrent events
        /// </summary>
        Dictionary<string, List<Event>> _googleExceptions;
        new IEnumerable<FieldHandlers> _fieldHandlers;
        /// <summary>
        /// Contains 0:00 of next date after defined interval. Next date is not inclusive
        /// </summary>
        DateTime _rangeEnd;
        DateTime _rangeStart;

        internal CalendarSynchronizer(
            UserCredential googleCredential,
            //Outlook.NameSpace outlookNamespace,
            string calendarFolderToSyncID)
        {
            this._googleService = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = googleCredential
            });

            this._outlookCalendarFolderToSyncID = calendarFolderToSyncID;
            this._googleExceptions = new Dictionary<string, List<Event>>();
            //this._googleBatchRequest = new BatchRequest(this.CalendarService);
            
            this.LoadSettings();
            var calendars = this.CalendarService.CalendarList.List().Execute().Items;
            // first load calendars synchronization settings to have Google calendar name to synchronize
            // in case no settings are loaded first found calendar will be used
            foreach (var cal in calendars)
            {
                Logger.Log("Found calendar: " + cal.Summary, EventType.Information);
            }
            var calendar = calendars.First(
                cal => (cal.Summary == this._googleCalendarName) || String.IsNullOrEmpty(this._googleCalendarName));
            this._googleCalendar = this.CalendarService.Calendars.Get(calendar.Id).Execute();
            this._defaultReminders = calendar.DefaultReminders;
            /// Prepare list of all available field handlers
            this.InitFieldHandlers();
        }

        /// <summary>
        /// Get last modification time for Google event. Check also time of exceptions modification 
        /// </summary>
        /// <param name="googleItem"></param>
        /// <returns></returns>
        protected override DateTime GetLastModificationTime(Event googleItem)
        {
            if (this._googleExceptions.ContainsKey(((Event)googleItem).Id))
                return
                    new DateTime(Math.Max(
                        GoogleUtilities.GetLastModificationTime(googleItem).Ticks,
                        this._googleExceptions[((Event)googleItem).Id].Max(exception => exception.Updated.Value.Ticks)));
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

        protected override ComparisonResult Compare(Event googleItem, object item)
        {
            /// If IDs are same items are same. Then we can check whether they are identical
            /// If IDs are different items are different
            /// Aggregate will invoke each method in _fieldComparers list. Boolean AND will ensure all comparers return true
            //try
            //{
            Outlook.AppointmentItem appointment = (Outlook.AppointmentItem)item;
                if (this.CompareIDs((Event)googleItem, appointment))
                {
                    var fieldsAreSame =
                        this._fieldHandlers.Aggregate(
                            true,
                            (res, handler) => {
                                Logger.Log(string.Format("Comparing '{0}' of {1} at {2}", handler.Comparer.Name, appointment.Subject, appointment.Start), EventType.Debug);
                                var result = res && (bool)handler.Comparer.Invoke(this, new object[] { googleItem, appointment });
                                Logger.Log(string.Format("\t{0}", result), EventType.Debug);
                                return result;
                            }
                        );
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

        protected override bool IsItemValid(Event googleItem)
        {
            /// We try to create EventSchedule of item. If due to any reason (Exception) it's impossible
            /// we treat item as invalid
            try
            {
                new EventSchedule((Event)googleItem);
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
                    var query = this.CalendarService.Events.List(this._googleCalendar.Id);
                    query.TimeMin = this._rangeStart;
                    query.TimeMax = this._rangeEnd;
                    query.ShowDeleted = true;
                    var events = query.Execute().Items;
                    this.GoogleItems = events.Where(e => e.Status != "cancelled").ToList();
                    foreach (var e in events.Where(e => !string.IsNullOrEmpty(e.RecurringEventId))) {
                        if (!this._googleExceptions.ContainsKey(e.RecurringEventId)) {
                            this._googleExceptions.Add(e.RecurringEventId, new List<Event>());
                        }
                        this._googleExceptions[e.RecurringEventId].Add(e);
                    }
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
                Outlook.AppointmentItem tmpEvent;
                // any MeetingItem is actually Appointment with special status. So we get correspondent appointmen
                if (item is Outlook.MeetingItem)
                    tmpEvent = ((Outlook.MeetingItem)item).GetAssociatedAppointment(false);
                else if (item is Outlook.AppointmentItem)
                    tmpEvent = item;
                // it's also possible non-calendar item can be placed in the Calendar folder. We'll just ignore all such items
                else
                    continue;
                if ((tmpEvent.Start >= this._rangeStart) && (tmpEvent.Start < this._rangeEnd))
                {
                    this.OutlookItems.Add(item);
                }
                //// release appoinment item since it's COM object 
                //Marshal.ReleaseComObject(item);
                //Marshal.ReleaseComObject(tmpEvent);
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
            Logger.Log("UpdateGoogleItem", EventType.Debug);
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
                    pair.Google = new Event()
                    {
                        End = new EventDateTime(),
                        ExtendedProperties = new Event.ExtendedPropertiesData()
                        {
                            Shared = new Dictionary<string, string>()
                        },
                        Location = "",
                        Reminders = new Event.RemindersData()
                        {
                            //Overrides = new 
                            UseDefault = false
                        },
                        Start = new EventDateTime()
                    };
                    //pair.Google.Service = this.CalendarService;
                    GoogleUtilities.SetOutlookID(pair.Google, OutlookUtilities.GetItemID(pair.Outlook));
                    try
                    {
#if DEBUG
                        Logger.Log("Setting event fields", EventType.Debug);
#endif
                        foreach (var fieldHandler in this._fieldHandlers)
                        {
                            fieldHandler.Setter.Invoke(this, new object[] { 
                            pair.Google,
                            pair.Outlook,
                            Target.Google});
                        }
#if DEBUG
                        Logger.Log("Trying to create Google event", EventType.Debug);
#endif
                        pair.Google = this.CalendarService.Events.Insert(pair.Google, this._googleCalendar.Id).Execute();
#if DEBUG
                        Logger.Log(string.Format("New Google event with eventId {0} is created", pair.Google.Id), EventType.Debug);
                        Logger.Log(Newtonsoft.Json.JsonConvert.SerializeObject(pair.Google), EventType.Debug);
#endif
                        /// When item is created it has an ID. 
                        /// If it's recurrent all recurrence instances are created. 
                        /// This gives a possibility to set recurrence exceptions.
                        /// Earlier exceptions can't be set
                        this.SetRecurrenceExceptions((Outlook.AppointmentItem)pair.Outlook, pair.Google);
                        /// As last point we set newly created Google event's ID to Outlook's item
                        OutlookUtilities.SetGoogleID(pair.Outlook, GoogleUtilities.GetItemID(pair.Google));
                        this.SaveOutlookItem(pair.Outlook);
                        this._syncResult.CreatedItems++;
                    }
                    catch (Exception exc )
                    {
                        ErrorHandler.Handle(exc);
                        this._googleBatchRequest.Queue<Event>(this.CalendarService.Events.Delete(
                            this._googleCalendar.Id, pair.Google.Id), BatchCallback);
                    }
                }
                else if (pair.SyncAction.Action == Action.Delete)
                {
                    //pair.Google.Delete();
                    this._googleBatchRequest.Queue<Event>(this.CalendarService.Events.Delete(
                        this._googleCalendar.Id, pair.Google.Id), BatchCallback);
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
                    pair.Google = this.CalendarService.Events.Update(pair.Google, this._googleCalendar.Id, pair.Google.Id).Execute();
                    /// When item is updated and it's became recurrent all recurrence instances are created. 
                    /// This gives a possibility to set recurrence exceptions.
                    /// Earlier exceptions can't be set
                    this.SetRecurrenceExceptions((Outlook.AppointmentItem)pair.Outlook, pair.Google);

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

        private void BatchCallback(Event content, RequestError error, int index, HttpResponseMessage message)
        {
            if (!message.IsSuccessStatusCode)
            {
                this._batchResult.Add(error);
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
                    var res = this.CalendarService.Events.Update(pair.Google, this._googleCalendar.Id, pair.Google.Id);
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
                Logger.Log(String.Format("{0} {1} '{2}'. {3}", Properties.Resources.Error_ItemSynchronizationFailure, pair.SyncAction, pair.Google.Summary, ErrorHandler.BuildExceptionDescription(exc)), EventType.Error);
                this._syncResult.ErrorItems++;
            }
        }

        protected override void SaveOutlookItem(object outlookItem)
        {
            ((Outlook.AppointmentItem)outlookItem).Save();
        }

        public override SyncResult Sync()
        {
            Logger.Log("Initializing " + this.GetType().Name, EventType.Debug);
            this.Init();
            LoadGoogleItems();
            LoadOutlookItems();
            Logger.Log(String.Format("Got {0} Google items and {1} Outlook items", this.GoogleItems.Count, this.OutlookItems.Count), EventType.Information);
            this._outlookGoogleIDsCache = new Dictionary<object, string>(OutlookItems.Count);
            Logger.Log("Comparing items", EventType.Debug);
            this.CombineItems();
            /// Because error items were marked as identical we should subtract errorous items from identical ones
            this._syncResult.IdenticalItems -= this._syncResult.ErrorItems;

            Logger.Log(String.Format("There are {0} items to update", this._itemsPairs.Count), EventType.Information);
            Logger.Log("Updating items", EventType.Debug);
            /// Update items
            foreach (var pair in this._itemsPairs)
            {
                var go = (Event)pair.Google;
                var startTime = pair.Google != null ? (pair.Google.Start.Date ?? pair.Google.Start.DateTime.Value.ToString("yyyy-MM-dd")) : "";
                Logger.Log(String.Format(
                    "Running action '{0}' on item '{1}' starting at {2}. Target: {3}",
                    pair.SyncAction.Action,
                    pair.Google == null ? ((Outlook.AppointmentItem)pair.Outlook).Subject : pair.Google.Summary,
                    pair.Google == null ? ((Outlook.AppointmentItem)pair.Outlook).Start.ToString() : startTime,
                    pair.SyncAction.Target), EventType.Debug);
                try
                {
                    if (pair.SyncAction.Target == Target.Google)
                        UpdateGoogleItem(pair);
                    else
                        UpdateOutlookItem(pair);
                }
                catch (COMException exc)
                {
                    if (exc.ErrorCode == unchecked((int)0x80010105))
                    {
                        Logger.Log(exc.Message, EventType.Error);
                        this._syncResult.ErrorItems++;
                        continue;
                    }
                    else
                        throw exc;
                }
            }
            /// Run batch update on all Google items requiring update
            this._googleBatchRequest.ExecuteAsync(CancellationToken.None).Wait();
            /// Extract errors from batch result and log it
            var errors =
                from entry in this._batchResult
                where
                    entry.Code != 200 &&
                    entry.Code != 201
                select entry.Message;
            foreach (var error in errors)
                Logger.Log(error, EventType.Error);

            return this._syncResult;
        }

        internal override void Unpair()
        {
            this.Init();
            this.LoadGoogleItems();
            var batch = new BatchRequest(this.CalendarService);
            foreach (var googleItem in this.GoogleItems)
            {
                GoogleUtilities.RemoveOutlookID(googleItem);
                batch.Queue<Event>(this.CalendarService.Events.Update(googleItem, this._googleCalendar.Id, googleItem.Id), BatchCallback);
            }
            batch.ExecuteAsync(CancellationToken.None).Wait();

            this.LoadOutlookItems();
            foreach (var outlookItem in this.OutlookItems)
            {
                OutlookUtilities.RemoveGoogleID(outlookItem);
                this.SaveOutlookItem(outlookItem);
            }
        }
    }
}

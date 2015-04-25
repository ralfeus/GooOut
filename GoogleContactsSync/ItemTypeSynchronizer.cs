using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using GoogleConn = Google.GData.Client;
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;

namespace R.GoogleOutlookSync
{
    internal abstract class ItemTypeSynchronizer
    {
        private Dictionary<object, string> _outlookGoogleIDsCache;

        protected IEnumerable<FieldHandlers> _fieldHandlers;
        protected GoogleConn.AtomFeed _googleBatchUpdateFeed;
        protected GoogleConn.Service _googleConnection;
        protected List<ItemMatcher> _itemsPairs;
        protected List<GoogleConn.AtomEntry> GoogleItems { get; set; }
        protected List<object> OutlookItems { get; set; }
        protected SyncResult _syncResult = new SyncResult();

        public SyncResult Sync()
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
                var go = (Google.GData.Calendar.EventEntry)pair.Google;
                Google.GData.Extensions.ExtensionCollection<Google.GData.Extensions.When> times = null;
                if (go != null)
                    times = go.Times;
                Logger.Log(String.Format(
                    "Running action '{0}' on item '{1}' starting at {2}. Target: {3}",
                    pair.SyncAction.Action,
                    pair.Google == null ? ((Outlook.AppointmentItem)pair.Outlook).Subject : pair.Google.Title.Text,
                    pair.Google == null ? ((Outlook.AppointmentItem)pair.Outlook).Start.ToString() : (times.Count > 0 ? times[0].StartTime.ToString() : Regex.Match("DTSTART=(.*)\r", go.Recurrence.Value).Groups[1].Value),
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
            var result = this._googleConnection.Batch(this._googleBatchUpdateFeed, new Uri(this._googleBatchUpdateFeed.Batch));
            /// Extract errors from batch result and log it
            var errors =
                from entry in result.Entries
                where
                    entry.BatchData.Status.Code != 200 &&
                    entry.BatchData.Status.Code != 201
                select entry.BatchData.Status.Reason;
            foreach (var error in errors)
                Logger.Log(error, EventType.Error);

            return this._syncResult;
        }

        private void Init()
        {
            this.GoogleItems = new List<GoogleConn.AtomEntry>();
            this.OutlookItems = new List<object>();
            //this.InitFieldHandlers();
        }

        protected void InitFieldHandlers()
        {
            /// 1. Get methods, which current instance have. Include non-public methods too 
            ///   (handlers are private since they are not used anywhere else)
            var methods = this.GetType().GetMethods(BindingFlags.Instance | BindingFlags.NonPublic);
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
                select new FieldHandlers(
                    (Func<GoogleConn.AtomEntry, object, bool>)Delegate.CreateDelegate(typeof(Func<GoogleConn.AtomEntry, object, bool>), this, comparerMethod.Name),
                    (Action<GoogleConn.AtomEntry, object, Target>)Delegate.CreateDelegate(typeof(Action<GoogleConn.AtomEntry, object, Target>), this, setterMethod.Name));
            ComparerDelegate test;
            foreach (var method in comparerMethods)
                test = (ComparerDelegate)Delegate.CreateDelegate(typeof(ComparerDelegate), this, method.Name);
        }

        /// <summary>
        /// Creates a combined list of Google and Outlook items, which differ in some way
        /// For same but changed items it's one record.
        /// For each item having no pair it's one record per item
        /// </summary>
        protected virtual void CombineItems()
        {
            this._itemsPairs = new List<ItemMatcher>(this.GoogleItems.Count + this.OutlookItems.Count);
            bool noFurtherCheckNeeded;
            ComparisonResult itemsComparisonResult = 0;
            int counter = 0; int outlookItemsCount = this.OutlookItems.Count;
            foreach (var outlookItem in new List<object>(this.OutlookItems))
            {
                Logger.Log(String.Format("Checking Outlook item {0} of {1}", ++counter, outlookItemsCount), EventType.Debug);
                noFurtherCheckNeeded = false;

                foreach (var googleItem in new List<GoogleConn.AtomEntry>(this.GoogleItems))
                {
                    try
                    {
                        itemsComparisonResult = this.Compare(googleItem, outlookItem);
                    }
                    catch (UnsynchronizableItemException exc)
                    {
                        Logger.Log(exc.Message + (exc.ItemType == Target.Google ? googleItem.Title.Text : OutlookUtilities.GetItemSubject(outlookItem)), EventType.Warning);
                        /// if already paired items is impossible to synchronized due to some reason
                        /// we treat them as identical
                        this._syncResult.ErrorItems++;
                        itemsComparisonResult = ComparisonResult.Identical;
                        Logger.Log(String.Format("While matching Google item '{0}' and Outlook item '{1}' the error has happened.", googleItem.Title.Text, OutlookUtilities.GetItemSubject(outlookItem)), EventType.Debug);
                    }
                    /// If items are same we just remove from both lists and forget about them
                    if (itemsComparisonResult == ComparisonResult.Identical)
                    {
                        Logger.Log(String.Format("Google item '{0}' and Outlook item '{1}' are identical.", googleItem.Title.Text, OutlookUtilities.GetItemSubject(outlookItem)), EventType.Debug);
                        this.GoogleItems.Remove(googleItem);
                        this.OutlookItems.Remove(outlookItem);
                        noFurtherCheckNeeded = true;
                        this._syncResult.IdenticalItems++;
                        break;
                    }
                    /// If items are same but changed we will store in pairs list for further synchronization
                    /// However if synchronization setting is one-way synchronization then all pairs, which have to
                    /// change source item will be ignored
                    else if (itemsComparisonResult == ComparisonResult.SameButChanged)
                    {
                        Logger.Log(String.Format("Google item '{0}' and Outlook item '{1}' are same but changed.", googleItem.Title.Text, OutlookUtilities.GetItemSubject(outlookItem)), EventType.Debug);
                        var itemMatcher = new ItemMatcher(googleItem, outlookItem);
                        if (Properties.Settings.Default.SynchronizationOption == SyncOption.GoogleToOutlookOnly)
                            itemMatcher.SyncAction.Target = Target.Outlook;
                        else if (Properties.Settings.Default.SynchronizationOption == SyncOption.OutlookToGoogleOnly)
                            itemMatcher.SyncAction.Target = Target.Google;
                        else
                            itemMatcher.SyncAction.Target = this.WhoLoses(googleItem, outlookItem);

                        itemMatcher.SyncAction.Action = Action.Update;
                        this._itemsPairs.Add(itemMatcher);

                        this.GoogleItems.Remove(googleItem);
                        this.OutlookItems.Remove(outlookItem);
                        noFurtherCheckNeeded = true;
                        break;
                    }
                }
                /// If no pair was found for the Outlook item and item is valid (IsItemValid() check) it will be added to the list alone for further synchronization
                /// If synchronization setting is OutlookToGoogleOnly only items with Target == Google will be added. No Outlook targeted items will be added
                if (!noFurtherCheckNeeded && this.IsItemValid(outlookItem))
                {
                    var itemMatcher = new ItemMatcher(null, outlookItem);
                    itemMatcher.SyncAction = this.GetNonPairedItemAction(outlookItem);
                    
                    Logger.Log(String.Format("Outlook item '{0}' has no pair.", OutlookUtilities.GetItemSubject(outlookItem)), EventType.Debug);
                    Logger.Log(String.Format("Action {0} will be performed on {1}", itemMatcher.SyncAction.Action, itemMatcher.SyncAction.Target), EventType.Debug);

                    if (((itemMatcher.SyncAction.Target == Target.Google) && (Properties.Settings.Default.SynchronizationOption != SyncOption.GoogleToOutlookOnly)) ||
                        ((itemMatcher.SyncAction.Target == Target.Outlook) && (Properties.Settings.Default.SynchronizationOption != SyncOption.OutlookToGoogleOnly)))
                        this._itemsPairs.Add(itemMatcher);
                }

                //break;
            }
            /// All remaining valid (IsItemValid() check) Outlook items without Google pair will be added to the list alone for further synchronization
            /// If synchronization setting is GoogleToOutlookOnly only items with Target == Outlook will be added. No Google targeted items will be added
            Logger.Log(String.Format("{0} Outlook items remain unmatched", this.OutlookItems.Count), EventType.Debug);
            foreach (var googleItem in this.GoogleItems)
            {
                if (this.IsItemValid(googleItem))
                {
                    var itemMatcher = new ItemMatcher(googleItem, null);
                    itemMatcher.SyncAction = this.GetNonPairedItemAction(googleItem);

                    Logger.Log(String.Format("Google item '{0}' has no pair.", googleItem.Title.Text), EventType.Debug);
                    Logger.Log(String.Format("Action {0} will be performed on {1}", itemMatcher.SyncAction.Action, itemMatcher.SyncAction.Target), EventType.Debug);

                    if (((itemMatcher.SyncAction.Target == Target.Google) && (Properties.Settings.Default.SynchronizationOption != SyncOption.GoogleToOutlookOnly)) ||
                        ((itemMatcher.SyncAction.Target == Target.Outlook) && (Properties.Settings.Default.SynchronizationOption != SyncOption.OutlookToGoogleOnly)))
                        this._itemsPairs.Add(itemMatcher);
                }
            }

            /// Clean up source Google and Outlook items lists
            this.GoogleItems.Clear();
            this.OutlookItems.Clear();
        }

        /// <summary>
        /// Compares IDs of Google and Outlook items. If IDs are same items are same (regardless synchronized they are or not)
        /// Checks value of item ID with value of corresponding extended property of the peer item
        /// </summary>
        /// <param name="googleItem">Google item</param>
        /// <param name="outlookItem">Outlook item</param>
        /// <returns>true if IDs are same, false if IDs are different</returns>
        protected bool CompareIDs(GoogleConn.AtomEntry googleItem, object outlookItem)
        {
            var outlookItemID = OutlookUtilities.GetItemID(outlookItem);
            var googleItemID = GoogleUtilities.GetItemID(googleItem);
            if (!this._outlookGoogleIDsCache.ContainsKey(outlookItem))
                this._outlookGoogleIDsCache.Add(outlookItem, OutlookUtilities.GetGoogleID(outlookItem));
            var outlookItemGoogleID = this._outlookGoogleIDsCache[outlookItem];
            var googleItemOutlookID = GoogleUtilities.GetOutlookID(googleItem);
            return (googleItemID == outlookItemGoogleID) || (outlookItemID == googleItemOutlookID);
        }

        /// <summary>
        /// Compares Google and Outlook items and check whether they are completely identical
        /// </summary>
        /// <param name="googleItem">Google item</param>
        /// <param name="outlookItem">Outlook item</param>
        /// <returns>true if items are identical, false if items differ</returns>
        protected virtual ComparisonResult Compare(GoogleConn.AtomEntry googleItem, object outlookItem)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Defines synchronization action for Google item, which has no Outlook pair 
        /// If it's new item the Outlook item will be created
        /// If it's old item and Outlook item is already deleted the Google item will be deleted as well
        /// </summary>
        /// <param name="outlookItem">Google item</param>
        /// <returns>Action and target this action should be performed on</returns>
        private SyncAction GetNonPairedItemAction(GoogleConn.AtomEntry googleItem)
        {
            if (String.IsNullOrEmpty(GoogleUtilities.GetOutlookID(googleItem)))
                return new SyncAction(Target.Outlook, Action.Create);
            else
                return new SyncAction(Target.Google, Action.Delete);
        }

        /// <summary>
        /// Defines synchronization action for Outlook item, which has no Google pair 
        /// If it's new item the Google item will be created
        /// If it's old item and Google item is already deleted the Outlook item will be deleted as well
        /// </summary>
        /// <param name="outlookItem">Outlook item</param>
        /// <returns>Action and target this action should be performed on</returns>
        private SyncAction GetNonPairedItemAction(object outlookItem)
        {
            if (String.IsNullOrEmpty(OutlookUtilities.GetGoogleID(outlookItem)))
                return new SyncAction(Target.Google, Action.Create);
            else
                return new SyncAction(Target.Outlook, Action.Delete);
        }

        protected abstract bool IsItemValid(GoogleConn.AtomEntry googleItem);
        protected abstract bool IsItemValid(object outlookItem);
        protected abstract void LoadGoogleItems();
        protected abstract void LoadOutlookItems();
        protected abstract void UpdateGoogleItem(ItemMatcher pair);
        protected abstract void UpdateOutlookItem(ItemMatcher pair);

        /// <summary>
        /// Defines what item - Google or Outlook should be updated
        /// First check synchronization settings.
        /// If synchronization settings allow both sides to be updated checks, at which side item is fresher
        /// </summary>
        /// <param name="googleItem">Google item</param>
        /// <param name="outlookItem">Outlook item</param>
        /// <returns></returns>
        private Target WhoLoses(GoogleConn.AtomEntry googleItem, object outlookItem)
        {
            /// If merge master is defined explicitly by option we don't care on other possibilities
            /// Right now no other merge options but default one is available
            //if (Properties.Settings.Default.SynchronizationOption == SyncOption.MergeGoogleWins)
            //    return Targets.Outlook;
            //else if (Properties.Settings.Default.SynchronizationOption == SyncOption.MergeOutlookWins)
            //    return Targets.Google;
            //else if (Properties.Settings.Default.SynchronizationOption == SyncOption.MergePrompt)
            //    return this.GetUserTargetDecision();
            /// In case no weird synchronization option is set (default is Merge)
            /// Winner (and loser) will be defined by last modification time.
            /// Freshest item wins
            var googleLastModificationTime = this.GetLastModificationTime(googleItem);
            var outlookLastModificationTime = this.GetLastModificationTime(outlookItem);
            if (googleLastModificationTime < outlookLastModificationTime)
                return Target.Google;
            else if (outlookLastModificationTime < googleLastModificationTime)
                return Target.Outlook;
            else
                throw new CannotDefineSynchronizationTargetException();
        }

        protected virtual DateTime GetLastModificationTime(GoogleConn.AtomEntry googleItem)
        {
            throw new NotImplementedException();
        }

        protected virtual DateTime GetLastModificationTime(object outlookItem)
        {
            return OutlookUtilities.GetLastModificationTime(outlookItem);
        }

        protected Outlook.Items GetOutlookItems(string outlookFolderID)
        {
            Logger.Log("Getting Outlook folder object", EventType.Debug);
            Outlook.MAPIFolder mapiFolder = OutlookConnection.Namespace.GetFolderFromID(outlookFolderID);
            try
            {
                Logger.Log("Getting Outlook items collection", EventType.Debug);
                return mapiFolder.Items;
            }
            finally
            {
                if (mapiFolder != null)
                    Marshal.ReleaseComObject(mapiFolder);
                mapiFolder = null;
            }
        }

        private Target GetUserTargetDecision()
        {
            throw new NotImplementedException();
        }

        internal void Unpair()
        {

            this.Init();
            this.LoadGoogleItems();
            foreach (var googleItem in this.GoogleItems)
            {
                GoogleUtilities.RemoveOutlookID(googleItem);
                googleItem.BatchData = new GoogleConn.GDataBatchEntryData(GoogleConn.GDataBatchOperationType.update);
                this._googleBatchUpdateFeed.Entries.Add(googleItem);
            }
            var result = this._googleConnection.Batch(this._googleBatchUpdateFeed, new Uri(this._googleBatchUpdateFeed.Batch));

            this.LoadOutlookItems();
            foreach (var outlookItem in this.OutlookItems)
            {
                OutlookUtilities.RemoveGoogleID(outlookItem);
                this.SaveOutlookItem(outlookItem);
            }
        }

        protected abstract void SaveOutlookItem(object outlookItem);
    }
}
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using System.Threading;
using Google.Apis.Util.Store;
//using AddinExpress.Outlook;

namespace R.GoogleOutlookSync
{
    internal class Synchronizer
	{
		public const int OutlookUserPropertyMaxLength = 32;
		public const string OutlookUserPropertyTemplate = "g/con/{0}/";
        internal const string myContactsGroup = "System Group: My Contacts";
		private static object _syncRoot = new object();
        private List<ItemTypeSynchronizer> _syncBatch = new List<ItemTypeSynchronizer>(1);
        //private SecurityManager _sm = new SecurityManager();

        internal SyncResult Result { get; private set; }

		private int _totalCount;
        public int TotalCount { get { return _totalCount; } }

        private int _skippedCount;
        public int SkippedCount
        {
            set { _skippedCount = value; }
            get { return _skippedCount; }
        }

        private int _skippedCountNotMatches;
        public int SkippedCountNotMatches
        {
            set { _skippedCountNotMatches = value; }
            get { return _skippedCountNotMatches; }
        }

		public delegate void DuplicatesFoundHandler(string title, string message);
		public delegate void ErrorNotificationHandler(string title, Exception ex, EventType eventType);
		public event DuplicatesFoundHandler DuplicatesFound;
		public event ErrorNotificationHandler ErrorEncountered;

        private string _propertyPrefix;
        public string OutlookPropertyPrefix
        {
            get { return _propertyPrefix; }
        }

        public string OutlookPropertyNameId
        {
            get { return _propertyPrefix + "id"; }
        }

        /*public string OutlookPropertyNameUpdated
        {
        	get { return _propertyPrefix + "up"; }
        }*/

        public string OutlookPropertyNameSynced
        {
            get { return _propertyPrefix + "up"; }
        }

        private SyncOption _syncOption = SyncOption.Merge;
        public SyncOption SyncOption
        {
            get { return _syncOption; }
            set { _syncOption = value; }
        }

        private string _syncProfile = "";
        public string SyncProfile
        {
            get { return _syncProfile; }
            set { _syncProfile = value; }
        }

        public string CalendarFolderToSyncID { get; set; }

        private bool _syncContacts;
        private UserCredential _googleCredential;

        /// <summary>
        /// If true sync also contacts
        /// </summary>
        public bool SyncContacts
        {
            get { return _syncContacts; }
            set { _syncContacts = value; }
        }
       
        /// <summary>
        /// If true sync Calendar
        /// </summary>
        public bool SyncCalendar { get; set; }

        /// <summary>
        /// If true sync Tasks
        /// </summary>
        public bool SyncTasks { get; set; }

		public void LoginToGoogle(string googleLogonAccount)
		{
			Logger.Log("Connecting to Google...", EventType.Information);
            var authorization = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets()
                {
                    ClientId = Properties.Settings.Default.GoogleAuth_ClientID,
                    ClientSecret = Properties.Settings.Default.GoogleAuth_ClientSecret
                }, 
                new[] { CalendarService.Scope.Calendar }, googleLogonAccount, CancellationToken.None,
                new FileDataStore("GooOut"));
            this._googleCredential = authorization.Result;

            int maxUserIdLength = Synchronizer.OutlookUserPropertyMaxLength - (Synchronizer.OutlookUserPropertyTemplate.Length - 3 + 2);//-3 = to remove {0}, +2 = to add length for "id" or "up"
            string userId = this._googleCredential.UserId;
			if (userId.Length > maxUserIdLength)
				userId = userId.GetHashCode().ToString("X"); //if a user id would overflow UserProperty name, then use that user id hash code as id.
            //Remove characters not allowed for Outlook user property names: []_#
            userId.Replace("#", "").Replace("[", "").Replace("]", "").Replace("_", "");

			_propertyPrefix = string.Format(Synchronizer.OutlookUserPropertyTemplate, userId);
		}

        //public void LoginToOutlook()
        //{
        //    Logger.Log("Connecting to Outlook...", EventType.Information);

        //    //try
        //    //{
        //        this._outlookApp = OutlookConnection.Application;
        //        this._outlookNamespace = OutlookConnection.Namespace;
        //        //CreateOutlookInstance();
        //    //}
        //    //catch (System.Runtime.InteropServices.COMException)
        //    //{
        //    //    try
        //    //    {
        //    //        // If outlook was closed/terminated inbetween, we will receive an Exception
        //    //        // System.Runtime.InteropServices.COMException (0x800706BA): The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)
        //    //        // so recreate outlook instance
        //    //        Logger.Log("Cannot connect to Outlook, creating new instance....", EventType.Information);
        //    //        /*_outlookApp = new Outlook.Application();
        //    //        _outlookNamespace = _outlookApp.GetNamespace("mapi");
        //    //        _outlookNamespace.Logon();*/
        //    //        _outlookApp = null;
        //    //        _outlookNamespace = null;
        //    //        CreateOutlookInstance();
                    
        //    //    }
        //    //    catch (Exception ex)
        //    //    {
        //    //        string message = "Cannot connect to Outlook.\r\nPlease restart " + Application.ProductName + " and try again. If error persists, please inform developer.";
        //    //        // Error again? We need full stacktrace, display it!
        //    //        throw new Exception(message, ex);
        //    //    }
        //    //}
        //}

        //private void CreateOutlookInstance()
        //{
        //    if (_outlookApp == null || _outlookNamespace == null)
        //    {

        //        //Try to create new Outlook application 3 times, because mostly it fails the first time, if not yet running
        //        for (int i = 0; i < 3; i++)
        //        {
        //            try
        //            {
        //                // First try to get the running application in case Outlook is already started
        //                try
        //                {
        //                    _outlookApp = Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
        //                }
        //                catch (COMException)
        //                {
        //                    // That failed - try to create a new application object, launching Outlook in the background
        //                    _outlookApp = new Outlook.Application();
        //                }
        //                break;  //Exit the for loop, if creating outllok application was successful
        //            }
        //            catch (COMException ex)
        //            {
        //                if (i == 2)
        //                    throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.", ex);
        //                else //wait ten seconds and try again
        //                    System.Threading.Thread.Sleep(1000 * 10);
        //            }
        //        }

        //        if (_outlookApp == null)
        //            throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");


        //        //Try to create new Outlook namespace 3 times, because mostly it fails the first time, if not yet running
        //        for (int i = 0; i < 3; i++)
        //        {
        //            try
        //            {
        //                _outlookNamespace = _outlookApp.GetNamespace("MAPI");
        //                break;  //Exit the for loop, if creating outllok application was successful
        //            }
        //            catch (COMException ex)
        //            {
        //                if (i == 2)
        //                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
        //                else //wait ten seconds and try again
        //                    System.Threading.Thread.Sleep(1000 * 10);
        //            }
        //        }                                   

        //        if (_outlookNamespace == null)
        //            throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");
        //    }

        //    /*
        //    // Get default profile name from registry, as this is not always "Outlook" and would popup a dialog to choose profile
        //    // no matter if default profile is set or not. So try to read the default profile, fallback is still "Outlook"
        //    string profileName = "Outlook";
        //    using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Outlook\SocialConnector", false))
        //    {
        //        if (k != null)
        //            profileName = k.GetValue("PrimaryOscProfile", "Outlook").ToString();
        //    }
        //    _outlookNamespace.Logon(profileName, null, true, false);*/

        //    //Just try to access the outlookNamespace to check, if it is still accessible, throws COMException, if not reachable           
        //    _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);

        //    //this._sm.ConnectTo(this._outlookApp);
        //    //this._sm.DisableOOMWarnings = true;
        //}

        //public void LogoffOutlook()
        //{
        //    try
        //    {
        //        Logger.Log("Disconnecting from Outlook...", EventType.Debug);
        //        //this._sm.DisableOOMWarnings = false;
        //        //this._sm.Disconnect(this._outlookApp);
        //        if (_outlookNamespace != null)
        //        {
        //            _outlookNamespace.Logoff();
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        // if outlook was closed inbetween, we get an System.InvalidCastException or similar exception, that indicates that outlook cannot be acced anymore
        //        // so as outlook is closed anyways, we just ignore the exception here
        //    }
        //    finally
        //    {
        //        if (_outlookNamespace != null)
        //            Marshal.ReleaseComObject(_outlookNamespace);
        //        if (_outlookApp != null)
        //        {
        //            Marshal.ReleaseComObject(_outlookApp);
        //        }
        //        _outlookNamespace = null;
        //        _outlookApp = null;
        //        Logger.Log("Disconnected from Outlook", EventType.Debug);
        //    }
        //}

        public void LogoffGoogle()
        {            
            //_contactsRequest = null;            
        }

		//private void LoadOutlookContacts()
		//{
		//	Logger.Log("Loading Outlook contacts...", EventType.Information);
  //          _outlookContacts = GetOutlookItems(Outlook.OlDefaultFolders.olFolderContacts);			
		//}

        //private void LoadOutlookNotes()
        //{
        //    Logger.Log("Loading Outlook Notes...", EventType.Information);
        //    _outlookNotes = GetOutlookItems(Outlook.OlDefaultFolders.olFolderNotes);
        //}

        private Outlook.Items GetOutlookItems(Outlook.OlDefaultFolders outlookDefaultFolder)
        {
            Outlook.MAPIFolder mapiFolder = OutlookConnection.Namespace.GetDefaultFolder(outlookDefaultFolder);
            try
            {
                return mapiFolder.Items;
            }
            finally
            {
                if (mapiFolder != null)
                    Marshal.ReleaseComObject(mapiFolder);
                mapiFolder = null;
            }
        }
        ///// <summary>
        ///// Moves duplicates from _outlookContacts to _outlookContactDuplicates
        ///// </summary>
        //private void FilterOutlookContactDuplicates()
        //{
        //    _outlookContactDuplicates = new Collection<Outlook.ContactItem>();

        //    if (_outlookContacts.Count < 2)
        //        return;

        //    Outlook.ContactItem main, other;
        //    bool found = true;
        //    int index = 0;

        //    while (found)
        //    {
        //        found = false;

        //        for (int i = index; i <= _outlookContacts.Count - 1; i++)
        //        {
        //            main = _outlookContacts[i] as Outlook.ContactItem;

        //            // only look forward
        //            for (int j = i + 1; j <= _outlookContacts.Count; j++)
        //            {
        //                other = _outlookContacts[j] as Outlook.ContactItem;

        //                if (other.FileAs == main.FileAs &&
        //                    other.Email1Address == main.Email1Address)
        //                {
        //                    _outlookContactDuplicates.Add(other);
        //                    _outlookContacts.Remove(j);
        //                    found = true;
        //                    index = i;
        //                    break;
        //                }
        //            }
        //            if (found)
        //                break;
        //        }
        //    }
        //}

        //private void LoadGoogleContacts()
        //{
        //          string message = "Error Loading Google Contacts. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";
        //	try
        //	{

        //		Logger.Log("Loading Google Contacts...", EventType.Information);

        //              _googleContacts = new Collection<Contact>();

        //		ContactsQuery query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
        //		query.NumberToRetrieve = 256;
        //		query.StartIndex = 0;

        //              //Only load Google Contacts in My Contacts group (to avoid syncing accounts added automatically to "Weitere Kontakte"/"Further Contacts")
        //              Group group = GetGoogleGroupByName(myContactsGroup);
        //              if (group != null)
        //                  query.Group = group.Id;

        //		//query.ShowDeleted = false;
        //		//query.OrderBy = "lastmodified";

        //              Feed<Contact> feed=_contactsRequest.Get<Contact>(query);

        //              while (feed != null)
        //              {
        //                  foreach (Contact a in feed.Entries)
        //                  {
        //                      _googleContacts.Add(a);
        //                  }
        //                  query.StartIndex += query.NumberToRetrieve;
        //                  feed = _contactsRequest.Get<Contact>(feed, FeedRequestType.Next);

        //              }                              

        //	}
        //          catch (System.Net.WebException ex)
        //          {
        //              //Logger.Log(message, EventType.Error);
        //              throw new GDataRequestException(message, ex);
        //          }
        //          catch (System.NullReferenceException ex)
        //          {
        //              //Logger.Log(message, EventType.Error);
        //              throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
        //          }
        //}
        //private void LoadGoogleGroups()
        //{
        //          string message = "Error Loading Google Groups. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";
        //          try
        //          {
        //              Logger.Log("Loading Google Groups...", EventType.Information);
        //              GroupsQuery query = new GroupsQuery(GroupsQuery.CreateGroupsUri("default"));
        //              query.NumberToRetrieve = 256;
        //              query.StartIndex = 0;
        //              //query.ShowDeleted = false;

        //              _googleGroups = new Collection<Group>();

        //              Feed<Group> feed = _contactsRequest.Get<Group>(query);               

        //              while (feed != null)
        //              {
        //                  foreach (Group a in feed.Entries)
        //                  {
        //                      _googleGroups.Add(a);
        //                  }
        //                  query.StartIndex += query.NumberToRetrieve;
        //                  feed = _contactsRequest.Get<Group>(feed, FeedRequestType.Next);

        //              }

        //              ////Only for debugging or reset purpose: Delete all Gougle Groups:
        //              //for (int i = _googleGroups.Count; i > 0;i-- )
        //              //    _googleService.Delete(_googleGroups[i-1]);
        //          }            
        //	catch (System.Net.WebException ex)
        //	{                               				
        //		//Logger.Log(message, EventType.Error);
        //              throw new GDataRequestException(message, ex);
        //	}
        //          catch (System.NullReferenceException ex)
        //          {
        //              //Logger.Log(message, EventType.Error);
        //              throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
        //          }

        //}

        //private void LoadGoogleNotes()
        //{
        //    LoadGoogleNotes(null);
        //}

        //internal Document LoadGoogleNotes(AtomId id)
        //{
        //    string message = "Error Loading Google Notes. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";

        //    Document ret = null;
        //    try
        //    {
        //        if (id == null) // Only log, if not specific Google Notes are searched
        //            Logger.Log("Loading Google Notes...", EventType.Information);

        //        _googleNotes = new Collection<Document>();

        //        //First get the Notes folder or create it, if not yet existing
        //        _googleNotesFolder = null;
        //        DocumentQuery query = new DocumentQuery(_documentsRequest.BaseUri);
        //        query.Categories.Add(new QueryCategory(new AtomCategory("folder")));
        //        query.Title = "Notes";//ToDo: Make the folder configurable in SettingsForm, for now hardcode to "Notes"
        //        Feed<Document> feed = _documentsRequest.Get<Document>(query);

        //        if (feed != null)
        //        {
        //            foreach (Document a in feed.Entries)
        //            {
        //                _googleNotesFolder = a;
        //                break;
        //            }                    
        //        }

        //        if (_googleNotesFolder == null)
        //        {
        //            _googleNotesFolder = new Document();
        //            _googleNotesFolder.Type = Document.DocumentType.Folder;
        //            //_googleNotesFolder.Categories.Add(new AtomCategory("http://schemas.google.com/docs/2007#folder"));
        //            _googleNotesFolder.Title = query.Title;
        //            _googleNotesFolder = SaveGoogleNote(_googleNotesFolder);
        //        }

        //        if (id == null)
        //            query = new DocumentQuery(_googleNotesFolder.DocumentEntry.Content.AbsoluteUri);
        //        else //if newly created
        //            query = new DocumentQuery(_documentsRequest.BaseUri);
        //        query.Categories.Add(new QueryCategory(new AtomCategory("document")));
        //        query.NumberToRetrieve = 256;
        //        query.StartIndex = 0;                

        //        //query.ShowDeleted = false;
        //        //query.OrderBy = "lastmodified";

        //        feed = _documentsRequest.Get<Document>(query);

        //        while (feed != null)
        //        {
        //            foreach (Document a in feed.Entries)
        //            {
        //                _googleNotes.Add(a);
        //                if (id != null && id.Equals(a.DocumentEntry.Id))
        //                    ret = a;
        //            }
        //            query.StartIndex += query.NumberToRetrieve;
        //            feed = _documentsRequest.Get<Document>(feed, FeedRequestType.Next);

        //        }

        //    }
        //    catch (System.Net.WebException ex)
        //    {
        //        //Logger.Log(message, EventType.Error);
        //        throw new GDataRequestException(message, ex);
        //    }
        //    catch (System.NullReferenceException ex)
        //    {
        //        //Logger.Log(message, EventType.Error);
        //        throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
        //    }

        //    return ret;
        //}

        ///// <summary>
        ///// Load the contacts from Google and Outlook
        ///// </summary>
        //public void LoadContacts()
        //{
        //    LoadOutlookContacts();
        //    LoadGoogleGroups();
        //    LoadGoogleContacts();
        //}

        //public void LoadNotes()
        //{
        //    LoadOutlookNotes();
        //    LoadGoogleNotes();
        //}

        /// <summary>
        /// Load the contacts from Google and Outlook and match them
        /// </summary>
  //      public void MatchContacts()
		//{
  //          LoadContacts();

		//	DuplicateDataException duplicateDataException;
		//	//_contactMatches = ContactsMatcher.MatchContacts(this, out duplicateDataException);
		//	if (duplicateDataException != null)
		//	{
				
		//		if (DuplicatesFound != null)
  //                  DuplicatesFound("Google duplicates found", duplicateDataException.Message);
  //              else
  //                  Logger.Log(duplicateDataException.Message, EventType.Warning);
		//	}
		//}


        /// <summary>
        /// Load the contacts from Google and Outlook and match them
        /// </summary>
        //public void MatchNotes()
        //{
        //    LoadNotes();
        //    _noteMatches = NotesMatcher.MatchNotes(this);
        //    /*DuplicateDataException duplicateDataException;
        //    _matches = ContactsMatcher.MatchContacts(this, out duplicateDataException);
        //    if (duplicateDataException != null)
        //    {

        //        if (DuplicatesFound != null)
        //            DuplicatesFound("Google duplicates found", duplicateDataException.Message);
        //        else
        //            Logger.Log(duplicateDataException.Message, EventType.Warning);
        //    }*/
        //}

		public void Sync()
		{
			lock (_syncRoot)
			{
                try { OutlookConnection.Connect(); }
                catch (OutlookConnectionException exc)
                {
                    ErrorHandler.Handle(exc);
                    return;
                }
                this.Initialize();
                if (!(this.SyncCalendar || this.SyncContacts))
                {
                    throw new NoDataToSyncronizeSpecifiedException("Neither calendar nor contacts are selected on for syncing. Please choose at least one option. Sync aborted!");
                }
                //SecurityManager sm = new SecurityManager();
                try
                {
                    /// We assume the synchronization fails and change our mind only in case no exceptions are thrown
                    this.Result.Succeeded = false;

                    if (_syncProfile.Length == 0)
                    {
                        Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                        return;
                    }
                    /// Disabling Outlook security warnings
                    //OutlookUtilities.TryDo(() => sm.ConnectTo(OutlookConnection.Application));
                    //OutlookUtilities.TryDo(() => sm.DisableOOMWarnings = true);

                    _skippedCount = 0;
                    _skippedCountNotMatches = 0;

                    //if (this._syncContacts)
                    //{
                    //    MatchContacts();
                    //    if (_contactMatches == null)
                    //        return;

                    //    _totalCount = _contactMatches.Count + _skippedCountNotMatches;

                    //    //Remove Google duplicates from matches to be synced
                    //    if (_googleContactDuplicates != null)
                    //    {
                    //        for (int i = _googleContactDuplicates.Count - 1; i >= 0; i--)
                    //        {
                    //            ContactMatch match = _googleContactDuplicates[i];
                    //            if (_contactMatches.Contains(match))
                    //            {
                    //                _skippedCount++;
                    //                _contactMatches.Remove(match);
                    //            }
                    //        }
                    //    }

                        ////Remove Outlook duplicates from matches to be synced
                        //if (_outlookContactDuplicates != null)
                        //{
                        //    for (int i = _outlookContactDuplicates.Count - 1; i >= 0; i--)
                        //    {
                        //        ContactMatch match = _outlookContactDuplicates[i];
                        //        if (_contactMatches.Contains(match))
                        //        {
                        //            _skippedCount++;
                        //            _contactMatches.Remove(match);
                        //        }
                        //    }
                        //}

                        ////Remove remaining google contacts not in My Contacts group (to avoid syncing accounts added automatically to "Weitere Kontakte"/"Further Contacts"
                        //Group syncGroup = GetGoogleGroupByName(myContactsGroup);
                        //if (syncGroup != null)
                        //{
                        //    for (int i = _googleContacts.Count -1 ;i >=0; i--)
                        //    {
                        //        Contact googleContact = _googleContacts[i];
                        //        Collection<Group> googleContactGroups = Utilities.GetGoogleGroups(this, googleContact);

                        //        if (!googleContactGroups.Contains(syncGroup))
                        //            _googleContacts.Remove(googleContact);

                        //    }
                        //}                                    


                        //Logger.Log("Syncing groups...", EventType.Information);
                        //ContactsMatcher.SyncGroups(this);

                        //Logger.Log("Syncing contacts...", EventType.Information);
                        //ContactsMatcher.SyncContacts(this);

                        //SaveContacts(_contactMatches);
                    //}

                    // This is new approach to synchronizing. Synchronizator shouldn't care on separate item types.
                    // It should prepare environment for the synchronization (login, get settings etc) and run synchronization for each 
                    // 

                    if (this.SyncCalendar)
                    {
                        this._syncBatch.Add(new CalendarSynchronizer(
                            this._googleCredential,
                            //this._outlookNamespace,
                            this.CalendarFolderToSyncID
                            ));
                    }

                    /// For future use
                    //if (this.SyncTasks)
                    //{
                    //    this._syncBatch.Add(new TaskSynchronizer(
                    //        this._googleAuthToken,
                    //        this._outlookNamespace
                    //        ));
                    //}

                    foreach (var sync in this._syncBatch)
                        this.Result += sync.Sync();

                    Logger.Log(String.Format("Synchronization is finished\r\n{0}", this.Result), EventType.Information);
                    this.Result.Succeeded = true;
                }
                catch (ArgumentException exc)
                {
                    ErrorHandler.Handle(exc);
                    Logger.Log("Supposedly wrong Outlook application argument. Outlook app profile name is '" + OutlookConnection.Application.DefaultProfileName + "'", EventType.Error);
                }
                catch (COMException exc)
                {
                    if (exc.ErrorCode == -2147467260) // E_ABORT
                    {
                    }
                }
                catch (Exception exc)
                {
                    ErrorHandler.Handle(exc);
                }
                finally
                {
                    //if (_outlookContacts != null)
                    //{
                    //    Marshal.ReleaseComObject(_outlookContacts);
                    //    _outlookContacts = null;
                    //}
                    //_googleContacts = null;
                    //_outlookContactDuplicates = null;
                    //_googleContactDuplicates = null;
                    //_googleGroups = null;
                    //_contactMatches = null;
                    //OutlookUtilities.TryDo(() => sm.DisableOOMWarnings = false);
                    //OutlookUtilities.TryDo(() => sm.Disconnect(OutlookConnection.Application));
                    OutlookConnection.Disconnect();
                }
			}
		}

		//public void SaveContacts(List<ContactMatch> contacts)
		//{
		//	foreach (ContactMatch match in contacts)
		//	{
		//		try
		//		{
		//			SaveContact(match);
		//		}
		//		catch (Exception ex)
		//		{
		//			if (ErrorEncountered != null)
		//			{
  //                      this._result.ErrorItems++;
  //                      this._result.UpdatedItems--;
  //                      string message = String.Format("Failed to synchronize contact: {0}. \nPlease check the contact, if any Email already exists on Google contacts side or if there is too much or invalid data in the notes field. \nIf the problem persists, please try recreating the contact or report the error on SourceForge.", match.OutlookContact.FileAs);
		//				Exception newEx = new Exception(message, ex);
		//				ErrorEncountered("Error", newEx, EventType.Error);
		//			}
		//			else
		//				throw;
		//		}
		//	}
		//}

        // NOTE: Outlook contacts are not saved here anymore, they have already been saved and counted
  //      public void SaveContact(ContactMatch match)
  //      {
  //          if (match.GoogleContact != null && match.OutlookContact != null)
		//	{
		//		//bool googleChanged, outlookChanged;
		//		//SaveContactGroups(match, out googleChanged, out outlookChanged);
  //              if (match.GoogleContact.ContactEntry.Dirty || match.GoogleContact.ContactEntry.IsDirty())
  //              {
  //                  //google contact was modified. save.
  //                  this._result.UpdatedItems++;					
		//			SaveGoogleContact(match);
		//			Logger.Log("Updated Google contact from Outlook: \"" + match.OutlookContact.FileAs + "\".", EventType.Information);
		//		}

  //              //if (!outlookContactItem.Saved)// || outlookChanged)
  //              //{
  //              //    //outlook contact was modified. save.
  //              //    SaveOutlookContact(match, outlookContactItem);
  //              //    Logger.Log("Updated Outlook contact from Google: \"" + outlookContactItem.FileAs + "\".", EventType.Information);
  //              //}                

		//		// save photos
		//		//SaveContactPhotos(match);
		//	}
  //          else if (match.GoogleContact == null && match.OutlookContact != null)
		//	{
  //              if (match.OutlookContact.UserProperties.GoogleContactId != null)
		//		{
  //                  string name = match.OutlookContact.FileAs;
  //                  if (_syncOption == SyncOption.OutlookToGoogleOnly)
  //                  {
  //                      _skippedCount++;
  //                      Logger.Log("Skipped Deletion of Outlook contact because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
  //                  }
  //                  else
  //                  {
  //                      // peer google contact was deleted, delete outlook contact
  //                      Outlook.ContactItem item = match.OutlookContact.GetOriginalItemFromOutlook(this);
  //                      try
  //                      {
  //                          item.Delete();
  //                          this._result.DeletedItems++;
  //                          Logger.Log("Deleted Outlook contact: \"" + name + "\".", EventType.Information);
  //                      }
  //                      finally
  //                      {
  //                          Marshal.ReleaseComObject(item);
  //                          item = null;
  //                      }
  //                  }
		//		}
		//	}
  //          else if (match.GoogleContact != null && match.OutlookContact == null)
		//	{
		//		if (ContactPropertiesUtils.GetGoogleOutlookContactId(SyncProfile, match.GoogleContact) != null)
		//		{
  //                  string name = match.GoogleContact.Title;
  //                  if (string.IsNullOrEmpty(name))
  //                      name = match.GoogleContact.Name.FullName;
  //                  if (string.IsNullOrEmpty(name) && match.GoogleContact.Organizations.Count > 0)
  //                      name = match.GoogleContact.Organizations[0].Name;
  //                  if (string.IsNullOrEmpty(name) && match.GoogleContact.Emails.Count > 0)
  //                      name = match.GoogleContact.Emails[0].Address;

  //                  if (_syncOption == SyncOption.GoogleToOutlookOnly)
  //                  {
  //                      _skippedCount++;
  //                      Logger.Log("Skipped Deletion of Google contact because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
  //                  }
  //                  else
  //                  {
  //                      // peer outlook contact was deleted, delete google contact
  //                      _contactsRequest.Delete(match.GoogleContact);
  //                      this._result.DeletedItems++;
  //                      Logger.Log("Deleted Google contact: \"" + name + "\".", EventType.Information);
  //                  }
		//		}
		//	}
		//	else
		//	{
		//		//TODO: ignore for now: 
  //              throw new ArgumentNullException("To save contacts, at least a GoogleContacat or OutlookContact must be present.");
		//		//Logger.Log("Both Google and Outlook contact: \"" + match.OutlookContact.FileAs + "\" have been changed! Not implemented yet.", EventType.Warning);
		//	}
		//}

        //public void SaveNote(NoteMatch match)
        //{
        //    if (match.GoogleNote != null && match.OutlookNote != null)
        //    {
        //        //bool googleChanged, outlookChanged;
        //        //SaveNoteGroups(match, out googleChanged, out outlookChanged);
        //        if (match.GoogleNote.DocumentEntry.Dirty || match.GoogleNote.DocumentEntry.IsDirty())
        //        {
        //            //google note was modified. save.
        //            _syncedCount++;
        //            SaveGoogleNote(match);
        //            Logger.Log("Updated Google note from Outlook: \"" + match.OutlookNote.Subject + "\".", EventType.Information);
        //        }

        //        if (!match.OutlookNote.Saved)// || outlookChanged)
        //        {
        //            //outlook note was modified. save.
        //            _syncedCount++;
        //            NotePropertiesUtils.SetOutlookGoogleNoteId(this, match.OutlookNote, match.GoogleNote);
        //            match.OutlookNote.Save();
        //            Logger.Log("Updated Outlook note from Google: \"" + match.OutlookNote.Subject + "\".", EventType.Information);
        //        }                

        //        // save photos
        //        //SaveNotePhotos(match);
        //    }
        //    else if (match.GoogleNote == null && match.OutlookNote != null)
        //    {
        //        if (match.OutlookNote.ItemProperties[this.OutlookPropertyNameId] != null)
        //        {
        //            string name = match.OutlookNote.Subject;
        //            if (_syncOption == SyncOption.OutlookToGoogleOnly)
        //            {
        //                _skippedCount++;
        //                Logger.Log("Skipped Deletion of Outlook note because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
        //            }
        //            else
        //            {
        //                // peer google note was deleted, delete outlook note
        //                Outlook.NoteItem item = match.OutlookNote;
        //                //try
        //                //{
        //                    item.Delete();
        //                    try
        //                    { //Delete also the according temporary NoteFile
        //                        File.Delete(NotePropertiesUtils.GetFileName(NotePropertiesUtils.GetOutlookGoogleNoteId(this,match.OutlookNote)));
        //                    }
        //                    catch (Exception)
        //                    { }
        //                    _deletedCount++;
        //                    Logger.Log("Deleted Outlook note: \"" + name + "\".", EventType.Information);
        //                //}
        //                //finally
        //                //{
        //                //    Marshal.ReleaseComObject(item);
        //                //    item = null;
        //                //}
        //            }
        //        }
        //    }
        //    else if (match.GoogleNote != null && match.OutlookNote == null)
        //    {
        //        if (NotePropertiesUtils.NoteFileExists(match.GoogleNote.Id))
        //        {
        //            string name = match.GoogleNote.Title;                    

        //            if (_syncOption == SyncOption.GoogleToOutlookOnly)
        //            {
        //                _skippedCount++;
        //                Logger.Log("Skipped Deletion of Google note because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
        //            }
        //            else
        //            {
        //                // peer outlook note was deleted, delete google note
        //                 _documentsRequest.Delete(match.GoogleNote);
                         
        //                //ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder, therefore the following workaround
        //                 Document deletedNote = LoadGoogleNotes(match.GoogleNote.DocumentEntry.Id);
        //                 if (deletedNote != null)
        //                     _documentsRequest.Delete(deletedNote);
                        
        //                try
        //                 {//Delete also the according temporary NoteFile
        //                     File.Delete(NotePropertiesUtils.GetFileName(match.GoogleNote.Id));
        //                 }
        //                 catch (Exception)
        //                 {}

        //                _deletedCount++;
        //                Logger.Log("Deleted Google note: \"" + name + "\".", EventType.Information);
        //            }
        //        }
        //    }
        //    else
        //    {
        //        //TODO: ignore for now: 
        //        throw new ArgumentNullException("To save notes, at least a GoogleContacat or OutlookNote must be present.");
        //        //Logger.Log("Both Google and Outlook note: \"" + match.OutlookNote.FileAs + "\" have been changed! Not implemented yet.", EventType.Warning);
        //    }
        //}

        //private void SaveOutlookContact(ref Contact googleContact, Outlook.ContactItem outlookContact)
        //{
        //    ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
        //    outlookContact.Save();
        //    ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, googleContact, outlookContact);

        //    Contact updatedEntry = SaveGoogleContact(googleContact);
        //    //try
        //    //{
        //    //    updatedEntry = _googleService.Update(match.GoogleContact);
        //    //}
        //    //catch (GDataRequestException tmpEx)
        //    //{
        //    //    // check if it's the known HTCData problem, or if there is any invalid XML element or any unescaped XML sequence
        //    //    //if (tmpEx.ResponseString.Contains("HTCData") || tmpEx.ResponseString.Contains("&#39") || match.GoogleContact.Content.Contains("<"))
        //    //    //{
        //    //    //    bool wasDirty = match.GoogleContact.ContactEntry.Dirty;
        //    //    //    // XML escape the content
        //    //    //    match.GoogleContact.Content = EscapeXml(match.GoogleContact.Content);
        //    //    //    // set dirty to back, cause we don't want the changed content go back to Google without reason
        //    //    //    match.GoogleContact.ContactEntry.Content.Dirty = wasDirty;
        //    //    //    updatedEntry = _googleService.Update(match.GoogleContact);
                    
        //    //    //}
        //    //    //else 
        //    //    if (!String.IsNullOrEmpty(tmpEx.ResponseString))
        //    //        throw new ApplicationException(tmpEx.ResponseString, tmpEx);
        //    //    else
        //    //        throw;
        //    //}            
        //    googleContact = updatedEntry;

        //    ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
        //    outlookContact.Save();
        //    SaveOutlookPhoto(googleContact, outlookContact);
        //}
		private string EscapeXml(string xml)
		{
			string encodedXml = System.Security.SecurityElement.Escape(xml);
			return encodedXml;
		}
		//public void SaveGoogleContact(ContactMatch match)
		//{
  //          Outlook.ContactItem outlookContactItem = match.OutlookContact.GetOriginalItemFromOutlook(this);
  //          try
  //          {
  //              ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, outlookContactItem);
  //              match.GoogleContact = SaveGoogleContact(match.GoogleContact);
  //              ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContactItem, match.GoogleContact);
  //              outlookContactItem.Save();

  //              //Now save the Photo
  //              SaveGooglePhoto(match, outlookContactItem);
  //          }
  //          finally
  //          {
  //              Marshal.ReleaseComObject(outlookContactItem);
  //              outlookContactItem = null;
  //          }
		//}

        //public void SaveGoogleNote(NoteMatch match)
        //{
        //    Outlook.NoteItem outlookNoteItem = match.OutlookNote;
        //    //try
        //    //{  

        //        //ToDo: Somewhow, the content is not uploaded to Google, only an empty document                
        //        //match.GoogleNote = SaveGoogleNote(match.GoogleNote);

        //        //ToDo: Therefoe I use DocumentService.UploadDocument instead and move it to the NotesFolder
        //        if (match.GoogleNote.DocumentEntry.Id.Uri != null)
        //        {
        //            _documentsRequest.Delete(match.GoogleNote);

        //            //ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder
        //            Document deletedNote = LoadGoogleNotes(match.GoogleNote.DocumentEntry.Id);
        //            if (deletedNote != null)
        //                _documentsRequest.Delete(deletedNote);

        //        }

        //        Google.GData.Documents.DocumentEntry entry = _documentsRequest.Service.UploadDocument(NotePropertiesUtils.GetFileName(outlookNoteItem.EntryID), match.GoogleNote.Title.Replace(":", String.Empty));                               
        //        Document newNote = LoadGoogleNotes(entry.Id);
        //        match.GoogleNote = _documentsRequest.MoveDocumentTo(_googleNotesFolder, newNote);               
                
        //        NotePropertiesUtils.SetOutlookGoogleNoteId(this, outlookNoteItem, match.GoogleNote);
        //        outlookNoteItem.Save();

        //        //As GoogleDocuments don't have UserProperties, we have to use the file to check, if Note was already synced or not
        //        File.Delete(NotePropertiesUtils.GetFileName(match.GoogleNote.Id));
        //        File.Move(NotePropertiesUtils.GetFileName(outlookNoteItem.EntryID), NotePropertiesUtils.GetFileName(match.GoogleNote.Id));
        //    //}
        //    //finally
        //    //{
        //    //    Marshal.ReleaseComObject(outlookNoteItem);
        //    //    outlookNoteItem = null;
        //    //}
        //}

		//private string GetXml(Contact contact)
		//{
		//	MemoryStream ms = new MemoryStream();
		//	contact.ContactEntry.SaveToXml(ms);
		//	StreamReader sr = new StreamReader(ms);
		//	ms.Seek(0, SeekOrigin.Begin);
		//	string xml = sr.ReadToEnd();
		//	return xml;
		//}

  //      private string GetXml(Document note)
  //      {
  //          MemoryStream ms = new MemoryStream();
  //          note.DocumentEntry.SaveToXml(ms);
  //          StreamReader sr = new StreamReader(ms);
  //          ms.Seek(0, SeekOrigin.Begin);
  //          string xml = sr.ReadToEnd();
  //          return xml;
  //      }

        /// <summary>
        /// Only save the google contact without photo update
        /// </summary>
        /// <param name="googleContact"></param>
		//internal Contact SaveGoogleContact(Contact googleContact)
		//{
		//	//check if this contact was not yet inserted on google.
		//	if (googleContact.ContactEntry.Id.Uri == null)
		//	{
		//		//insert contact.
		//		Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

		//		try
		//		{
		//			Contact createdEntry = _contactsRequest.Insert(feedUri, googleContact);
  //                  return createdEntry;
		//		}
  //              catch (Exception ex)
  //              {
  //                  string responseString = "";
  //                  if (ex is GDataRequestException)
  //                      responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
  //                  string xml = GetXml(googleContact);
  //                  string newEx = String.Format("Error saving NEW Google contact: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
  //                  throw new ApplicationException(newEx, ex);
  //              }
		//	}
		//	else
		//	{
		//		try
		//		{
		//			//contact already present in google. just update
					
  //                  // User can create an empty label custom field on the web, but when I retrieve, and update, it throws this:
  //                  // Data Request Error Response: [Line 12, Column 44, element gContact:userDefinedField] Missing attribute: &#39;key&#39;
  //                  // Even though I didn't touch it.  So, I will search for empty keys, and give them a simple name.  Better than deleting...
  //                  int fieldCount = 0;
  //                  foreach (UserDefinedField userDefinedField in googleContact.ContactEntry.UserDefinedFields)
  //                  {
  //                      fieldCount++;
  //                      if (String.IsNullOrEmpty(userDefinedField.Key))
  //                      {
  //                          userDefinedField.Key = "UserField" + fieldCount.ToString();
  //                      }
  //                  }

  //                  //TODO: this will fail if original contact had an empty name or rpimary email address.
  //                  Contact updated = _contactsRequest.Update(googleContact);
  //                  return updated;
		//		}
  //              catch (Exception ex)
  //              {
  //                  string responseString = "";
  //                  if (ex is GDataRequestException)
  //                      responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
  //                  string xml = GetXml(googleContact);
  //                  string newEx = String.Format("Error saving EXISTING Google contact: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
  //                  throw new ApplicationException(newEx, ex);
  //              }
		//	}
		//}

        /// <summary>
        /// save the google note
        /// </summary>
        /// <param name="googleNote"></param>
        //internal Document SaveGoogleNote(Document googleNote)
        //{
        //    //check if this contact was not yet inserted on google.
        //    if (googleNote.DocumentEntry.Id.Uri == null)
        //    {
        //        //insert contact.
        //        Uri feedUri = null;

        //        try
        //        {//In case of Notes folder creation, the _googleNotesFolder.DocumentEntry.Content.AbsoluteUri throws a NullReferenceException
        //            feedUri = new Uri(_googleNotesFolder.DocumentEntry.Content.AbsoluteUri);
        //        }
        //        catch (Exception)
        //        { }

        //        if (feedUri == null)                
        //            feedUri = new Uri(_documentsRequest.BaseUri);               

        //        try
        //        {
        //            Document createdEntry = _documentsRequest.Insert(feedUri, googleNote);
        //            //ToDo: Workaround also doesn't help: Utilities.SaveGoogleNoteContent(this, createdEntry, googleNote);    
        //            return createdEntry;
        //        }
        //        catch (Exception ex)
        //        {
        //            string responseString = "";
        //            if (ex is GDataRequestException)
        //                responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
        //            string xml = GetXml(googleNote);
        //            string newEx = String.Format("Error saving NEW Google note: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
        //            throw new ApplicationException(newEx, ex);
        //        }
        //    }
        //    else
        //    {
        //        try
        //        {
        //            //note already present in google. just update
        //            Document updated = _documentsRequest.Update(googleNote);

        //            //ToDo: Workaround also doesn't help: Utilities.SaveGoogleNoteContent(this, updated, googleNote);                   

        //            return updated;
        //        }
        //        catch (Exception ex)
        //        {
        //            string responseString = "";
        //            if (ex is GDataRequestException)
        //                responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
        //            string xml = GetXml(googleNote);
        //            string newEx = String.Format("Error saving EXISTING Google note: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
        //            throw new ApplicationException(newEx, ex);
        //        }
        //    }
        //}         

  //      public void SaveGooglePhoto(ContactMatch match, Outlook.ContactItem outlookContactitem)
  //      {
  //          bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
  //          bool hasOutlookPhoto = Utilities.HasPhoto(outlookContactitem);

  //          if (hasOutlookPhoto)
  //          {
  //              // add outlook photo to google
  //              Image outlookPhoto = Utilities.GetOutlookPhoto(outlookContactitem);

  //              if (outlookPhoto != null)
  //              {
  //                  using (MemoryStream stream = new MemoryStream(Utilities.BitmapToBytes(new Bitmap(outlookPhoto))))
  //                  {
  //                      // Save image to stream.
  //                      //outlookPhoto.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);

  //                      //Don'T crop, because maybe someone wants to keep his photo like it is on Outlook
  //                      //outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
  //                      _contactsRequest.SetPhoto(match.GoogleContact, stream);

  //                      //Just save the Outlook Contact to have the same lastUpdate date as Google
  //                      ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContactitem, match.GoogleContact);
  //                      outlookContactitem.Save();
  //                      outlookPhoto.Dispose();
                        
  //                  }
  //              }
  //          }
  //          else if (hasGooglePhoto)
  //          {
  //              //Delete Photo on Google side, if no Outlook photo exists
  //              _contactsRequest.Delete(match.GoogleContact.PhotoUri, match.GoogleContact.PhotoEtag);
  //          }

  //          Utilities.DeleteTempPhoto();
  //      }

  //      public void SaveOutlookPhoto(Contact googleContact, Outlook.ContactItem outlookContact)
  //      {
  //          bool hasGooglePhoto = Utilities.HasPhoto(googleContact);
  //          bool hasOutlookPhoto = Utilities.HasPhoto(outlookContact);

  //          if (hasGooglePhoto)
  //          {
  //              // add google photo to outlook
  //              //ToDo: add google photo to outlook with new Google API
  //              //Stream stream = _googleService.GetPhoto(match.GoogleContact);
  //              Image googlePhoto = Utilities.GetGooglePhoto(this, googleContact);
  //              if (googlePhoto != null)    // Google may have an invalid photo
  //              {
  //                  Utilities.SetOutlookPhoto(outlookContact, googlePhoto);
  //                  ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
  //                  outlookContact.Save();

  //                  googlePhoto.Dispose();
  //              }
  //          }
  //          else if (hasOutlookPhoto)
  //          {
  //              outlookContact.RemovePicture();
  //              ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
  //              outlookContact.Save();
  //          }
  //      }

	
		//public Group SaveGoogleGroup(Group group)
		//{
		//	//check if this group was not yet inserted on google.
		//	if (group.GroupEntry.Id.Uri == null)
		//	{
		//		//insert group.
		//		Uri feedUri = new Uri(GroupsQuery.CreateGroupsUri("default"));

		//		try
		//		{
		//			return _contactsRequest.Insert(feedUri, group);
		//		}
		//		catch
		//		{
		//			//TODO: save google group xml for diagnistics
		//			throw;
		//		}
		//	}
		//	else
		//	{
		//		try
		//		{
		//			//group already present in google. just update
		//			return _contactsRequest.Update(group);
		//		}
		//		catch
		//		{
		//			//TODO: save google group xml for diagnistics
		//			throw;
		//		}
		//	}
		//}

        /// <summary>
        /// Updates Google contact from Outlook (including groups/categories)
        /// </summary>
        //public void UpdateContact(Outlook.ContactItem master, Contact slave)
        //{
        //    ContactSync.UpdateContact(master, slave);
        //    OverwriteContactGroups(master, slave);
        //}

        /// <summary>
        /// Updates Outlook contact from Google (including groups/categories)
        /// </summary>
        //public void UpdateContact(Contact master, Outlook.ContactItem slave)
        //{
        //    ContactSync.UpdateContact(master, slave);
        //    OverwriteContactGroups(master, slave);

        //    // -- Immediately save the Outlook contact (including groups) so it can be released, and don't do it in the save loop later
        //    SaveOutlookContact(ref master, slave);
        //    this._result.UpdatedItems++;
        //    Logger.Log("Updated Outlook contact from Google: \"" + slave.FileAs + "\".", EventType.Information);
        //}

        /// <summary>
        /// Updates Google note from Outlook
        /// </summary>
        //public void UpdateNote(Outlook.NoteItem master, Document slave)
        //{
        //    slave.Title = master.Subject;

        //    string fileName = NotePropertiesUtils.CreateNoteFile(master.EntryID, master.Body);

        //    string contentType = MediaFileSource.GetContentTypeForFileName(fileName);

        //    //ToDo: Somewhow, the content is not uploaded to Google, only an empty document
        //    //Therefoe I use DocumentService.UploadDocument instead.
        //    slave.MediaSource = new MediaFileSource(fileName, contentType);

        //}

        


        /// <summary>
        /// Updates Outlook contact from Google
        /// </summary>
        //public void UpdateNote(Document master, Outlook.NoteItem slave)
        //{
        //    //slave.Subject = master.Title; //The Subject is readonly and set automatically by Outlook
        //    string body = NotePropertiesUtils.GetBody(this, master);

        //    if (string.IsNullOrEmpty(body) && slave.Body != null)
        //    {
        //        //DialogResult result = MessageBox.Show("The body of Google note '" + master.Title + "' is empty. Do you really want to syncronize an empty Google note to a not yet empty Outlook note?", "Empty Google Note", MessageBoxButtons.YesNo);

        //        //if (result != DialogResult.Yes)
        //        //{
        //        //    Logger.Log("The body of Google note '" + master.Title + "' is empty. The user decided to skip this note and not to syncronize an empty Google note to a not yet empty Outlook note.", EventType.Information);
        //            Logger.Log("The body of Google note '" + master.Title + "' is empty. It is skipped from syncing, because Outlook note is not empty.", EventType.Warning);
        //            SkippedCount++;
        //            return;
        //        //}
        //        //Logger.Log("The body of Google note '" + master.Title + "' is empty. The user decided to syncronize an empty Google note to a not yet empty Outlook note (" + slave.Body + ").", EventType.Warning);                
                
        //    }

        //    slave.Body = body;

        //    NotePropertiesUtils.CreateNoteFile(master.Id, body);

        //}

		/// <summary>
		/// Updates Google contact's groups from Outlook contact
		/// </summary>
		//private void OverwriteContactGroups(Outlook.ContactItem master, Contact slave)
		//{
		//	Collection<Group> currentGroups = Utilities.GetGoogleGroups(this, slave);

		//	// get outlook categories
		//	string[] cats = Utilities.GetOutlookGroups(master.Categories);

		//	// remove obsolete groups
		//	Collection<Group> remove = new Collection<Group>();
		//	bool found;
		//	foreach (Group group in currentGroups)
		//	{
		//		found = false;
		//		foreach (string cat in cats)
		//		{
		//			if (group.Title == cat)
		//			{
		//				found = true;
		//				break;
		//			}
		//		}
		//		if (!found)
		//			remove.Add(group);
		//	}
		//	while (remove.Count != 0)
		//	{
		//		Utilities.RemoveGoogleGroup(slave, remove[0]);
		//		remove.RemoveAt(0);
		//	}

		//	// add new groups
		//	Group g;
		//	foreach (string cat in cats)
		//	{
		//		if (!Utilities.ContainsGroup(this, slave, cat))
		//		{
		//			// add group to contact
		//			g = GetGoogleGroupByName(cat);
		//			if (g == null)
		//			{
		//				throw new Exception(string.Format("Google Groups were supposed to be created prior to saving", cat));
		//			}
		//			Utilities.AddGoogleGroup(slave, g);
		//		}
		//	}

  //          //add system Group My Contacts            
  //          if (!Utilities.ContainsGroup(this, slave, myContactsGroup))
  //          {
  //              // add group to contact
  //              g = GetGoogleGroupByName(myContactsGroup);
  //              if (g == null)
  //              {
  //                  throw new Exception(string.Format("Google System Group: My Contacts doesn't exist", myContactsGroup));
  //              }
  //              Utilities.AddGoogleGroup(slave, g);
  //          }
		//}

		///// <summary>
		///// Updates Outlook contact's categories (groups) from Google groups
		///// </summary>
		//private void OverwriteContactGroups(Contact master, Outlook.ContactItem slave)
		//{
		//	Collection<Group> newGroups = Utilities.GetGoogleGroups(this, master);

		//	List<string> newCats = new List<string>(newGroups.Count);
		//	foreach (Group group in newGroups)
  //          {   //Only add groups that are no SystemGroup (e.g. "System Group: Meine Kontakte") automatically tracked by Google
  //              if (group.Title != null && !group.Title.Equals(myContactsGroup))
		//		    newCats.Add(group.Title);
		//	}

		//	slave.Categories = string.Join(", ", newCats.ToArray());
		//}

		/// <summary>
		/// Resets associantions of Outlook contacts with Google contacts via user props
		/// and resets associantions of Google contacts with Outlook contacts via extended properties.
		/// </summary>
		//public void ResetContactMatches()
		//{
		//	Debug.Assert(_outlookContacts != null, "Outlook Contacts object is null - this should not happen. Please inform Developers.");
  //          Debug.Assert(_googleContacts != null, "Google Contacts object is null - this should not happen. Please inform Developers.");

  //          try
  //          {
  //              if (_syncProfile.Length == 0)
  //              {
  //                  Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
  //                  return;
  //              }
               

		//	    lock (_syncRoot)
		//	    {
  //                  Logger.Log("Resetting Google Contact matches...", EventType.Information);
		//		    foreach (Contact googleContact in _googleContacts)
		//		    {
  //                      try
  //                      {
  //                          if (googleContact != null)
  //                              ResetMatch(googleContact);
  //                      }
  //                      catch (Exception ex)
  //                      {
  //                          string name =googleContact.Title;
  //                          if (string.IsNullOrEmpty(name))
  //                              name = googleContact.Name.FullName;
  //                          if (string.IsNullOrEmpty(name) && googleContact.Organizations.Count > 0)
  //                              name = googleContact.Organizations[0].Name;
  //                          if (string.IsNullOrEmpty(name) && googleContact.Emails.Count > 0)
  //                              name = googleContact.Emails[0].Address;

  //                          Logger.Log("The match of Google contact " + name + " couldn't be reset: " + ex.Message, EventType.Warning);
  //                      }
		//		    }

  //                  Logger.Log("Resetting Outlook Contact matches...", EventType.Information);
  //                  //1 based array
  //                  for (int i=1; i <= _outlookContacts.Count; i++)
  //                  {
  //                      Outlook.ContactItem outlookContact = null;

  //                      try
  //                      {
  //                          outlookContact = _outlookContacts[i] as Outlook.ContactItem;
  //                          if (outlookContact == null)
  //                          {
  //                              Logger.Log("Empty Outlook contact found (maybe distribution list). Skipping", EventType.Warning);
  //                              continue;
  //                          }
  //                      }
  //                      catch (Exception ex)
  //                      {
  //                          //this is needed because some contacts throw exceptions
  //                          Logger.Log("Accessing Outlook contact threw and exception. Skipping: " + ex.Message, EventType.Warning);                               
  //                          continue;
  //                      }

  //                      try
  //                      {
  //                          ResetMatch(outlookContact);                            
  //                      }
  //                      catch (Exception ex)
  //                      {
  //                          Logger.Log("The match of Outlook contact " + outlookContact.FileAs + " couldn't be reset: " + ex.Message, EventType.Warning);
  //                      }
  //                  }

  //              }
  //          }
  //          finally
  //          {
  //              if (_outlookContacts != null)
  //              {
  //                  Marshal.ReleaseComObject(_outlookContacts);
  //                  _outlookContacts = null;
  //              }
  //              _googleContacts = null;
  //          }
						
		//}

  //      /// <summary>
  //      /// Resets associantions of Outlook notes with Google contacts via user props
  //      /// and resets associantions of Google contacts with Outlook contacts via extended properties.
  //      /// </summary>
  //      public void ResetNoteMatches()
  //      {
  //          Debug.Assert(_outlookNotes != null, "Outlook Notes object is null - this should not happen. Please inform Developers.");            

  //          //try
  //          //{
  //              if (_syncProfile.Length == 0)
  //              {
  //                  Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
  //                  return;
  //              }


  //              lock (_syncRoot)
  //              {
  //                  Logger.Log("Resetting Google Note matches...", EventType.Information);

  //                  try
  //                  {
  //                      NotePropertiesUtils.DeleteNoteFiles();
  //                  }
  //                  catch (Exception ex)
  //                  {                           
  //                      Logger.Log("The Google Note matches couldn't be reset: " + ex.Message, EventType.Warning);
  //                  }
                    

  //                  Logger.Log("Resetting Outlook Note matches...", EventType.Information);
  //                  //1 based array
  //                  for (int i = 1; i <= _outlookNotes.Count; i++)
  //                  {
  //                      Outlook.NoteItem outlookNote = null;

  //                      try
  //                      {
  //                          outlookNote = _outlookNotes[i] as Outlook.NoteItem;
  //                          if (outlookNote == null)
  //                          {
  //                              Logger.Log("Empty Outlook Note found (maybe distribution list). Skipping", EventType.Warning);
  //                              continue;
  //                          }
  //                      }
  //                      catch (Exception ex)
  //                      {
  //                          //this is needed because some notes throw exceptions
  //                          Logger.Log("Accessing Outlook Note threw and exception. Skipping: " + ex.Message, EventType.Warning);
  //                          continue;
  //                      }

  //                      try
  //                      {
  //                          ResetMatch(outlookNote);
  //                      }
  //                      catch (Exception ex)
  //                      {
  //                          Logger.Log("The match of Outlook note " + outlookNote.Subject + " couldn't be reset: " + ex.Message, EventType.Warning);
  //                      }
  //                  }

  //              }
  //          //}
  //          //finally
  //          //{
  //          //    if (_outlookContacts != null)
  //          //    {
  //          //        Marshal.ReleaseComObject(_outlookContacts);
  //          //        _outlookContacts = null;
  //          //    }
  //          //    _googleContacts = null;
  //          //}

  //      }


        /// <summary>
        /// Reset the match link between Google and Outlook contact        
        /// </summary>
        //public void ResetMatch(Contact googleContact)
        //{
            
        //    if (googleContact != null)
        //    {
        //        ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, googleContact);
        //        SaveGoogleContact(googleContact);
        //    }
        //}

        /// <summary>
        /// Reset the match link between Outlook and Google contact
        /// </summary>
        //public void ResetMatch(Outlook.ContactItem outlookContact)
        //{           

        //    if (outlookContact != null)
        //    {
        //        try
        //        {
        //            ContactPropertiesUtils.ResetOutlookGoogleContactId(this, outlookContact);
        //            outlookContact.Save();
        //        }
        //        finally
        //        {
        //            Marshal.ReleaseComObject(outlookContact);
        //            outlookContact = null;
        //        }
                
        //    }
        //}

        /// <summary>
        /// Reset the match link between Outlook and Google note
        /// </summary>
        //public void ResetMatch(Outlook.NoteItem outlookNote)
        //{

        //    if (outlookNote != null)
        //    {
        //        //try
        //        //{
        //            NotePropertiesUtils.ResetOutlookGoogleNoteId(this, outlookNote);
        //            outlookNote.Save();
        //        //}
        //        //finally
        //        //{
        //        //    Marshal.ReleaseComObject(outlookNote);
        //        //    outlookNote = null;
        //        //}

        //    }


        //}

        //public ContactMatch ContactByProperty(string name, string email)
        //{
        //    foreach (ContactMatch m in Contacts)
        //    {
        //        if (m.GoogleContact != null &&
        //            ((m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email) ||
        //            m.GoogleContact.Title == name))
        //        {
        //            return m;
        //        }
        //        else if (m.OutlookContact != null && (
        //            (m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email) ||
        //            m.OutlookContact.FileAs == name))
        //        {
        //            return m;
        //        }
        //    }
        //    return null;
        //}
		
        ////public ContactMatch ContactEmail(string email)
        //{
        //    foreach (ContactMatch m in Contacts)
        //    {
        //        if (m.GoogleContact != null &&
        //            (m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email))
        //        {
        //            return m;
        //        }
        //        else if (m.OutlookContact != null && (
        //            m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email))
        //        {
        //            return m;
        //        }
        //    }
        //    return null;
        //}

		/// <summary>
		/// Used to find duplicates.
		/// </summary>
		/// <param name="name"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		//public Collection<OutlookContactInfo> OutlookContactByProperty(string name, string value)
		//{
  //          Collection<OutlookContactInfo> col = new Collection<OutlookContactInfo>();
  //          //foreach (Outlook.ContactItem outlookContact in OutlookContacts)
  //          //{
  //          //    if (outlookContact != null && (
  //          //        (outlookContact.Email1Address != null && outlookContact.Email1Address == email) ||
  //          //        outlookContact.FileAs == name))
  //          //    {
  //          //        col.Add(outlookContact);
  //          //    }
  //          //}
  //          Outlook.ContactItem item = null;
  //          try
  //          {
  //              item = OutlookContacts.Find("["+name+"] = \"" + value + "\"") as Outlook.ContactItem;
  //              while (item != null)
  //              {
  //                  col.Add(new OutlookContactInfo(item, this));
  //                  Marshal.ReleaseComObject(item);
  //                  item = OutlookContacts.FindNext() as Outlook.ContactItem;
  //              }
  //          }
  //          catch (Exception)
		//	{
		//		//TODO: should not get here.
		//	}

		//	return col;
		//}
        
		//public Group GetGoogleGroupById(string id)
		//{
		//	//return _googleGroups.FindById(new AtomId(id)) as Group;
  //          foreach (Group group in _googleGroups)
  //          {
  //              if (group.GroupEntry.Id.Equals(new AtomId(id)))
  //                  return group;
  //          }
  //          return null;
		//}

		//public Group GetGoogleGroupByName(string name)
		//{
		//	foreach (Group group in _googleGroups)
		//	{
		//		if (group.Title == name)
		//			return group;
		//	}
		//	return null;
		//}

  //      public Contact GetGoogleContactById(string id)
  //      {
  //          foreach (Contact contact in _googleContacts)
  //          {
  //              if (contact.ContactEntry.Id.Equals(new AtomId(id)))
  //                  return contact;
  //          }
  //          return null;
  //      }

		//public Group CreateGroup(string name)
		//{
		//	Group group = new Group();
		//	group.Title = name;
		//	group.GroupEntry.Dirty = true;
		//	return group;
		//}

		//public static bool AreEqual(Outlook.ContactItem c1, Outlook.ContactItem c2)
		//{
		//	return c1.Email1Address == c2.Email1Address;
		//}
		//public static int IndexOf(Collection<Outlook.ContactItem> col, Outlook.ContactItem outlookContact)
		//{

		//	for (int i = 0; i < col.Count; i++)
		//	{
		//		if (AreEqual(col[i], outlookContact))
		//			return i;
		//	}
		//	return -1;
		//}

        internal void Initialize()
        {
            this.Result = new SyncResult();
            this._syncBatch.Clear();
        }

        internal void Unpair()
        {
            //try
            //{
                this.Initialize();
                try { OutlookConnection.Connect(); }
                catch (OutlookConnectionException exc)
                {
                    ErrorHandler.Handle(exc);
                    return;
                }
                //if (this.SyncContacts)
                //{
                //    this.LoadContacts();
                //    this.ResetContactMatches();
                //}
                if (this.SyncCalendar)
                {
                    this._syncBatch.Add(new CalendarSynchronizer(
                        this._googleCredential,
                        this.CalendarFolderToSyncID
                        ));
                }
                foreach (var syncItem in this._syncBatch)
                    syncItem.Unpair();
                 OutlookConnection.Disconnect();
           //}
           // finally
           // {
           // }
        }
    }
}
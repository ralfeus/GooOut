using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Threading;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Net;
//using R.Utilities;
using System.Globalization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace R.GoogleOutlookSync
{
    internal partial class SettingsForm : Form
	{
		private Synchronizer _sync;
        internal Synchronizer Synchronizer
        {
            get
            {
                if (this._sync == null)
                {
                    this._sync = new Synchronizer();
                    _sync.DuplicatesFound += new Synchronizer.DuplicatesFoundHandler(OnDuplicatesFound);
                    _sync.ErrorEncountered += new Synchronizer.ErrorNotificationHandler(OnErrorEncountered);
                    //this._sync.SyncProfile = this.tbSyncProfile.Text;
                    this._sync.SyncProfile = this.cmbOutlookProfiles.Text;
                }
                return this._sync;
            }
        }

		private SyncOption _syncOption;
		private DateTime lastSync;
		private bool requestClose = false;
        private bool boolShowBalloonTip = true;
        private LogForm _log = new LogForm();
        private Thread _syncThread;
        private string _calendarFolderToSyncID;

        //register window for lock/unlock messages of workstation
        private bool registered = false;

		delegate void TextHandler(string text);

		public SettingsForm()
		{
			InitializeComponent();
#if DEBUG
            this.resetMatchesLinkLabel.Visible = true;
#endif
			Logger.LogUpdated += new Logger.LogUpdatedHandler(Logger_LogUpdated);

            /// Need to let menu item "Show settings" know form is closed so it updates check status accordingly
            this.FormClosing += new FormClosingEventHandler(this.mniSettings_Click);

            /// Tie page buttons to panels
            this.tbtnAdvanced.Tag = this.pnlAdvanced;
            this.tbtnGeneral.Tag = this.pnlGeneral;
            this.tbtnProxy.Tag = this.pnlProxy;
            this.tbtnGeneral.PerformClick();

            //ContactsMatcher.NotificationReceived += new ContactsMatcher.NotificationHandler(OnNotificationReceived);
            //NotesMatcher.NotificationReceived += new NotesMatcher.NotificationHandler(OnNotificationReceived);

            this.InitializeLanguages();
            this.InitializeOutlookProfiles();
            this.LoadSettings();
            this.InitializeProxy();

            /// Need to let menu item "Show log" know form is closed so it updates check status accordingly
            this._log.FormClosing +=new FormClosingEventHandler(this.mniLog_Click);
			lastSync = DateTime.Now/*.AddSeconds(90)*/ - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
            TimerSwitch(this.mniAutoSync.Checked);
			this._log.LastSyncLabel.Text = "Not synced";

			ValidateSyncButton();

            // requires Windows XP or higher
            bool XpOrHigher = Environment.OSVersion.Platform == PlatformID.Win32NT &&
                                (Environment.OSVersion.Version.Major > 5 ||
                                    (Environment.OSVersion.Version.Major == 5 &&
                                     Environment.OSVersion.Version.Minor >= 1));

            if (XpOrHigher)
                registered = WTSRegisterSessionNotification(Handle, 0);
		}

        private void InitializeOutlookProfiles()
        {
            this.cmbOutlookProfiles.Items.AddRange(OutlookUtilities.GetOutlookProfiles());
        }

        //private void SetLanguage()
        //{
        //    RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
        //    if (regKeyAppRoot.GetValue("Language") != null)
        //        Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo((int)regKeyAppRoot.GetValue("Language"));
        //}

        private void InitializeLanguages()
        {
            string executablePath = Path.GetDirectoryName(Application.ExecutablePath);
            string[] directories = Directory.GetDirectories(executablePath);
            int pathLength = executablePath.Length + 1;
            foreach (string s in directories)
            {
                try
                {
                    this.cmbLanguage.Items.Add(CultureInfo.GetCultureInfo(s.Remove(0, pathLength)));
                }
                catch (Exception) { }
            }
            this.cmbLanguage.Items.Add(CultureInfo.GetCultureInfo("en"));
            if (this.cmbLanguage.Items.Contains(Thread.CurrentThread.CurrentUICulture))
                this.cmbLanguage.SelectedItem = Thread.CurrentThread.CurrentUICulture;
            else
                this.cmbLanguage.SelectedItem = CultureInfo.GetCultureInfo("en");
        }

//        private UpdateResult CheckUpdates()
//        {
//            try
//            {
//                var result = Updater.Update(
//                    String.Format("{0}/{1}/{2}",
//                    Properties.Settings.Default.URL_Update,
//                    Assembly.GetExecutingAssembly().GetName().Version.Major,
//                    Utilities.GetAssemblyArchitecture().ToString().ToLower()));
//                //var result = Updater.Update("http://stone/download/1/");
//                if (result == UpdateResult.UpToDate)
//                {
//#if DEBUG
//                    Logger.Log("The application is up to date", EventType.Information);
//#endif
//                }
//                else if (result == UpdateResult.Restart)
//                {
//                    try
//                    {
//                        Logger.Log("The application is updated. Restart is needed", EventType.Information);
//                    }
//                    catch (Exception)
//                    { }
//                    MessageBox.Show("The application is updated. Restart is needed", Application.ProductName);
//                    this.Restart();
//                }
//                else if (result == UpdateResult.Success)
//                {
//                    try
//                    {
//                        Logger.Log("The application is updated successfully. Current version is " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version, EventType.Information);
//                    }
//                    catch (Exception)
//                    { }
//                    MessageBox.Show("The application is updated successfully. Current version is " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version, Application.ProductName);
//                }
//                else if (result == UpdateResult.Fail)
//                {
//                    try
//                    {
//                        Logger.Log("Application failed to update.", EventType.Error);
//                    }
//                    catch (Exception)
//                    { }
//                }
//                else if (result == UpdateResult.AccessDenied)
//                {
//                    if (MessageBox.Show(Properties.Resources.Confirm_RestartAsAdminNeeded, Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
//                        this.Restart(true);
//                }

//                return result;
//            }
//            catch (Exception exc)
//            {
//                ErrorHandler.Handle(exc);
//                return UpdateResult.Fail;
//            }
//        }

        private void Restart(bool asAdmin = false)
        {
            ProcessStartInfo si = new ProcessStartInfo(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            if (asAdmin)
                si.Verb = "runas";
            Process.Start(si);
            Program.ProgramMutex.ReleaseMutex();
            this.Exit();
        }

        private void InitializeProxy()
        {
            this.LoadProxySettings();
            this.CustomProxy_Changed(null, null);
            this.SetProxy();
        }

        ~SettingsForm()
        {
            if(registered)
            {
                WTSUnRegisterSessionNotification(Handle);
                registered = false;
            }
            Logger.Close();
        }

        private void LoadProxySettings()
        {   
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
            if (regKeyAppRoot.GetValue("ProxyUsage") != null)
            {
                if (Convert.ToBoolean(regKeyAppRoot.GetValue("ProxyUsage")))
                {
                    this.CustomProxy.Checked = true;
                    this.SystemProxy.Checked = !CustomProxy.Checked;

                    if (regKeyAppRoot.GetValue("ProxyURL") != null)
                        this.Address.Text = (string)regKeyAppRoot.GetValue("ProxyURL");

                    if (regKeyAppRoot.GetValue("ProxyPort") != null)
                        this.Port.Text = (string)regKeyAppRoot.GetValue("ProxyPort");

                    if (Convert.ToBoolean(regKeyAppRoot.GetValue("ProxyAuth")))
                    {
                        this.Authorization.Checked = true;

                        if (regKeyAppRoot.GetValue("ProxyUsername") != null)
                        {
                            this.txtProxyUserName.Text = regKeyAppRoot.GetValue("ProxyUsername") as string;
                            if (regKeyAppRoot.GetValue("ProxyPassword") != null)
                                this.txtProxyPassword.Text = Encryption.DecryptPassword(this.txtProxyUserName.Text, regKeyAppRoot.GetValue("ProxyPassword") as string);
                        }
                    }
                }
            }
        }

		private void LoadSettings()
		{
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
            /// Setting of necessary Outlook profile is essential for next settings
            if (regKeyAppRoot.GetValue("OutlookProfile") != null)
                this.SetOutlookProfile((string)regKeyAppRoot.GetValue("OutlookProfile"));
            else
                this.SetOutlookProfile(OutlookUtilities.GetDefaultOutlookProfile());
            // default
            SetSyncOption(0);
            OutlookConnection.Connect();
            SetCalendarFolder(OutlookConnection.Namespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).EntryID);

            // load
            if (regKeyAppRoot.GetValue("SyncOption") != null)
            {
                _syncOption = (SyncOption)regKeyAppRoot.GetValue("SyncOption");
                SetSyncOption((int)_syncOption);
            }
            if (regKeyAppRoot.GetValue("GoogleAccountName") != null)
            {
                this.txtGoogleAccountName.Text = this._googleAccountName = regKeyAppRoot.GetValue("GoogleAccountName") as string;
            }
            if (regKeyAppRoot.GetValue("AutoSync") != null)
                this.SetAutoSync(Convert.ToBoolean(regKeyAppRoot.GetValue("AutoSync")));
            if (regKeyAppRoot.GetValue("AutoSyncInterval") != null)
                autoSyncInterval.Value = Convert.ToDecimal(regKeyAppRoot.GetValue("AutoSyncInterval"));
            if (regKeyAppRoot.GetValue("ReportSyncResult") != null)
                reportSyncResultCheckBox.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("ReportSyncResult"));
            //if (regKeyAppRoot.GetValue("SyncDeletion") != null)
            //    btSyncDelete.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncDeletion"));
            if (regKeyAppRoot.GetValue("SyncCalendar") != null)
                btSyncCalendar.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncCalendar"));
            if (regKeyAppRoot.GetValue("SyncContacts") != null)
                btSyncContacts.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncContacts"));

            //if (regKeyAppRoot.GetValue("SyncNotes") != null)
            //    btSyncNotes.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncNotes"));

            try
            {
                if (regKeyAppRoot.GetValue("CalendarFolderID") != null)
                    this.SetCalendarFolder((string)regKeyAppRoot.GetValue("CalendarFolderID"));
            }
            catch (COMException exc)
            {
                if (exc.ErrorCode == Properties.Settings.Default.Constant_ID_not_found)
                    Logger.Log(Properties.Resources.Warning_SavedFolderIdNotFound, EventType.Warning);
            }

            OutlookConnection.Disconnect();
            //autoSyncCheckBox_CheckedChanged(null, null);

            //if (!Properties.Settings.Default.ApplicationAllowedToRun)
            //{
            //    btSyncCalendar.Checked = false;
            //    btSyncCalendar.Enabled = false;
            //}
		}

        private void SetCalendarFolder(string calendarFolderID)
        {
            this._calendarFolderToSyncID = calendarFolderID;
            this.lnkCalendarFolder.Text = GetOutlookFolderPath(calendarFolderID);
            this.Synchronizer.CalendarFolderToSyncID = calendarFolderID;
        }

        private string GetOutlookFolderPath(string folderID)
        {
            Outlook.MAPIFolder currFolder = OutlookConnection.Namespace.GetFolderFromID(folderID);
            var path = new StringBuilder();
            while (currFolder.Parent as Outlook.MAPIFolder != null)
            {
                path.Insert(0, "\\" + currFolder.Name);
                currFolder = (Outlook.MAPIFolder)currFolder.Parent;
            }
            return path.ToString();
        }

        private void SaveProxySettings()
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
            regKeyAppRoot.SetValue("ProxyUsage", this.CustomProxy.Checked);

            if (this.CustomProxy.Checked)
            {

                if (!string.IsNullOrEmpty(this.Address.Text))
                {
                    regKeyAppRoot.SetValue("ProxyURL", this.Address.Text);
                    if (!string.IsNullOrEmpty(Port.Text))
                        regKeyAppRoot.SetValue("ProxyPort", this.Port.Text);
                }

                regKeyAppRoot.SetValue("ProxyAuth", this.Authorization.Checked);
                if (this.Authorization.Checked)
                {
                    if (!string.IsNullOrEmpty(this.txtProxyUserName.Text))
                    {
                        regKeyAppRoot.SetValue("ProxyUsername", this.txtProxyUserName.Text);
                        if (!string.IsNullOrEmpty(this.txtProxyPassword.Text))
                            regKeyAppRoot.SetValue("ProxyPassword", Encryption.EncryptPassword(this.txtProxyUserName.Text, this.txtProxyPassword.Text));
                    }
                }
            }
        }

		private void SaveSettings()
		{
			RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
            regKeyAppRoot.SetValue("Language", ((CultureInfo)this.cmbLanguage.SelectedItem).LCID, RegistryValueKind.DWord);
			regKeyAppRoot.SetValue("SyncOption", (int)_syncOption);
            if (!string.IsNullOrEmpty(this._googleAccountName))
            {
                regKeyAppRoot.SetValue("GoogleAccountName", this._googleAccountName);
            }
            regKeyAppRoot.SetValue("AutoSync", this.mniAutoSync.Checked);
			regKeyAppRoot.SetValue("AutoSyncInterval", autoSyncInterval.Value.ToString());
			regKeyAppRoot.SetValue("ReportSyncResult", reportSyncResultCheckBox.Checked);
            regKeyAppRoot.SetValue("SyncCalendar", btSyncCalendar.Checked);
            regKeyAppRoot.SetValue("SyncContacts", btSyncContacts.Checked);
            //regKeyAppRoot.SetValue("SyncNotes", btSyncNotes.Checked);
            regKeyAppRoot.SetValue("CalendarFolderID", this._calendarFolderToSyncID);
            regKeyAppRoot.SetValue("OutlookProfile", this.cmbOutlookProfiles.SelectedItem);
		}

        private void SetProxy()
        {
            if (this.CustomProxy.Checked)
            {
                try
                {
                    var myProxy = new WebProxy(this.Address.Text, Convert.ToInt16(Port.Text))
                    {
                        BypassProxyOnLocal = true
                    };

                    if (this.Authorization.Checked)
                    {
                        myProxy.Credentials = new NetworkCredential(this.txtProxyUserName.Text, this.txtProxyPassword.Text);
                    }
                    WebRequest.DefaultWebProxy = myProxy;
                }
                catch (Exception ex)
                {
                    ErrorHandler.Handle(ex);
                }
            }
            else // to do set default system proxy
                WebRequest.DefaultWebProxy = WebRequest.GetSystemWebProxy();
        }

		private bool ValidCredentials
		{
			get
			{
				//bool userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\'\%\._\+\-]+)@(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);
				bool syncProfileNameIsValid = cmbOutlookProfiles.Text.Length != 0;

				//setBgColor(UserName, userNameIsValid);
				SetBgColor(cmbOutlookProfiles, syncProfileNameIsValid);
				return /*userNameIsValid && passwordIsValid &&*/ syncProfileNameIsValid;
			}
		}
		private void SetBgColor(Control box, bool isValid)
		{
			if (!isValid)
				box.BackColor = Color.LightPink;
			else
				box.BackColor = Color.White;
		}

		private void Sync()
		{
			try
			{
				if (!ValidCredentials)
					return;

                ThreadStart starter = new ThreadStart(Sync_ThreadStarter);
                this._syncThread = new Thread(starter);
                this._syncThread.SetApartmentState(ApartmentState.STA);
                this._syncThread.Start();

                /// wait for thread to start
                while (!this._syncThread.IsAlive)
                    Thread.Sleep(500);

                lastSync = DateTime.Now;
                SetLastSyncText("Last synced at " + lastSync.ToString());
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}

        [STAThread]
		private void Sync_ThreadStarter()
		{
            try
            {
                //var app = new Microsoft.Office.Interop.Outlook.Application();
                //var ns = app.GetNamespace("MAPI");
                //ns.Logon();
                //var sm = new AddinExpress.Outlook.SecurityManager();
                //sm.ConnectTo(app.Application);

                TimerSwitch(false);
                SetLastSyncText("Syncing...");
                notifyIcon.Text = Application.ProductName + "\nSyncing...";
                SetFormEnabled(false);

                Logger.ClearLog();
                SetSyncConsoleText("");
                Logger.Log("Sync started.", EventType.Information);
                //SetSyncConsoleText(Logger.GetText());
                //this.Synchronizer.SyncProfile = this.tbSyncProfile.Text;
                this.Synchronizer.SyncOption = this._syncOption;
                /// So far I prefer synchronization settings to be available throughout application
                Properties.Settings.Default.SynchronizationOption = this._syncOption;
                //_sync.SyncDelete = true;
                this.Synchronizer.SyncCalendar = this.btSyncCalendar.Checked;
                this.Synchronizer.SyncContacts = this.btSyncContacts.Checked;
                //_sync.SyncNotes = false;



                this.LoginToGoogle();
                //_sync.LoginToOutlook();

                this.Synchronizer.Sync();

                if (reportSyncResultCheckBox.Checked)
                {
                    /*
                    notifyIcon.BalloonTipTitle = Application.ProductName;
                    notifyIcon.BalloonTipText = string.Format("{0}. {1}", DateTime.Now, message);
                    */
                    ToolTipIcon icon;
                    if (this.Synchronizer.Result.ErrorItems > 0)
                        icon = ToolTipIcon.Error;
                    else if (this.Synchronizer.SkippedCount > 0)
                        icon = ToolTipIcon.Warning;
                    else
                        icon = ToolTipIcon.Info;
                    /*notifyIcon.ShowBalloonTip(5000);
                    */
                    ShowBalloonToolTip(Application.ProductName,
                        string.Format("{0}. {1}\r\n{2}", DateTime.Now, Properties.Resources.Notification_SyncComplete, this.Synchronizer.Result),
                        icon,
                        5000);

                }
                string toolTip = string.Format("{0}\nLast sync: {1}", Application.ProductName, DateTime.Now.ToString("dd.MM. HH:mm"));
                if (this.Synchronizer.Result.ErrorItems + this.Synchronizer.SkippedCount > 0)
                    toolTip += string.Format("\nWarnings: {0}.", this.Synchronizer.Result.ErrorItems + this.Synchronizer.SkippedCount);
                if (toolTip.Length >= 64)
                    toolTip = toolTip.Substring(0, 63);
                notifyIcon.Text = toolTip;
            }
            catch (WebException exc)
            {
                ErrorHandler.Handle(exc);
                Logger.Log(String.Format("{2}\r\n{0}\r\n{1}", exc.Message, exc.Response, Properties.Resources.Error_ConnectionFailure), EventType.Error);
                SetLastSyncText(Properties.Resources.Notification_SyncFailed);
            }
            catch (GoogleConnectionException)
            {
                string message = Properties.Resources.Error_ConnectionFailure;
                Logger.Log(message, EventType.Error);
                Program.Instance.ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000);
            }
            catch (NoDataToSyncronizeSpecifiedException)
            {
                MessageBox.Show(Properties.Resources.Error_NoDataToSyncronizeSpecified, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (ThreadAbortException) { }
            /// This very likely means the application is being closed
            catch (ObjectDisposedException) { }
            catch (Exception ex)
            {
                /// Some Google exception. Couldn't find analog so far
                //SetLastSyncText(Properties.Resources.Notification_SyncFailed);
                //notifyIcon.Text = Application.ProductName + "\n" + Properties.Resources.Notification_SyncFailed;

                //string responseString = (null != ex.InnerException) ? ex.ResponseString : ex.Message;

                //if (ex.InnerException is System.Net.WebException)
                //{
                //    string message = Properties.Resources.Error_ConnectionFailure + "\n" + ((System.Net.WebException)ex.InnerException).Message + "\r\n" + responseString;
                //    Logger.Log(message, EventType.Warning);
                //    Program.Instance.ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000);
                //}
                //else
                //{
                //    ErrorHandler.Handle(ex);
                //}


                Logger.Log("The error occured, which was not caught by Synchronizer. Probably it has occurred in \"finally\" section", EventType.Debug);
                SetLastSyncText(Properties.Resources.Notification_SyncFailed);
                notifyIcon.Text = Application.ProductName + "\n" + Properties.Resources.Notification_SyncFailed;
                ErrorHandler.Handle(ex);
            }							
			finally
			{                        
                lastSync = DateTime.Now;
                TimerSwitch(this.mniAutoSync.Checked);
				SetFormEnabled(true);
                if (this.Synchronizer != null)
                {
                    //_sync.LogoffOutlook();
                    this.Synchronizer.LogoffGoogle();
                    GC.Collect();
                }
			}
            if (!this.Synchronizer.Result.Succeeded)
            {
                if (MessageBox.Show(
                                    Properties.Resources.Confirm_SendSessionLog,
                                    Application.ProductName,
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question,
                                    MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes)
                {
                    Logger.SendSessionLog((sendResult) => NotifyLogSendResult(sendResult));
                }
            }
		}

        public void ShowBalloonToolTip(string title, string message, ToolTipIcon icon, int timeout)
        {
            if (this.InvokeRequired)
                this.Invoke((MethodInvoker)(() => this.ShowBalloonToolTip(title, message, icon, timeout)));
            else
            {
                //if user is active on workstation
                if (boolShowBalloonTip)
                {
                    notifyIcon.BalloonTipTitle = title;
                    notifyIcon.BalloonTipText = message;
                    notifyIcon.BalloonTipIcon = icon;
                    notifyIcon.ShowBalloonTip(timeout);
                }
            }
        }

		void Logger_LogUpdated(string Message)
		{
			AppendSyncConsoleText(Message);
		}

		void OnErrorEncountered(string title, Exception ex, EventType eventType)
		{
            Logger.SendSessionLog((sendResult) => NotifyLogSendResult(sendResult));
		}

		void OnDuplicatesFound(string title, string message)
		{
            Logger.Log(message, EventType.Warning);
            ShowBalloonToolTip(title,message,ToolTipIcon.Warning,5000);
            /*
			notifyIcon.BalloonTipTitle = title;
			notifyIcon.BalloonTipText = message;
			notifyIcon.BalloonTipIcon = ToolTipIcon.Warning;
			notifyIcon.ShowBalloonTip(5000);
             */
		}

        void OnNotificationReceived(string message)
        {
            SetLastSyncText(message);           
        }

		public void SetFormEnabled(bool enabled)
		{
			if (this.InvokeRequired)
			{
                this.Invoke((MethodInvoker)delegate { SetFormEnabled(enabled); });
			}
			else
			{
				resetMatchesLinkLabel.Enabled = enabled;
				//settingsGroupBox.Enabled = enabled;
				this.mniSync.Enabled = enabled;
                this.mniStopSync.Enabled = !enabled;
			}
		}
		public void SetLastSyncText(string text)
		{
			if (this.InvokeRequired)
			{
				TextHandler h = new TextHandler(SetLastSyncText);
				this.Invoke(h, new object[] { text });
			}
			else
				this._log.LastSyncLabel.Text = text;
		}
		public void SetSyncConsoleText(string text)
		{
			if (this.InvokeRequired)
			{
				TextHandler h = new TextHandler(SetSyncConsoleText);
				this.Invoke(h, new object[] { text });
			}
			else
            {
				this._log.SyncConsole.Text = text;
                //Scroll to bottom to always see the last log entry
                this._log.SyncConsole.SelectionStart = this._log.SyncConsole.TextLength;
                this._log.SyncConsole.ScrollToCaret();
            }

		}
		public void AppendSyncConsoleText(string text)
		{
			if (this.InvokeRequired)
			{
				TextHandler h = new TextHandler(AppendSyncConsoleText);
				this.Invoke(h, new object[] { text });
			}
			else
            {
				this._log.SyncConsole.Text += text;
                //Scroll to bottom to always see the last log entry
                this._log.SyncConsole.SelectionStart = this._log.SyncConsole.TextLength;
                this._log.SyncConsole.ScrollToCaret();
            }
		}

		public void TimerSwitch(bool value)
		{
			if (this.InvokeRequired)
			{
                this.Invoke((MethodInvoker)delegate { TimerSwitch(value); });
			}
			else
			{
                this.grpAutoSync.Enabled = value;
                this.syncTimer.Enabled = value;
				this.nextSyncLabel.Visible = value;
			}
		}

        //to detect if the user locks or unlocks the workstation
        [DllImport("wtsapi32.dll")]
        private static extern bool WTSRegisterSessionNotification(IntPtr hWnd, int dwFlags);

        [DllImport("wtsapi32.dll")]
        private static extern bool WTSUnRegisterSessionNotification(IntPtr hWnd);

		// Fix for WinXP and older systems, that do not continue with shutdown until all programs have closed
		// FormClosing would hold system shutdown, when it sets the cancel to true
		private const int WM_QUERYENDSESSION = 0x11;

        //Code to find out if workstation is locked
        private const int WM_WTSSESSION_CHANGE = 0x02B1;
        private const int WTS_SESSION_LOCK = 0x7;
        private const int WTS_SESSION_UNLOCK = 0x8;
        private string _googleAccountName;

        /*
        protected void OnSessionLock()
        {
            Logger.Log("Locked at " + DateTime.Now + Environment.NewLine, EventType.Information);
        }

        protected void OnSessionUnlock()
        {
            Logger.Log("Unlocked at " + DateTime.Now + Environment.NewLine, EventType.Information);
        }
        */

        protected override void WndProc(ref System.Windows.Forms.Message m)
		{
            //Logger.Log(m.Msg, EventType.Information);
            switch(m.Msg) 
            {
                
                case WM_QUERYENDSESSION:
                    requestClose = true;
                    break;
                case WM_WTSSESSION_CHANGE:
                    {
                        if (m.WParam.ToInt32() == WTS_SESSION_LOCK)
                        {
                            //Logger.Log("\nBenutzer aktiv -> ToolTip", EventType.Information);
                            //OnSessionLock();
                            boolShowBalloonTip = false; // Do something when locked
                        }
                        else if (m.WParam.ToInt32() == WTS_SESSION_UNLOCK)
                        {
                            //Logger.Log("\nBenutzer inaktiv -> kein ToolTip", EventType.Information);
                            //OnSessionUnlock();
                            boolShowBalloonTip = true; // Do something when unlocked
                        }
                     break;
                    }
                default:
                    break;
            }
            /*
			if (m.Msg == WM_QUERYENDSESSION)
				requestClose = true;
            if (m.Msg == SESSIONCHANGEMESSAGE)
            {
                if (m.WParam.ToInt32() == SESSIONLOCKPARAM)
                    OnSessionLock(); // Do something when locked
                else if (m.WParam.ToInt32() == SESSIONUNLOCKPARAM)
                    OnSessionUnlock(); // Do something when unlocked
            }*/
			// If this is WM_QUERYENDSESSION, the form must exit and not just hide
			base.WndProc(ref m);
		} 

		private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (!requestClose)
			{
                this.StopSynchronization();
				SaveSettings();
				e.Cancel = true;
			}
		}
		private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
		{
			try
			{
				SaveSettings();

				notifyIcon.Dispose();
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}

		private void syncOptionBox_SelectedIndexChanged(object sender, EventArgs e)
		{
			try
			{
				int index = syncOptionBox.SelectedIndex;
				if (index == -1)
					return;

				SetSyncOption(index);
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}
		private void SetSyncOption(int index)
		{
			_syncOption = (SyncOption)index;
			for (int i = 0; i < syncOptionBox.Items.Count; i++)
			{
				if (i == index)
					syncOptionBox.SetItemCheckState(i, CheckState.Checked);
				else
					syncOptionBox.SetItemCheckState(i, CheckState.Unchecked);
			}
		}

		private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
		{
            if ((this._syncThread != null) && this._syncThread.IsAlive)
                this.mniLog_Click(sender, e);
            else
                this.mniSettings_Click(sender, e);
		}

        //private void autoSyncCheckBox_CheckedChanged(object sender, EventArgs e)
        //{
        //    lastSync = DateTime.Now.AddSeconds(15) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
        //    autoSyncInterval.Enabled = autoSyncCheckBox.Checked;
        //    syncTimer.Enabled = autoSyncCheckBox.Checked;
        //    nextSyncLabel.Visible = autoSyncCheckBox.Checked;
        //}

		private void syncTimer_Tick(object sender, EventArgs e)
		{
            //Logger.Log("Autosync timer tick", EventType.Information);
			if (lastSync != null)
			{
				TimeSpan syncTime = DateTime.Now - lastSync;
				TimeSpan limit = new TimeSpan(0, (int)autoSyncInterval.Value, 0);
				if (syncTime < limit)
				{
					TimeSpan diff = limit - syncTime;
					string str = "Next sync in";
					if (diff.Hours != 0)
						str += " " + diff.Hours + " h";
					if (diff.Minutes != 0 || diff.Hours != 0)
						str += " " + diff.Minutes + " min";
					if (diff.Seconds != 0)
						str += " " + diff.Seconds + " s";
					nextSyncLabel.Text = str;
					return;
				}
			}
			Sync();
		}

		private void resetMatchesLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
			// force deactivation to show up
			Application.DoEvents();
			try
			{
                TimerSwitch(false);
                SetLastSyncText("Resetting matches...");
                notifyIcon.Text = Application.ProductName + "\nResetting matches...";
                SetFormEnabled(false);
                this.hideButton.Enabled = false;

                Logger.ClearLog();
                SetSyncConsoleText("");
                Logger.Log("Reset Matches started.", EventType.Information);

                //_sync.SyncNotes = false;
                this.Synchronizer.SyncContacts = btSyncContacts.Checked;
                this.Synchronizer.SyncCalendar = btSyncCalendar.Checked;

				this.LoginToGoogle();
				//this.Synchronizer.LoginToOutlook();
                //this.Synchronizer.SyncProfile = tbSyncProfile.Text;

                this.Synchronizer.Unpair();



                lastSync = DateTime.Now;
                SetLastSyncText("Matches reset at " + lastSync.ToString());
                Logger.Log("Matches reset.", EventType.Information);                
			}
			catch (Exception ex)
            {
                SetLastSyncText("Reset Matches failed");
                Logger.Log("Reset Matches failed", EventType.Error);
				ErrorHandler.Handle(ex);
			}
			finally
			{                
                lastSync = DateTime.Now;
                TimerSwitch(true);
				SetFormEnabled(true);
                this.hideButton.Enabled = true;
                if (this.Synchronizer != null)
                {
                    //this.Synchronizer.LogoffOutlook();
                    this.Synchronizer.LogoffGoogle();
                }
			}
		}

        private void ShowForm()
        {
            WindowState = FormWindowState.Normal;
            Show();
            this.Activate();
        }
		private void HideForm()
		{
			WindowState = FormWindowState.Minimized;
			Hide();
		}

		private void mniSettings_Click(object sender, EventArgs e)
		{
            this.SwitchWindow(mniSettings, this);
		}

		private void mniLog_Click(object sender, EventArgs e)
		{
            this.SwitchWindow(mniLog, this._log);
		}

        private void SwitchWindow(ToolStripMenuItem menuItem, Form window)
        {
            if (menuItem.Checked)
                window.Hide();
            else
            {
                window.Show();
                window.BringToFront();
            }
            menuItem.Checked = !menuItem.Checked;
        }

		private void toolStripMenuItem2_Click(object sender, EventArgs e)
		{
            this.Exit();
		}

        private void Exit()
        {
            requestClose = true;
            this.Close();
        }

		private void mniSync_Click(object sender, EventArgs e)
		{
			Sync();
		}

		private void SettingsForm_Load(object sender, EventArgs e)
		{
            //if ((this.CheckUpdates() != UpdateResult.Success) && (this.CheckUpdates() != UpdateResult.UpToDate))
            //    return;

			if (/*string.IsNullOrEmpty(UserName.Text) ||
				string.IsNullOrEmpty(Password.Text) || */
				string.IsNullOrEmpty(cmbOutlookProfiles.Text))
			{
				// this is the first load, show form
				ShowForm();
				//UserName.Focus();
			}
			else
				HideForm();
		}

        //private void runAtStartupCheckBox_CheckedChanged(object sender, EventArgs e)
        //{
        //    RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");

        //    if (runAtStartupCheckBox.Checked)
        //    {
        //        // add to registry
        //        regKeyAppRoot.SetValue("GoogleContactSync", Application.ExecutablePath);
        //    }
        //    else
        //    {
        //        // remove from registry
        //        regKeyAppRoot.DeleteValue("GoogleContactSync");
        //    }
        //}

		private void UserName_TextChanged(object sender, EventArgs e)
		{
			ValidateSyncButton();
		}
		private void Password_TextChanged(object sender, EventArgs e)
		{
			ValidateSyncButton();
		}

		private void ValidateSyncButton()
		{
			this.mniSync.Enabled = ValidCredentials;
		}

		private void deleteDuplicatesButton_Click(object sender, EventArgs e)
		{
			//DeleteDuplicatesForm f = new DeleteDuplicatesForm(_sync
		}

		//private void tbSyncProfile_TextChanged(object sender, EventArgs e)
		//{
		//	ValidateSyncButton();
  //          this.Synchronizer.SyncProfile = this.tbSyncProfile.Text;
		//}        

		private void hideButton_Click(object sender, EventArgs e)
		{
			this.Close();
		}

        private void lnkCalendarOptions_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new CalendarSyncSettings().Show();
        }

        private void SettingsForm_HelpRequested(object sender, EventArgs e)
        {
            new AboutBox().ShowDialog();
        }

        private void PageToolButton_Click(object sender, EventArgs e)
        {
            ToolStripButton button = (ToolStripButton)sender;
            SetActivePanel((Panel)button.Tag);
            tbtnAdvanced.Checked = tbtnAdvanced == button;
            tbtnGeneral.Checked = tbtnGeneral == button;
            tbtnProxy.Checked = tbtnProxy == button;
        }

        private void SetActivePanel(Panel panel)
        {
            pnlAdvanced.Visible = panel == pnlAdvanced;
            pnlGeneral.Visible = panel == pnlGeneral;
            pnlProxy.Visible = panel == pnlProxy;
        }

        private bool ValidateProxySettings()
        {
            bool userNameIsValid = Regex.IsMatch(this.txtProxyUserName.Text, @"^(?'id'[a-z0-9\\\/\@\'\%\._\+\- ]+)$", RegexOptions.IgnoreCase);
            bool passwordIsValid = this.txtProxyPassword.Text.Length != 0;
            bool AddressIsValid = Regex.IsMatch(this.Address.Text, @"^(?'url'[\w\d#@%;$()~_?\\\.&]+)$", RegexOptions.IgnoreCase);
            bool PortIsValid = Regex.IsMatch(this.Port.Text, @"^(?'port'[0-9]{2,6})$", RegexOptions.IgnoreCase);

            SetBgColor(this.txtProxyUserName, userNameIsValid);
            SetBgColor(this.txtProxyPassword, passwordIsValid);
            SetBgColor(this.Address, AddressIsValid);
            SetBgColor(this.Port, PortIsValid);
            return (userNameIsValid && passwordIsValid || !Authorization.Checked) && AddressIsValid && PortIsValid || SystemProxy.Checked;
        }

        private void CustomProxy_Changed(object sender, EventArgs e)
        {
            this.Address.Enabled = this.CustomProxy.Checked;
            this.Port.Enabled = this.CustomProxy.Checked;
            this.Authorization.Enabled = this.CustomProxy.Checked;
            this.txtProxyUserName.Enabled = this.CustomProxy.Checked && this.Authorization.Checked;
            this.txtProxyPassword.Enabled = this.CustomProxy.Checked && this.Authorization.Checked;

            if (this.ValidateProxySettings())
                this.SetProxy();
        }

        private void Authorization_CheckedChanged(object sender, EventArgs e)
        {
            this.txtProxyUserName.Enabled = this.Authorization.Checked;
            this.txtProxyPassword.Enabled = this.Authorization.Checked;
        }

        private void pnlHelp_Click(object sender, EventArgs e)
        {
            Process.Start(Properties.Settings.Default.URL_Help);
        }

        private void mniStopSync_Click(object sender, EventArgs e)
        {

            if (MessageBox.Show(
                Properties.Resources.Confirm_StopSyncronization,
                Application.ProductName,
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                this.StopSynchronization();
            }
        }

        private void StopSynchronization()
        {
            if ((this._syncThread != null) && this._syncThread.IsAlive)
                this._syncThread.Abort();
        }

        private void lnkBuyNow_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(Properties.Settings.Default.URL_Register);
        }

        private void mniAutoSync_Click(object sender, EventArgs e)
        {
            this.SetAutoSync(!this.mniAutoSync.Checked);
        }

        private void SetAutoSync(bool enableAutoSync)
        {
            this.mniAutoSync.Checked = enableAutoSync;
            this.TimerSwitch(enableAutoSync);
       }

        private void mniRegister_Click(object sender, EventArgs e)
        {
            new RegisterForm().ShowDialog();
        }

        private void lnkCheckUpdates_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //var result = this.CheckUpdates();
            //switch (result)
            //{
            //    case UpdateResult.Fail:
            //        Logger.Log(Properties.Resources.Notification_UpdateFailure, EventType.Warning);
            //        break;
            //    case UpdateResult.Restart:
            //        Logger.Log(Properties.Resources.Notification_UpdateRestartNeeded, EventType.Information);
            //        break;
            //    case UpdateResult.Success:
            //        Logger.Log(Properties.Resources.Notification_UpdateSuccess, EventType.Information);
            //        break;
            //    case UpdateResult.UpToDate:
            //        MessageBox.Show(Properties.Resources.Notification_UpdateUpToDate, Application.ProductName);
            //        break;
            //}
        }

        private void cmbLanguage_SelectedValueChanged(object sender, EventArgs e)
        {
            //this.SaveSettings();
            if (Thread.CurrentThread.CurrentUICulture != cmbLanguage.SelectedItem) 
                this.lblLanguageChangeNotification.Visible = true;
        }

        private void lnkSendSessionLog_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new Thread(() => Logger.SendSessionLog(sendResult => NotifyLogSendResult(sendResult))).Start();
        }

        private void lnkSendLogFile_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new Thread(() => Logger.SendLogFile(sendResult => NotifyLogSendResult(sendResult))).Start();
        }

        private void NotifyLogSendResult(bool successful)
        {
            if (successful)
                this.ShowBalloonToolTip(Application.ProductName, Properties.Resources.Notification_LogSent, ToolTipIcon.Info, 5000);
            else
                this.ShowBalloonToolTip(Application.ProductName, Properties.Resources.Error_LogSendingFailure, ToolTipIcon.Error, 5000);
        }

        private void lnkCalendarFolder_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.ChooseCalendarFolder();
        }

        private void ChooseCalendarFolder()
        {
            try { OutlookConnection.Connect(); }
            catch (OutlookConnectionException) { return; }
            Outlook.MAPIFolder chosenFolder = OutlookConnection.Namespace.PickFolder();
            if (chosenFolder != null)
                this.SetCalendarFolder(chosenFolder.EntryID);

            OutlookConnection.Disconnect();
        }

        private void SetOutlookProfile(string profile)
        {
            OutlookConnection.Profile = profile;
            if (this.cmbOutlookProfiles.Text != profile)
                this.cmbOutlookProfiles.SelectedItem = profile;
        }

        private void cmbOutlookProfiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.SetOutlookProfile(cmbOutlookProfiles.Text);
        }

        private void lnkGoogleLogon_Click(object sender, EventArgs e)
        {
            this.LoginToGoogle();
        }

        private void LoginToGoogle()
        {
            if (this.IsGoogleAccountValid())
            {
                this.Synchronizer.LoginToGoogle(this._googleAccountName);
            }
        }

        private bool IsGoogleAccountValid()
        {
            return 
                !string.IsNullOrEmpty(this._googleAccountName) && 
                (Regex.IsMatch(this._googleAccountName, @"[\w!#$%&'*+\-\/=?\^_`{|}~][\w!#$%&'*+\-\/=?\^_`{|}~\.]*[\w!#$%&'*+\-\/=?\^_`{|}~]@[\w!#$%&'*+\-\/=?\^_`{|}~][\w!#$%&'*+\-\/=?\^_`{|}~\.]*"));
        }

        private void txtGoogleAccountName_TextChanged(object sender, EventArgs e)
        {
            this._googleAccountName = this.txtGoogleAccountName.Text;
            if (IsGoogleAccountValid())
            {
                SetBgColor(txtGoogleAccountName, true);
            }
            else
            {
                SetBgColor(txtGoogleAccountName, false);
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

namespace R.GoogleOutlookSync
{
    class OutlookConnection : IDisposable
    {
        private Microsoft.Office.Interop.Outlook.Application _application = null;
        private Microsoft.Office.Interop.Outlook.NameSpace _namespace = null;
        private static OutlookConnection _instance;

        public static Application Application
        {
            get
            {
                if (_instance == null)
                    throw new System.Exception("Outlook connection is not initialized. Initialise first with OutlookConnection.Connect() function");
                return _instance._application;
            }
        }

        public static NameSpace Namespace
        {
            get
            {
                if (_instance == null)
                    throw new System.Exception("Outlook connection is not initialized. Initialise first with OutlookConnection.Connect() function");
                var attemptsAmount = Properties.Settings.Default.AttemptsAmount;
                return _instance._namespace;
            }
        }

        public static string Profile { get; set; }

        private OutlookConnection()
        {
            for (int attemptsLeft = Properties.Settings.Default.AttemptsAmount; attemptsLeft > 0; attemptsLeft--)
            {
                try
                {
                    /// First try to get running Outlook instance
                    try { this._application = (Application)Marshal.GetActiveObject("Outlook.Application"); }
                    /// If it fails - try to create new object
                    catch (COMException) { this._application = new Microsoft.Office.Interop.Outlook.Application(); }
                    this._namespace = this._application.GetNamespace("MAPI");
                    this._namespace.Logon(Profile);
                    var version = this._application.Version;
                    Properties.Settings.Default.OutlookVersion = Convert.ToInt32(version.Substring(0, version.IndexOf('.')));
                    break;
                }
                catch (COMException exc)
                {
                    if (exc.ErrorCode == Properties.Settings.Default.Constant_RPC_server_is_not_available)
                    {
                        Logger.Log("Couldn't connect to Outlook. Error: " + ErrorHandler.BuildExceptionDescription(exc), EventType.Debug);
                        System.Threading.Thread.Sleep(10000);
                        attemptsLeft--;
                    }
                    else if (exc.ErrorCode == Properties.Settings.Default.Constant_Profile_is_not_set_up)
                    {
                        throw new ProfileNotConfiguredException(Profile);
                    }
                    else
                        throw exc;
                }
            }
        }

        public static void Connect()
        {
            Logger.Log("Connecting to Outlook", EventType.Information);
            try
            { _instance = new OutlookConnection(); }
            catch (System.Exception exc)
            {
                ErrorHandler.Handle(exc);
                throw new OutlookConnectionException(exc);
            }
        }

        public static void Disconnect()
        {
            Logger.Log("Disconnecting from Outlook", EventType.Debug);
            _instance.Dispose();
            _instance = null;
        }
    
        public void  Dispose()
        {
            if (this._namespace != null)
            {
                this._namespace.Logoff();
                Marshal.ReleaseComObject(this._namespace);
                this._namespace = null;
            }
            if (this._application != null)
            {
                Marshal.ReleaseComObject(this._application);
                this._application = null;
            }
        }
    }
}

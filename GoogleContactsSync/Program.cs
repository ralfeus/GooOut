using System;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Threading;
using System.Diagnostics;
using System.Globalization;


namespace R.GoogleOutlookSync
{
    static class Program
    {
		private static SettingsForm instance;
        internal static Mutex ProgramMutex;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
#if DEBUG
            /// Thus I can attach to the process at the remote machine
            //MessageBox.Show("Press OK to continue");
#endif
            SetLanguage();
            //prevent more than one instance of the program
            /// Old way
            //bool ok;
            //ProgramMutex = new System.Threading.Mutex(true, "acbbbc09-f76c-4874-aaff-4f3353a5a5a6", out ok);
            //if (!ok)

            ProgramMutex = new Mutex(false, "acbbbc09-f76c-4874-aaff-4f3353a5a5a6");
            if (!ProgramMutex.WaitOne(1000, true))
            {
                MessageBox.Show(String.Format(Properties.Resources.Error_AnotherInstanceIsRunning, Application.ProductName), Application.ProductName, MessageBoxButtons.OK);
                return;
            }

            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //SetRegistrationStatus();
            //Properties.Settings.Default.ApplicationAllowedToRun = !Properties.Settings.Default.IsNotRegistered || !Program.IsExpired();
            //if (!Properties.Settings.Default.ApplicationAllowedToRun)
            //{
            //    MessageBox.Show(Properties.Resources.Warning_ApplicationExpired, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    new TrialityNotifierForm().ShowDialog();
            //}
            
            Application.Run(new SettingsForm());
            if (ProgramMutex.WaitOne(0, true))
                ProgramMutex.ReleaseMutex();
            //GC.KeepAlive(m);
        }

        private static bool RebootedAfterInstallation()
        {
            DateTime installTime;
            RegistryKey key = null;
            try
            {
                key = Registry.CurrentUser.OpenSubKey(Properties.Settings.Default.ApplicationRegistryKey, true);
                installTime = DateTime.FromBinary((long)key.GetValue(Properties.Settings.Default.Registry_InstallTimeValue));
            }
            catch (Exception exc)
            {
                Logger.Log(String.Format("Couldn't get install time. Error is '{0}'", exc.Message), EventType.Debug);
                return true;
            }
            using (var uptime = new PerformanceCounter("System", "System Up Time"))
            {
                uptime.NextValue();       //Call this an extra time before reading its value
                if (DateTime.Now - TimeSpan.FromSeconds(uptime.NextValue()) > installTime)
                {
                    key.DeleteValue(Properties.Settings.Default.Registry_InstallTimeValue);
                    return true;
                }
                else
                    return false;
            }
        }

		internal static SettingsForm Instance
		{
			get { return instance; }
		}

        /// <summary>
        /// Fallback. If there is some try/catch missing we will handle it here, just before the application quits unhandled
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception)
                ErrorHandler.Handle((Exception)e.ExceptionObject);
        }

        private static bool IsExpired()
        {
            var key = Registry.CurrentUser.OpenSubKey(Properties.Settings.Default.ApplicationRegistryKey, true);
            if (key == null)
                key = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
            var expirationDate = key.GetValue(Properties.Settings.Default.RegistrationSettings_ExpirationDate);
            if (expirationDate == null)
            {
                expirationDate = DateTime.Today.AddDays(Properties.Settings.Default.RegistrationSettings_TrialityDuration).ToBinary() ^ Properties.Settings.Default.CPUSerial;
                key.SetValue(Properties.Settings.Default.RegistrationSettings_ExpirationDate, expirationDate, RegistryValueKind.QWord);
                return false;
            }
            else
            {
                var startDate = (long)expirationDate ^ Properties.Settings.Default.CPUSerial;
                return DateTime.Today >= DateTime.FromBinary(startDate);
            }
        }

        private static void SetRegistrationStatus()
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(Properties.Settings.Default.ApplicationRegistryKey);
            try
            {
                var storedKey = key.GetValue(Properties.Settings.Default.RegistrationSettings_RegistrationNumberKey);
                var number = (long)storedKey ^ Properties.Settings.Default.CPUSerial;
                Properties.Settings.Default.IsNotRegistered = number % 3 != 0;
            }
            /// If any operation fails it means the key is either not created or is wrong. Both mean the registration is failed
            catch (Exception)
            {
                Properties.Settings.Default.IsNotRegistered = true;
            }
        }

        private static void SetLanguage()
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
            if (regKeyAppRoot.GetValue("Language") != null)
                Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo((int)regKeyAppRoot.GetValue("Language"));
        }
    }
}
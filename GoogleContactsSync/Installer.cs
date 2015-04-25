using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;
using System.Net;


namespace R.GoogleOutlookSync
{
    [RunInstaller(true)]
    public partial class Installer : System.Configuration.Install.Installer
    {
        public Installer()
        {
#if DEBUG
            MessageBox.Show("!");
#endif
            InitializeComponent();

            this.Committed += new InstallEventHandler(Installer_Committed);
        }

        void Installer_Committed(object sender, InstallEventArgs e)
        {
#if DEBUG
            MessageBox.Show("!!");
#endif
            /// Perform user defined actions
            if (this.Context.Parameters["chkautostart"] != "1")
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\GooOut.lnk");
            if (this.Context.Parameters["chkrunnow"] == "1")
                Process.Start(this.Context.Parameters["assemblypath"]);

            /// Perform actions dependent of bitness of target environment
            //File.Delete("secman64.dll");

            /// Set install time to compare with reboot time
            var key = Registry.CurrentUser.OpenSubKey(Properties.Settings.Default.ApplicationRegistryKey, true);
            if (key == null)
                key = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
            key.SetValue(Properties.Settings.Default.Registry_InstallTimeValue, DateTime.Now.ToBinary(), RegistryValueKind.QWord);

            this.CleanupOldInstallation();
            this.SendInstallNotification();
        }

        private void CleanupOldInstallation()
        {
            Utilities.TryDeleteFile(Path.GetDirectoryName(this.Context.Parameters["assemblypath"]) + "\\secman.dll");
            Utilities.TryDeleteFile(Path.GetDirectoryName(this.Context.Parameters["assemblypath"]) + "\\secman64.dll");
        }

        private void SendInstallNotification()
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Properties.Settings.Default.URL_InstallNotification);
            for (int i = 0; i < 2; i++)
            {            
                try
                {
                    request.BeginGetResponse(null, null);
                    break;
                }
                catch (Exception) 
                { 
                    System.Threading.Thread.Sleep(5000);
                }            
            }
        }
    }
}

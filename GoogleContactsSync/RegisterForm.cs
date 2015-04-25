using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Security.Cryptography;
using System.Diagnostics;

namespace R.GoogleOutlookSync
{
    public partial class RegisterForm : Form
    {
        public RegisterForm()
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            var key = Registry.CurrentUser.OpenSubKey(Properties.Settings.Default.ApplicationRegistryKey, true);
            if (key == null)
                key = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
            var parts = this.txtRegistrationNumber.Text.Split('-');
            try
            {
                var number = int.Parse(parts[3]) + int.Parse(parts[4]) + int.Parse(parts[5]);
                if (number % 3 != 0)
                    throw new Exception();
                key.SetValue(Properties.Settings.Default.RegistrationSettings_RegistrationNumberKey, number ^ Properties.Settings.Default.CPUSerial, RegistryValueKind.QWord);
                MessageBox.Show(Properties.Resources.RegistrationSettings_RegistrationNotificationText);
            }
            catch (Exception)
            {
                MessageBox.Show(Properties.Resources.Error_RegistrationFailure);
            }
        }

        private void lnkGetRegNumber_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(Properties.Settings.Default.URL_Register);
        }
    }
}

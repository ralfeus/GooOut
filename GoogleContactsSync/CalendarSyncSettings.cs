using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace R.GoogleOutlookSync
{
    public partial class CalendarSyncSettings : Form
    {
        public CalendarSyncSettings()
        {
            InitializeComponent();
        }

        private void CalendarSyncSettings_Deactivate(object sender, EventArgs e)
        {
            this.SaveSettings();
            this.Close();
        }

        private void CalendarSyncSettings_Activated(object sender, EventArgs e)
        {
            this.LoadSettings();
            this.Location = Cursor.Position;
        }

        private void LoadSettings()
        {
            try
            {
                var regKeyAppRoot = Registry.CurrentUser.OpenSubKey(Properties.Settings.Default.ApplicationRegistryKey);
                if (regKeyAppRoot == null)
                    throw new Exception();
                this.nudSyncDaysBefore.Value = (int)regKeyAppRoot.GetValue(Properties.Settings.Default.CalendarSettings_RangeBefore);
                this.nudSyncDaysAfter.Value = (int)regKeyAppRoot.GetValue(Properties.Settings.Default.CalendarSettings_RangeAfter);
                this.txtCalendarName.Text = (string)regKeyAppRoot.GetValue(Properties.Settings.Default.CalendarSettings_CalendarName);
                var test = regKeyAppRoot.GetValue(Properties.Settings.Default.CalendarSettings_CalendarName);
            }
            catch (Exception)
            {
                // If registry key can't be opened due to any reason just use default values
            }
        }

        private void SaveSettings()
        {
            try
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(Properties.Settings.Default.ApplicationRegistryKey);
                regKeyAppRoot.SetValue(Properties.Settings.Default.CalendarSettings_RangeBefore, this.nudSyncDaysBefore.Value, RegistryValueKind.DWord);
                regKeyAppRoot.SetValue(Properties.Settings.Default.CalendarSettings_RangeAfter, this.nudSyncDaysAfter.Value, RegistryValueKind.DWord);
                regKeyAppRoot.SetValue(Properties.Settings.Default.CalendarSettings_CalendarName, this.txtCalendarName.Text);
            }
            catch (Exception exc)
            {
                ErrorHandler.Handle(exc);
            }
        }

        private void CalendarSyncSettings_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void nudSyncDaysAfter_ValueChanged(object sender, EventArgs e)
        {

        }

        private void nudSyncDaysBefore_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}

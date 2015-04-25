using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace R.GoogleOutlookSync
{
    public partial class TrialityNotifierForm : Form
    {
        public TrialityNotifierForm()
        {
            InitializeComponent();
        }

        private void TrialityNotifierForm_Load(object sender, EventArgs e)
        {
            var closeEnabler = new BackgroundWorker();
            closeEnabler.DoWork += new DoWorkEventHandler(closeEnabler_DoWork);
            closeEnabler.RunWorkerAsync();
        }

        void closeEnabler_DoWork(object sender, DoWorkEventArgs e)
        {
            Thread.Sleep(5000);
            if (this.InvokeRequired)
                this.Invoke((System.Action)(() => this.btnClose.Enabled = true));
            else
                this.btnClose.Enabled = true;
            this.FormClosing -= new FormClosingEventHandler(TrialityNotifierForm_FormClosing);
        }

        private void TrialityNotifierForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
        }
    }
}

namespace R.GoogleOutlookSync
{
    partial class LogForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LogForm));
            this.logGroupBox = new System.Windows.Forms.GroupBox();
            this.SyncConsole = new System.Windows.Forms.TextBox();
            this.LastSyncLabel = new System.Windows.Forms.Label();
            this.logGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // logGroupBox
            // 
            resources.ApplyResources(this.logGroupBox, "logGroupBox");
            this.logGroupBox.Controls.Add(this.SyncConsole);
            this.logGroupBox.Controls.Add(this.LastSyncLabel);
            this.logGroupBox.Name = "logGroupBox";
            this.logGroupBox.TabStop = false;
            // 
            // SyncConsole
            // 
            resources.ApplyResources(this.SyncConsole, "SyncConsole");
            this.SyncConsole.BackColor = System.Drawing.SystemColors.Info;
            this.SyncConsole.Name = "SyncConsole";
            this.SyncConsole.ReadOnly = true;
            // 
            // LastSyncLabel
            // 
            resources.ApplyResources(this.LastSyncLabel, "LastSyncLabel");
            this.LastSyncLabel.Name = "LastSyncLabel";
            // 
            // LogForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.logGroupBox);
            this.Name = "LogForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.LogForm_FormClosing);
            this.logGroupBox.ResumeLayout(false);
            this.logGroupBox.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox logGroupBox;
        internal System.Windows.Forms.Label LastSyncLabel;
        internal System.Windows.Forms.TextBox SyncConsole;
    }
}
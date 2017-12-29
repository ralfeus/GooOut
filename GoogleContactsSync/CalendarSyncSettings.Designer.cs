namespace R.GoogleOutlookSync
{
    partial class CalendarSyncSettings
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
            System.Windows.Forms.Label label1;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CalendarSyncSettings));
            System.Windows.Forms.GroupBox groupBox2;
            System.Windows.Forms.Label label3;
            this.txtCalendarName = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.nudSyncDaysAfter = new System.Windows.Forms.NumericUpDown();
            this.nudSyncDaysBefore = new System.Windows.Forms.NumericUpDown();
            label1 = new System.Windows.Forms.Label();
            groupBox2 = new System.Windows.Forms.GroupBox();
            label3 = new System.Windows.Forms.Label();
            groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudSyncDaysAfter)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudSyncDaysBefore)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            resources.ApplyResources(label1, "label1");
            label1.Name = "label1";
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(label3);
            groupBox2.Controls.Add(this.txtCalendarName);
            groupBox2.Controls.Add(this.groupBox1);
            resources.ApplyResources(groupBox2, "groupBox2");
            groupBox2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            groupBox2.Name = "groupBox2";
            groupBox2.TabStop = false;
            // 
            // label3
            // 
            resources.ApplyResources(label3, "label3");
            label3.Name = "label3";
            // 
            // txtCalendarName
            // 
            resources.ApplyResources(this.txtCalendarName, "txtCalendarName");
            this.txtCalendarName.Name = "txtCalendarName";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.nudSyncDaysAfter);
            this.groupBox1.Controls.Add(label1);
            this.groupBox1.Controls.Add(this.nudSyncDaysBefore);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // nudSyncDaysAfter
            // 
            resources.ApplyResources(this.nudSyncDaysAfter, "nudSyncDaysAfter");
            this.nudSyncDaysAfter.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.nudSyncDaysAfter.Name = "nudSyncDaysAfter";
            this.nudSyncDaysAfter.ValueChanged += new System.EventHandler(this.nudSyncDaysAfter_ValueChanged);
            // 
            // nudSyncDaysBefore
            // 
            resources.ApplyResources(this.nudSyncDaysBefore, "nudSyncDaysBefore");
            this.nudSyncDaysBefore.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.nudSyncDaysBefore.Name = "nudSyncDaysBefore";
            this.nudSyncDaysBefore.ValueChanged += new System.EventHandler(this.nudSyncDaysBefore_ValueChanged);
            // 
            // CalendarSyncSettings
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(groupBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.KeyPreview = true;
            this.Name = "CalendarSyncSettings";
            this.Activated += new System.EventHandler(this.CalendarSyncSettings_Activated);
            this.Deactivate += new System.EventHandler(this.CalendarSyncSettings_Deactivate);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.CalendarSyncSettings_KeyDown);
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudSyncDaysAfter)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudSyncDaysBefore)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.NumericUpDown nudSyncDaysBefore;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown nudSyncDaysAfter;
        private System.Windows.Forms.TextBox txtCalendarName;
    }
}
namespace R.GoogleOutlookSync
{
    partial class SettingsForm
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.Label label5;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            System.Windows.Forms.GroupBox grpGoogleSettings;
            System.Windows.Forms.GroupBox grpOutlookSettings;
            System.Windows.Forms.Label label11;
            System.Windows.Forms.Label label10;
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.UserName = new System.Windows.Forms.TextBox();
            this.Password = new System.Windows.Forms.TextBox();
            this.cmbOutlookProfiles = new System.Windows.Forms.ComboBox();
            this.lnkCalendarFolder = new System.Windows.Forms.LinkLabel();
            this.syncOptionBox = new System.Windows.Forms.CheckedListBox();
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.systemTrayMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mniAutoSync = new System.Windows.Forms.ToolStripMenuItem();
            this.mniSync = new System.Windows.Forms.ToolStripMenuItem();
            this.mniStopSync = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.mniSettings = new System.Windows.Forms.ToolStripMenuItem();
            this.mniLog = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.mniRegister = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.autoSyncInterval = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.grpAutoSync = new System.Windows.Forms.GroupBox();
            this.nextSyncLabel = new System.Windows.Forms.Label();
            this.reportSyncResultCheckBox = new System.Windows.Forms.CheckBox();
            this.syncTimer = new System.Windows.Forms.Timer(this.components);
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btSyncCalendar = new System.Windows.Forms.CheckBox();
            this.btSyncContacts = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tbSyncProfile = new System.Windows.Forms.TextBox();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.resetMatchesLinkLabel = new System.Windows.Forms.LinkLabel();
            this.hideButton = new System.Windows.Forms.Button();
            this.tsTabs = new System.Windows.Forms.ToolStrip();
            this.tbtnGeneral = new System.Windows.Forms.ToolStripButton();
            this.tbtnProxy = new System.Windows.Forms.ToolStripButton();
            this.tbtnAdvanced = new System.Windows.Forms.ToolStripButton();
            this.pnlGeneral = new System.Windows.Forms.Panel();
            this.pnlAdvanced = new System.Windows.Forms.Panel();
            this.lnkSendLogFile = new System.Windows.Forms.LinkLabel();
            this.lnkSendSessionLog = new System.Windows.Forms.LinkLabel();
            this.lblLanguageChangeNotification = new System.Windows.Forms.Label();
            this.cmbLanguage = new System.Windows.Forms.ComboBox();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.lnkCheckUpdates = new System.Windows.Forms.LinkLabel();
            this.pnlProxy = new System.Windows.Forms.Panel();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.Port = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label8 = new System.Windows.Forms.Label();
            this.Address = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.Authorization = new System.Windows.Forms.CheckBox();
            this.txtProxyUserName = new System.Windows.Forms.TextBox();
            this.CustomProxy = new System.Windows.Forms.RadioButton();
            this.txtProxyPassword = new System.Windows.Forms.TextBox();
            this.SystemProxy = new System.Windows.Forms.RadioButton();
            this.lnkHelp = new System.Windows.Forms.LinkLabel();
            this.pnlHelp = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lnkBuyNow = new System.Windows.Forms.LinkLabel();
            label5 = new System.Windows.Forms.Label();
            grpGoogleSettings = new System.Windows.Forms.GroupBox();
            grpOutlookSettings = new System.Windows.Forms.GroupBox();
            label11 = new System.Windows.Forms.Label();
            label10 = new System.Windows.Forms.Label();
            grpGoogleSettings.SuspendLayout();
            grpOutlookSettings.SuspendLayout();
            this.systemTrayMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.autoSyncInterval)).BeginInit();
            this.grpAutoSync.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tsTabs.SuspendLayout();
            this.pnlGeneral.SuspendLayout();
            this.pnlAdvanced.SuspendLayout();
            this.pnlProxy.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel2.SuspendLayout();
            this.pnlHelp.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label5
            // 
            resources.ApplyResources(label5, "label5");
            label5.Name = "label5";
            // 
            // grpGoogleSettings
            // 
            resources.ApplyResources(grpGoogleSettings, "grpGoogleSettings");
            grpGoogleSettings.Controls.Add(this.label2);
            grpGoogleSettings.Controls.Add(this.label3);
            grpGoogleSettings.Controls.Add(this.UserName);
            grpGoogleSettings.Controls.Add(this.Password);
            grpGoogleSettings.Name = "grpGoogleSettings";
            grpGoogleSettings.TabStop = false;
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // UserName
            // 
            resources.ApplyResources(this.UserName, "UserName");
            this.UserName.Name = "UserName";
            this.UserName.TextChanged += new System.EventHandler(this.UserName_TextChanged);
            // 
            // Password
            // 
            resources.ApplyResources(this.Password, "Password");
            this.Password.Name = "Password";
            this.Password.TextChanged += new System.EventHandler(this.Password_TextChanged);
            // 
            // grpOutlookSettings
            // 
            grpOutlookSettings.Controls.Add(this.cmbOutlookProfiles);
            grpOutlookSettings.Controls.Add(label11);
            grpOutlookSettings.Controls.Add(this.lnkCalendarFolder);
            grpOutlookSettings.Controls.Add(label10);
            resources.ApplyResources(grpOutlookSettings, "grpOutlookSettings");
            grpOutlookSettings.Name = "grpOutlookSettings";
            grpOutlookSettings.TabStop = false;
            // 
            // cmbOutlookProfiles
            // 
            this.cmbOutlookProfiles.FormattingEnabled = true;
            resources.ApplyResources(this.cmbOutlookProfiles, "cmbOutlookProfiles");
            this.cmbOutlookProfiles.Name = "cmbOutlookProfiles";
            this.cmbOutlookProfiles.Sorted = true;
            this.cmbOutlookProfiles.SelectedIndexChanged += new System.EventHandler(this.cmbOutlookProfiles_SelectedIndexChanged);
            // 
            // label11
            // 
            resources.ApplyResources(label11, "label11");
            label11.Name = "label11";
            // 
            // lnkCalendarFolder
            // 
            resources.ApplyResources(this.lnkCalendarFolder, "lnkCalendarFolder");
            this.lnkCalendarFolder.Name = "lnkCalendarFolder";
            this.lnkCalendarFolder.TabStop = true;
            this.lnkCalendarFolder.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkCalendarFolder_LinkClicked);
            // 
            // label10
            // 
            resources.ApplyResources(label10, "label10");
            label10.Name = "label10";
            // 
            // syncOptionBox
            // 
            resources.ApplyResources(this.syncOptionBox, "syncOptionBox");
            this.syncOptionBox.BackColor = System.Drawing.SystemColors.Control;
            this.syncOptionBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.syncOptionBox.CheckOnClick = true;
            this.syncOptionBox.FormattingEnabled = true;
            this.syncOptionBox.Name = "syncOptionBox";
            this.toolTip.SetToolTip(this.syncOptionBox, resources.GetString("syncOptionBox.ToolTip"));
            this.syncOptionBox.SelectedIndexChanged += new System.EventHandler(this.syncOptionBox_SelectedIndexChanged);
            // 
            // notifyIcon
            // 
            this.notifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Warning;
            this.notifyIcon.ContextMenuStrip = this.systemTrayMenu;
            resources.ApplyResources(this.notifyIcon, "notifyIcon");
            this.notifyIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_MouseDoubleClick);
            // 
            // systemTrayMenu
            // 
            this.systemTrayMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mniAutoSync,
            this.mniSync,
            this.mniStopSync,
            this.toolStripSeparator2,
            this.mniSettings,
            this.mniLog,
            this.toolStripSeparator1,
            this.toolStripMenuItem5,
            this.mniRegister,
            this.toolStripMenuItem2});
            this.systemTrayMenu.Name = "systemTrayMenu";
            resources.ApplyResources(this.systemTrayMenu, "systemTrayMenu");
            // 
            // mniAutoSync
            // 
            this.mniAutoSync.Name = "mniAutoSync";
            resources.ApplyResources(this.mniAutoSync, "mniAutoSync");
            this.mniAutoSync.Click += new System.EventHandler(this.mniAutoSync_Click);
            // 
            // mniSync
            // 
            this.mniSync.Name = "mniSync";
            resources.ApplyResources(this.mniSync, "mniSync");
            this.mniSync.Click += new System.EventHandler(this.mniSync_Click);
            // 
            // mniStopSync
            // 
            resources.ApplyResources(this.mniStopSync, "mniStopSync");
            this.mniStopSync.Name = "mniStopSync";
            this.mniStopSync.Click += new System.EventHandler(this.mniStopSync_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            resources.ApplyResources(this.toolStripSeparator2, "toolStripSeparator2");
            // 
            // mniSettings
            // 
            resources.ApplyResources(this.mniSettings, "mniSettings");
            this.mniSettings.Name = "mniSettings";
            this.mniSettings.Click += new System.EventHandler(this.mniSettings_Click);
            // 
            // mniLog
            // 
            this.mniLog.Name = "mniLog";
            resources.ApplyResources(this.mniLog, "mniLog");
            this.mniLog.Click += new System.EventHandler(this.mniLog_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            resources.ApplyResources(this.toolStripSeparator1, "toolStripSeparator1");
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            resources.ApplyResources(this.toolStripMenuItem5, "toolStripMenuItem5");
            this.toolStripMenuItem5.Click += new System.EventHandler(this.SettingsForm_HelpRequested);
            // 
            // mniRegister
            // 
            this.mniRegister.Name = "mniRegister";
            resources.ApplyResources(this.mniRegister, "mniRegister");
            this.mniRegister.Visible = global::R.GoogleOutlookSync.Properties.Settings.Default.IsNotRegistered;
            this.mniRegister.Click += new System.EventHandler(this.mniRegister_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            resources.ApplyResources(this.toolStripMenuItem2, "toolStripMenuItem2");
            this.toolStripMenuItem2.Click += new System.EventHandler(this.toolStripMenuItem2_Click);
            // 
            // autoSyncInterval
            // 
            resources.ApplyResources(this.autoSyncInterval, "autoSyncInterval");
            this.autoSyncInterval.Maximum = new decimal(new int[] {
            1440,
            0,
            0,
            0});
            this.autoSyncInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.autoSyncInterval.Name = "autoSyncInterval";
            this.autoSyncInterval.Value = new decimal(new int[] {
            120,
            0,
            0,
            0});
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            // 
            // grpAutoSync
            // 
            resources.ApplyResources(this.grpAutoSync, "grpAutoSync");
            this.grpAutoSync.Controls.Add(this.nextSyncLabel);
            this.grpAutoSync.Controls.Add(this.autoSyncInterval);
            this.grpAutoSync.Controls.Add(this.label4);
            this.grpAutoSync.Controls.Add(this.label1);
            this.grpAutoSync.Name = "grpAutoSync";
            this.grpAutoSync.TabStop = false;
            // 
            // nextSyncLabel
            // 
            resources.ApplyResources(this.nextSyncLabel, "nextSyncLabel");
            this.nextSyncLabel.Name = "nextSyncLabel";
            // 
            // reportSyncResultCheckBox
            // 
            resources.ApplyResources(this.reportSyncResultCheckBox, "reportSyncResultCheckBox");
            this.reportSyncResultCheckBox.Name = "reportSyncResultCheckBox";
            this.reportSyncResultCheckBox.UseVisualStyleBackColor = true;
            // 
            // syncTimer
            // 
            this.syncTimer.Interval = 1000;
            this.syncTimer.Tick += new System.EventHandler(this.syncTimer_Tick);
            // 
            // groupBox2
            // 
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Controls.Add(this.btSyncCalendar);
            this.groupBox2.Controls.Add(this.btSyncContacts);
            this.groupBox2.Controls.Add(this.panel1);
            this.groupBox2.Controls.Add(this.syncOptionBox);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // btSyncCalendar
            // 
            resources.ApplyResources(this.btSyncCalendar, "btSyncCalendar");
            this.btSyncCalendar.Checked = true;
            this.btSyncCalendar.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btSyncCalendar.Name = "btSyncCalendar";
            this.toolTip.SetToolTip(this.btSyncCalendar, resources.GetString("btSyncCalendar.ToolTip"));
            this.btSyncCalendar.UseVisualStyleBackColor = true;
            // 
            // btSyncContacts
            // 
            resources.ApplyResources(this.btSyncContacts, "btSyncContacts");
            this.btSyncContacts.Name = "btSyncContacts";
            this.toolTip.SetToolTip(this.btSyncContacts, resources.GetString("btSyncContacts.ToolTip"));
            this.btSyncContacts.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            resources.ApplyResources(this.panel1, "panel1");
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Name = "panel1";
            // 
            // tbSyncProfile
            // 
            resources.ApplyResources(this.tbSyncProfile, "tbSyncProfile");
            this.tbSyncProfile.Name = "tbSyncProfile";
            this.tbSyncProfile.TextChanged += new System.EventHandler(this.tbSyncProfile_TextChanged);
            // 
            // resetMatchesLinkLabel
            // 
            resources.ApplyResources(this.resetMatchesLinkLabel, "resetMatchesLinkLabel");
            this.resetMatchesLinkLabel.Name = "resetMatchesLinkLabel";
            this.resetMatchesLinkLabel.TabStop = true;
            this.toolTip.SetToolTip(this.resetMatchesLinkLabel, resources.GetString("resetMatchesLinkLabel.ToolTip"));
            this.resetMatchesLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.resetMatchesLinkLabel_LinkClicked);
            // 
            // hideButton
            // 
            resources.ApplyResources(this.hideButton, "hideButton");
            this.hideButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.hideButton.Name = "hideButton";
            this.hideButton.UseVisualStyleBackColor = true;
            this.hideButton.Click += new System.EventHandler(this.hideButton_Click);
            // 
            // tsTabs
            // 
            this.tsTabs.BackColor = System.Drawing.SystemColors.Control;
            this.tsTabs.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tbtnGeneral,
            this.tbtnProxy,
            this.tbtnAdvanced});
            resources.ApplyResources(this.tsTabs, "tsTabs");
            this.tsTabs.Name = "tsTabs";
            // 
            // tbtnGeneral
            // 
            this.tbtnGeneral.CheckOnClick = true;
            this.tbtnGeneral.Image = global::R.GoogleOutlookSync.Properties.Resources.Sync1;
            resources.ApplyResources(this.tbtnGeneral, "tbtnGeneral");
            this.tbtnGeneral.Name = "tbtnGeneral";
            this.tbtnGeneral.Click += new System.EventHandler(this.PageToolButton_Click);
            // 
            // tbtnProxy
            // 
            this.tbtnProxy.CheckOnClick = true;
            this.tbtnProxy.Image = global::R.GoogleOutlookSync.Properties.Resources.Proxies;
            resources.ApplyResources(this.tbtnProxy, "tbtnProxy");
            this.tbtnProxy.Name = "tbtnProxy";
            this.tbtnProxy.Click += new System.EventHandler(this.PageToolButton_Click);
            // 
            // tbtnAdvanced
            // 
            this.tbtnAdvanced.CheckOnClick = true;
            resources.ApplyResources(this.tbtnAdvanced, "tbtnAdvanced");
            this.tbtnAdvanced.Name = "tbtnAdvanced";
            this.tbtnAdvanced.Click += new System.EventHandler(this.PageToolButton_Click);
            // 
            // pnlGeneral
            // 
            this.pnlGeneral.Controls.Add(grpOutlookSettings);
            this.pnlGeneral.Controls.Add(grpGoogleSettings);
            this.pnlGeneral.Controls.Add(this.groupBox2);
            resources.ApplyResources(this.pnlGeneral, "pnlGeneral");
            this.pnlGeneral.Name = "pnlGeneral";
            // 
            // pnlAdvanced
            // 
            this.pnlAdvanced.Controls.Add(this.lnkSendLogFile);
            this.pnlAdvanced.Controls.Add(this.lnkSendSessionLog);
            this.pnlAdvanced.Controls.Add(this.lblLanguageChangeNotification);
            this.pnlAdvanced.Controls.Add(this.cmbLanguage);
            this.pnlAdvanced.Controls.Add(this.lblLanguage);
            this.pnlAdvanced.Controls.Add(this.lnkCheckUpdates);
            this.pnlAdvanced.Controls.Add(this.reportSyncResultCheckBox);
            this.pnlAdvanced.Controls.Add(this.grpAutoSync);
            this.pnlAdvanced.Controls.Add(this.resetMatchesLinkLabel);
            resources.ApplyResources(this.pnlAdvanced, "pnlAdvanced");
            this.pnlAdvanced.Name = "pnlAdvanced";
            // 
            // lnkSendLogFile
            // 
            resources.ApplyResources(this.lnkSendLogFile, "lnkSendLogFile");
            this.lnkSendLogFile.Name = "lnkSendLogFile";
            this.lnkSendLogFile.TabStop = true;
            this.lnkSendLogFile.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkSendLogFile_LinkClicked);
            // 
            // lnkSendSessionLog
            // 
            resources.ApplyResources(this.lnkSendSessionLog, "lnkSendSessionLog");
            this.lnkSendSessionLog.Name = "lnkSendSessionLog";
            this.lnkSendSessionLog.TabStop = true;
            this.lnkSendSessionLog.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkSendSessionLog_LinkClicked);
            // 
            // lblLanguageChangeNotification
            // 
            resources.ApplyResources(this.lblLanguageChangeNotification, "lblLanguageChangeNotification");
            this.lblLanguageChangeNotification.Name = "lblLanguageChangeNotification";
            // 
            // cmbLanguage
            // 
            this.cmbLanguage.FormattingEnabled = true;
            resources.ApplyResources(this.cmbLanguage, "cmbLanguage");
            this.cmbLanguage.Name = "cmbLanguage";
            this.cmbLanguage.Sorted = true;
            this.cmbLanguage.SelectedValueChanged += new System.EventHandler(this.cmbLanguage_SelectedValueChanged);
            // 
            // lblLanguage
            // 
            resources.ApplyResources(this.lblLanguage, "lblLanguage");
            this.lblLanguage.Name = "lblLanguage";
            // 
            // lnkCheckUpdates
            // 
            resources.ApplyResources(this.lnkCheckUpdates, "lnkCheckUpdates");
            this.lnkCheckUpdates.Name = "lnkCheckUpdates";
            this.lnkCheckUpdates.TabStop = true;
            this.lnkCheckUpdates.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkCheckUpdates_LinkClicked);
            // 
            // pnlProxy
            // 
            this.pnlProxy.Controls.Add(this.groupBox3);
            resources.ApplyResources(this.pnlProxy, "pnlProxy");
            this.pnlProxy.Name = "pnlProxy";
            // 
            // groupBox3
            // 
            resources.ApplyResources(this.groupBox3, "groupBox3");
            this.groupBox3.Controls.Add(this.panel3);
            this.groupBox3.Controls.Add(this.panel2);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.Authorization);
            this.groupBox3.Controls.Add(this.txtProxyUserName);
            this.groupBox3.Controls.Add(this.CustomProxy);
            this.groupBox3.Controls.Add(this.txtProxyPassword);
            this.groupBox3.Controls.Add(this.SystemProxy);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.TabStop = false;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.Port);
            this.panel3.Controls.Add(this.label9);
            resources.ApplyResources(this.panel3, "panel3");
            this.panel3.Name = "panel3";
            // 
            // Port
            // 
            resources.ApplyResources(this.Port, "Port");
            this.Port.Name = "Port";
            this.Port.TextChanged += new System.EventHandler(this.CustomProxy_Changed);
            // 
            // label9
            // 
            resources.ApplyResources(this.label9, "label9");
            this.label9.Name = "label9";
            // 
            // panel2
            // 
            resources.ApplyResources(this.panel2, "panel2");
            this.panel2.Controls.Add(this.label8);
            this.panel2.Controls.Add(this.Address);
            this.panel2.Name = "panel2";
            // 
            // label8
            // 
            resources.ApplyResources(this.label8, "label8");
            this.label8.Name = "label8";
            // 
            // Address
            // 
            resources.ApplyResources(this.Address, "Address");
            this.Address.Name = "Address";
            this.Address.TextChanged += new System.EventHandler(this.CustomProxy_Changed);
            // 
            // label6
            // 
            resources.ApplyResources(this.label6, "label6");
            this.label6.Name = "label6";
            // 
            // label7
            // 
            resources.ApplyResources(this.label7, "label7");
            this.label7.Name = "label7";
            // 
            // Authorization
            // 
            resources.ApplyResources(this.Authorization, "Authorization");
            this.Authorization.Name = "Authorization";
            this.Authorization.UseVisualStyleBackColor = true;
            this.Authorization.CheckedChanged += new System.EventHandler(this.Authorization_CheckedChanged);
            // 
            // txtProxyUserName
            // 
            resources.ApplyResources(this.txtProxyUserName, "txtProxyUserName");
            this.txtProxyUserName.Name = "txtProxyUserName";
            this.txtProxyUserName.TextChanged += new System.EventHandler(this.CustomProxy_Changed);
            // 
            // CustomProxy
            // 
            resources.ApplyResources(this.CustomProxy, "CustomProxy");
            this.CustomProxy.Name = "CustomProxy";
            this.CustomProxy.UseVisualStyleBackColor = true;
            this.CustomProxy.CheckedChanged += new System.EventHandler(this.CustomProxy_Changed);
            // 
            // txtProxyPassword
            // 
            resources.ApplyResources(this.txtProxyPassword, "txtProxyPassword");
            this.txtProxyPassword.Name = "txtProxyPassword";
            this.txtProxyPassword.TextChanged += new System.EventHandler(this.CustomProxy_Changed);
            // 
            // SystemProxy
            // 
            resources.ApplyResources(this.SystemProxy, "SystemProxy");
            this.SystemProxy.Checked = true;
            this.SystemProxy.Name = "SystemProxy";
            this.SystemProxy.TabStop = true;
            this.SystemProxy.UseVisualStyleBackColor = true;
            // 
            // lnkHelp
            // 
            resources.ApplyResources(this.lnkHelp, "lnkHelp");
            this.lnkHelp.Name = "lnkHelp";
            this.lnkHelp.TabStop = true;
            this.lnkHelp.VisitedLinkColor = System.Drawing.Color.Blue;
            this.lnkHelp.Click += new System.EventHandler(this.pnlHelp_Click);
            // 
            // pnlHelp
            // 
            resources.ApplyResources(this.pnlHelp, "pnlHelp");
            this.pnlHelp.Controls.Add(this.lnkHelp);
            this.pnlHelp.Controls.Add(this.pictureBox1);
            this.pnlHelp.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pnlHelp.Name = "pnlHelp";
            this.pnlHelp.Click += new System.EventHandler(this.pnlHelp_Click);
            // 
            // pictureBox1
            // 
            resources.ApplyResources(this.pictureBox1, "pictureBox1");
            this.pictureBox1.Image = global::R.GoogleOutlookSync.Properties.Resources.QuestionMarkIcon;
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pnlHelp_Click);
            // 
            // lnkBuyNow
            // 
            resources.ApplyResources(this.lnkBuyNow, "lnkBuyNow");
            this.lnkBuyNow.BackColor = System.Drawing.SystemColors.Control;
            this.lnkBuyNow.DataBindings.Add(new System.Windows.Forms.Binding("Visible", global::R.GoogleOutlookSync.Properties.Settings.Default, "IsNotRegistered", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.lnkBuyNow.Name = "lnkBuyNow";
            this.lnkBuyNow.TabStop = true;
            this.lnkBuyNow.Visible = global::R.GoogleOutlookSync.Properties.Settings.Default.IsNotRegistered;
            this.lnkBuyNow.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkBuyNow_LinkClicked);
            // 
            // SettingsForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.CancelButton = this.hideButton;
            this.Controls.Add(this.hideButton);
            this.Controls.Add(this.pnlGeneral);
            this.Controls.Add(this.pnlAdvanced);
            this.Controls.Add(this.pnlProxy);
            this.Controls.Add(this.lnkBuyNow);
            this.Controls.Add(this.tbSyncProfile);
            this.Controls.Add(label5);
            this.Controls.Add(this.tsTabs);
            this.Controls.Add(this.pnlHelp);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SettingsForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SettingsForm_FormClosed);
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.SettingsForm_HelpRequested);
            grpGoogleSettings.ResumeLayout(false);
            grpGoogleSettings.PerformLayout();
            grpOutlookSettings.ResumeLayout(false);
            grpOutlookSettings.PerformLayout();
            this.systemTrayMenu.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.autoSyncInterval)).EndInit();
            this.grpAutoSync.ResumeLayout(false);
            this.grpAutoSync.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tsTabs.ResumeLayout(false);
            this.tsTabs.PerformLayout();
            this.pnlGeneral.ResumeLayout(false);
            this.pnlAdvanced.ResumeLayout(false);
            this.pnlAdvanced.PerformLayout();
            this.pnlProxy.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.pnlHelp.ResumeLayout(false);
            this.pnlHelp.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox Password;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckedListBox syncOptionBox;
        internal System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.NumericUpDown autoSyncInterval;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox grpAutoSync;
        private System.Windows.Forms.Timer syncTimer;
        private System.Windows.Forms.Label nextSyncLabel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ContextMenuStrip systemTrayMenu;
        private System.Windows.Forms.ToolStripMenuItem mniSettings;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem mniLog;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem mniSync;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.TextBox tbSyncProfile;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.LinkLabel resetMatchesLinkLabel;
        private System.Windows.Forms.Button hideButton;
        private System.Windows.Forms.CheckBox reportSyncResultCheckBox;
        private System.Windows.Forms.CheckBox btSyncContacts;
        private System.Windows.Forms.CheckBox btSyncCalendar;
        private System.Windows.Forms.ToolStrip tsTabs;
        private System.Windows.Forms.ToolStripButton tbtnGeneral;
        private System.Windows.Forms.ToolStripButton tbtnProxy;
        private System.Windows.Forms.ToolStripButton tbtnAdvanced;
        private System.Windows.Forms.Panel pnlGeneral;
        private System.Windows.Forms.Panel pnlAdvanced;
        private System.Windows.Forms.Panel pnlProxy;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox Authorization;
        private System.Windows.Forms.TextBox txtProxyUserName;
        private System.Windows.Forms.RadioButton CustomProxy;
        private System.Windows.Forms.TextBox txtProxyPassword;
        private System.Windows.Forms.RadioButton SystemProxy;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox Address;
        private System.Windows.Forms.TextBox Port;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.LinkLabel lnkHelp;
        private System.Windows.Forms.Panel pnlHelp;
        private System.Windows.Forms.ToolStripMenuItem mniStopSync;
        private System.Windows.Forms.LinkLabel lnkBuyNow;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.ToolStripMenuItem mniRegister;
        private System.Windows.Forms.LinkLabel lnkCheckUpdates;
        internal System.Windows.Forms.ToolStripMenuItem mniAutoSync;
        private System.Windows.Forms.ComboBox cmbLanguage;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.Label lblLanguageChangeNotification;
        private System.Windows.Forms.LinkLabel lnkSendSessionLog;
        private System.Windows.Forms.LinkLabel lnkSendLogFile;
        private System.Windows.Forms.LinkLabel lnkCalendarFolder;
        private System.Windows.Forms.ComboBox cmbOutlookProfiles;
    }
}


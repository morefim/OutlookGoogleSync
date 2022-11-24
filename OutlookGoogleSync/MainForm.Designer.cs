/*
 * Created by SharpDevelop.
 * User: zsianti
 * Date: 14.08.2012
 * Time: 07:54
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace OutlookGoogleSync
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.LogBox = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.tbOutlookPassword = new System.Windows.Forms.TextBox();
            this.tbOutlookUser = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.cbAddReminders = new System.Windows.Forms.CheckBox();
            this.cbAddAttendees = new System.Windows.Forms.CheckBox();
            this.cbAddDescription = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.cbStartWithWindows = new System.Windows.Forms.CheckBox();
            this.cbMinimizeToTray = new System.Windows.Forms.CheckBox();
            this.cbStartInTray = new System.Windows.Forms.CheckBox();
            this.cbCreateFiles = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tbSyncPeriod = new System.Windows.Forms.TextBox();
            this.cbShowBubbleTooltips = new System.Windows.Forms.CheckBox();
            this.cbSyncEvery = new System.Windows.Forms.CheckBox();
            this.bSave = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.bGetMyCalendars = new System.Windows.Forms.Button();
            this.cbCalendars = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbDaysInTheFuture = new System.Windows.Forms.TextBox();
            this.tbDaysInThePast = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonIconize = new System.Windows.Forms.Button();
            this.bSyncNow = new System.Windows.Forms.Button();
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.notifyIconContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.showToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.sincNowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cbByPassKLALogoutPolicy = new System.Windows.Forms.CheckBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.notifyIconContextMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(543, 437);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.LogBox);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(535, 390);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Sync";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // LogBox
            // 
            this.LogBox.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.LogBox.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LogBox.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.LogBox.Location = new System.Drawing.Point(6, 6);
            this.LogBox.Multiline = true;
            this.LogBox.Name = "LogBox";
            this.LogBox.ReadOnly = true;
            this.LogBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.LogBox.Size = new System.Drawing.Size(523, 374);
            this.LogBox.TabIndex = 1;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox6);
            this.tabPage2.Controls.Add(this.groupBox5);
            this.tabPage2.Controls.Add(this.groupBox4);
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Controls.Add(this.bSave);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(535, 411);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Settings";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.tbOutlookPassword);
            this.groupBox6.Controls.Add(this.tbOutlookUser);
            this.groupBox6.Controls.Add(this.label5);
            this.groupBox6.Controls.Add(this.label4);
            this.groupBox6.Location = new System.Drawing.Point(6, 171);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(523, 62);
            this.groupBox6.TabIndex = 13;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Outlook";
            // 
            // tbOutlookPassword
            // 
            this.tbOutlookPassword.Location = new System.Drawing.Point(337, 25);
            this.tbOutlookPassword.Name = "tbOutlookPassword";
            this.tbOutlookPassword.Size = new System.Drawing.Size(173, 20);
            this.tbOutlookPassword.TabIndex = 3;
            this.tbOutlookPassword.TextChanged += new System.EventHandler(this.tbOutlookPassword_TextChanged);
            // 
            // tbOutlookUser
            // 
            this.tbOutlookUser.Location = new System.Drawing.Point(47, 25);
            this.tbOutlookUser.Name = "tbOutlookUser";
            this.tbOutlookUser.Size = new System.Drawing.Size(201, 20);
            this.tbOutlookUser.TabIndex = 2;
            this.tbOutlookUser.TextChanged += new System.EventHandler(this.tbOutlookUser_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(275, 28);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(56, 13);
            this.label5.TabIndex = 1;
            this.label5.Text = "Password:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 28);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Email:";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.cbAddReminders);
            this.groupBox5.Controls.Add(this.cbAddAttendees);
            this.groupBox5.Controls.Add(this.cbAddDescription);
            this.groupBox5.Location = new System.Drawing.Point(272, 239);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(257, 129);
            this.groupBox5.TabIndex = 12;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "When creating Google Calendar Entries...   ";
            // 
            // cbAddReminders
            // 
            this.cbAddReminders.Location = new System.Drawing.Point(12, 75);
            this.cbAddReminders.Name = "cbAddReminders";
            this.cbAddReminders.Size = new System.Drawing.Size(139, 24);
            this.cbAddReminders.TabIndex = 8;
            this.cbAddReminders.Text = "Add Reminders";
            this.cbAddReminders.UseVisualStyleBackColor = true;
            this.cbAddReminders.CheckedChanged += new System.EventHandler(this.CbAddRemindersCheckedChanged);
            // 
            // cbAddAttendees
            // 
            this.cbAddAttendees.Location = new System.Drawing.Point(12, 45);
            this.cbAddAttendees.Name = "cbAddAttendees";
            this.cbAddAttendees.Size = new System.Drawing.Size(235, 24);
            this.cbAddAttendees.TabIndex = 6;
            this.cbAddAttendees.Text = "Add Attendees at the end of the Description";
            this.cbAddAttendees.UseVisualStyleBackColor = true;
            this.cbAddAttendees.CheckedChanged += new System.EventHandler(this.cbAddAttendees_CheckedChanged);
            // 
            // cbAddDescription
            // 
            this.cbAddDescription.Location = new System.Drawing.Point(12, 19);
            this.cbAddDescription.Name = "cbAddDescription";
            this.cbAddDescription.Size = new System.Drawing.Size(209, 24);
            this.cbAddDescription.TabIndex = 7;
            this.cbAddDescription.Text = "Add Description";
            this.cbAddDescription.UseVisualStyleBackColor = true;
            this.cbAddDescription.CheckedChanged += new System.EventHandler(this.CbAddDescriptionCheckedChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.cbByPassKLALogoutPolicy);
            this.groupBox4.Controls.Add(this.cbStartWithWindows);
            this.groupBox4.Controls.Add(this.cbMinimizeToTray);
            this.groupBox4.Controls.Add(this.cbStartInTray);
            this.groupBox4.Controls.Add(this.cbCreateFiles);
            this.groupBox4.Location = new System.Drawing.Point(7, 239);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(256, 161);
            this.groupBox4.TabIndex = 11;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Options";
            // 
            // cbStartWithWindows
            // 
            this.cbStartWithWindows.Location = new System.Drawing.Point(12, 15);
            this.cbStartWithWindows.Name = "cbStartWithWindows";
            this.cbStartWithWindows.Size = new System.Drawing.Size(209, 24);
            this.cbStartWithWindows.TabIndex = 8;
            this.cbStartWithWindows.Text = "Start with Windows";
            this.cbStartWithWindows.UseVisualStyleBackColor = true;
            this.cbStartWithWindows.CheckedChanged += new System.EventHandler(this.cbStartWithWindows_CheckedChanged);
            // 
            // cbMinimizeToTray
            // 
            this.cbMinimizeToTray.Location = new System.Drawing.Point(12, 75);
            this.cbMinimizeToTray.Name = "cbMinimizeToTray";
            this.cbMinimizeToTray.Size = new System.Drawing.Size(104, 24);
            this.cbMinimizeToTray.TabIndex = 0;
            this.cbMinimizeToTray.Text = "Minimize to Tray";
            this.cbMinimizeToTray.UseVisualStyleBackColor = true;
            this.cbMinimizeToTray.CheckedChanged += new System.EventHandler(this.CbMinimizeToTrayCheckedChanged);
            // 
            // cbStartInTray
            // 
            this.cbStartInTray.Location = new System.Drawing.Point(12, 45);
            this.cbStartInTray.Name = "cbStartInTray";
            this.cbStartInTray.Size = new System.Drawing.Size(104, 24);
            this.cbStartInTray.TabIndex = 1;
            this.cbStartInTray.Text = "Start in Tray";
            this.cbStartInTray.UseVisualStyleBackColor = true;
            this.cbStartInTray.CheckedChanged += new System.EventHandler(this.CbStartInTrayCheckedChanged);
            // 
            // cbCreateFiles
            // 
            this.cbCreateFiles.Location = new System.Drawing.Point(12, 105);
            this.cbCreateFiles.Name = "cbCreateFiles";
            this.cbCreateFiles.Size = new System.Drawing.Size(235, 24);
            this.cbCreateFiles.TabIndex = 7;
            this.cbCreateFiles.Text = "Create text files with found/identified entries";
            this.cbCreateFiles.UseVisualStyleBackColor = true;
            this.cbCreateFiles.CheckedChanged += new System.EventHandler(this.cbCreateFiles_CheckedChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.tbSyncPeriod);
            this.groupBox3.Controls.Add(this.cbShowBubbleTooltips);
            this.groupBox3.Controls.Add(this.cbSyncEvery);
            this.groupBox3.Location = new System.Drawing.Point(225, 80);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(304, 85);
            this.groupBox3.TabIndex = 10;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Sync Regularly";
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(143, 24);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(89, 17);
            this.label6.TabIndex = 9;
            this.label6.Text = "minutes";
            // 
            // tbSyncPeriod
            // 
            this.tbSyncPeriod.Location = new System.Drawing.Point(98, 21);
            this.tbSyncPeriod.Name = "tbSyncPeriod";
            this.tbSyncPeriod.Size = new System.Drawing.Size(39, 20);
            this.tbSyncPeriod.TabIndex = 8;
            this.tbSyncPeriod.TextChanged += new System.EventHandler(this.tbSyncPeriod_TextChanged);
            // 
            // cbShowBubbleTooltips
            // 
            this.cbShowBubbleTooltips.Location = new System.Drawing.Point(9, 49);
            this.cbShowBubbleTooltips.Name = "cbShowBubbleTooltips";
            this.cbShowBubbleTooltips.Size = new System.Drawing.Size(259, 24);
            this.cbShowBubbleTooltips.TabIndex = 7;
            this.cbShowBubbleTooltips.Text = "Show Bubble Tooltip in Taskbar after Syncing";
            this.cbShowBubbleTooltips.UseVisualStyleBackColor = true;
            this.cbShowBubbleTooltips.CheckedChanged += new System.EventHandler(this.CbShowBubbleTooltipsCheckedChanged);
            // 
            // cbSyncEvery
            // 
            this.cbSyncEvery.Location = new System.Drawing.Point(9, 19);
            this.cbSyncEvery.Name = "cbSyncEvery";
            this.cbSyncEvery.Size = new System.Drawing.Size(83, 24);
            this.cbSyncEvery.TabIndex = 6;
            this.cbSyncEvery.Text = "Sync every:";
            this.cbSyncEvery.UseVisualStyleBackColor = true;
            this.cbSyncEvery.CheckedChanged += new System.EventHandler(this.CbSyncEveryHourCheckedChanged);
            // 
            // bSave
            // 
            this.bSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bSave.Image = global::OutlookGoogleSync.Properties.Resources.save;
            this.bSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.bSave.Location = new System.Drawing.Point(439, 374);
            this.bSave.Name = "bSave";
            this.bSave.Size = new System.Drawing.Size(90, 26);
            this.bSave.TabIndex = 8;
            this.bSave.Text = "Save      ";
            this.bSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.bSave.UseVisualStyleBackColor = true;
            this.bSave.Click += new System.EventHandler(this.Save_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.bGetMyCalendars);
            this.groupBox2.Controls.Add(this.cbCalendars);
            this.groupBox2.Location = new System.Drawing.Point(6, 6);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(523, 68);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Google Calendar";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(6, 33);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 23);
            this.label3.TabIndex = 3;
            this.label3.Text = "Use Google Calendar:";
            // 
            // bGetMyCalendars
            // 
            this.bGetMyCalendars.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bGetMyCalendars.Image = global::OutlookGoogleSync.Properties.Resources.google;
            this.bGetMyCalendars.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.bGetMyCalendars.Location = new System.Drawing.Point(356, 26);
            this.bGetMyCalendars.Name = "bGetMyCalendars";
            this.bGetMyCalendars.Size = new System.Drawing.Size(154, 26);
            this.bGetMyCalendars.TabIndex = 2;
            this.bGetMyCalendars.Text = "Get My Google Calendars";
            this.bGetMyCalendars.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.bGetMyCalendars.UseVisualStyleBackColor = true;
            this.bGetMyCalendars.Click += new System.EventHandler(this.GetMyGoogleCalendars_Click);
            // 
            // cbCalendars
            // 
            this.cbCalendars.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbCalendars.FormattingEnabled = true;
            this.cbCalendars.Location = new System.Drawing.Point(124, 30);
            this.cbCalendars.Name = "cbCalendars";
            this.cbCalendars.Size = new System.Drawing.Size(226, 21);
            this.cbCalendars.TabIndex = 1;
            this.cbCalendars.SelectedIndexChanged += new System.EventHandler(this.ComboBox1SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tbDaysInTheFuture);
            this.groupBox1.Controls.Add(this.tbDaysInThePast);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(6, 80);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(213, 85);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Sync Date Range";
            // 
            // tbDaysInTheFuture
            // 
            this.tbDaysInTheFuture.Location = new System.Drawing.Point(124, 51);
            this.tbDaysInTheFuture.Name = "tbDaysInTheFuture";
            this.tbDaysInTheFuture.Size = new System.Drawing.Size(39, 20);
            this.tbDaysInTheFuture.TabIndex = 4;
            this.tbDaysInTheFuture.TextChanged += new System.EventHandler(this.TbDaysInTheFutureTextChanged);
            // 
            // tbDaysInThePast
            // 
            this.tbDaysInThePast.Location = new System.Drawing.Point(124, 19);
            this.tbDaysInThePast.Name = "tbDaysInThePast";
            this.tbDaysInThePast.Size = new System.Drawing.Size(39, 20);
            this.tbDaysInThePast.TabIndex = 3;
            this.tbDaysInThePast.TextChanged += new System.EventHandler(this.TbDaysInThePastTextChanged);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(6, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 23);
            this.label2.TabIndex = 0;
            this.label2.Text = "Days in the Future:";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(6, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Days in the Past:";
            // 
            // buttonIconize
            // 
            this.buttonIconize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonIconize.Image = global::OutlookGoogleSync.Properties.Resources.calendar;
            this.buttonIconize.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonIconize.Location = new System.Drawing.Point(465, 455);
            this.buttonIconize.Name = "buttonIconize";
            this.buttonIconize.Size = new System.Drawing.Size(90, 26);
            this.buttonIconize.TabIndex = 2;
            this.buttonIconize.Text = "Iconize     ";
            this.buttonIconize.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonIconize.UseVisualStyleBackColor = true;
            this.buttonIconize.Click += new System.EventHandler(this.buttonIconize_Click);
            // 
            // bSyncNow
            // 
            this.bSyncNow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bSyncNow.Image = global::OutlookGoogleSync.Properties.Resources.sync;
            this.bSyncNow.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.bSyncNow.Location = new System.Drawing.Point(369, 455);
            this.bSyncNow.Name = "bSyncNow";
            this.bSyncNow.Size = new System.Drawing.Size(90, 26);
            this.bSyncNow.TabIndex = 0;
            this.bSyncNow.Text = "Sync now  ";
            this.bSyncNow.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.bSyncNow.UseVisualStyleBackColor = true;
            this.bSyncNow.Click += new System.EventHandler(this.SyncNow_Click);
            // 
            // notifyIcon
            // 
            this.notifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.notifyIcon.BalloonTipText = "Ready";
            this.notifyIcon.BalloonTipTitle = "Outlook Google Sync";
            this.notifyIcon.ContextMenuStrip = this.notifyIconContextMenu;
            this.notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));
            this.notifyIcon.Text = "Outlook Google Sync: Ready";
            this.notifyIcon.DoubleClick += new System.EventHandler(this.NotifyIconDoubleClick);
            // 
            // notifyIconContextMenu
            // 
            this.notifyIconContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.showToolStripMenuItem,
            this.toolStripMenuItem2,
            this.sincNowToolStripMenuItem,
            this.toolStripMenuItem1,
            this.exitToolStripMenuItem});
            this.notifyIconContextMenu.Name = "notifyIconContextMenu";
            this.notifyIconContextMenu.Size = new System.Drawing.Size(125, 82);
            // 
            // showToolStripMenuItem
            // 
            this.showToolStripMenuItem.Name = "showToolStripMenuItem";
            this.showToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.showToolStripMenuItem.Text = "Show...";
            this.showToolStripMenuItem.Click += new System.EventHandler(this.showToolStripMenuItem_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(121, 6);
            // 
            // sincNowToolStripMenuItem
            // 
            this.sincNowToolStripMenuItem.Name = "sincNowToolStripMenuItem";
            this.sincNowToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.sincNowToolStripMenuItem.Text = "Sinc Now";
            this.sincNowToolStripMenuItem.Click += new System.EventHandler(this.sincNowToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(121, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // cbByPassKLALogoutPolicy
            // 
            this.cbByPassKLALogoutPolicy.Location = new System.Drawing.Point(12, 131);
            this.cbByPassKLALogoutPolicy.Name = "cbByPassKLALogoutPolicy";
            this.cbByPassKLALogoutPolicy.Size = new System.Drawing.Size(171, 24);
            this.cbByPassKLALogoutPolicy.TabIndex = 9;
            this.cbByPassKLALogoutPolicy.Text = "Bypass KLA Logout Policy";
            this.cbByPassKLALogoutPolicy.UseVisualStyleBackColor = true;
            this.cbByPassKLALogoutPolicy.CheckedChanged += new System.EventHandler(this.cbByPassKLALogoutPolicy_CheckedChanged);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(567, 490);
            this.Controls.Add(this.buttonIconize);
            this.Controls.Add(this.bSyncNow);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = " Outlook Google Sync";
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.Resize += new System.EventHandler(this.MainFormResize);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.notifyIconContextMenu.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		private System.Windows.Forms.CheckBox cbAddReminders;
		private System.Windows.Forms.CheckBox cbAddDescription;
		private System.Windows.Forms.CheckBox cbShowBubbleTooltips;
		private System.Windows.Forms.CheckBox cbSyncEvery;
		private System.Windows.Forms.CheckBox cbMinimizeToTray;
		private System.Windows.Forms.CheckBox cbStartInTray;
		private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.NotifyIcon notifyIcon;
		private System.Windows.Forms.CheckBox cbAddAttendees;
		private System.Windows.Forms.CheckBox cbCreateFiles;
		private System.Windows.Forms.TextBox LogBox;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label3;
		public System.Windows.Forms.ComboBox cbCalendars;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox tbDaysInThePast;
		private System.Windows.Forms.TextBox tbDaysInTheFuture;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button bSave;
		private System.Windows.Forms.Button bGetMyCalendars;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.Button bSyncNow;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.CheckBox cbStartWithWindows;
        private System.Windows.Forms.ContextMenuStrip notifyIconContextMenu;
        private System.Windows.Forms.ToolStripMenuItem showToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.Button buttonIconize;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem sincNowToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.TextBox tbOutlookPassword;
        private System.Windows.Forms.TextBox tbOutlookUser;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbSyncPeriod;
        private System.Windows.Forms.CheckBox cbByPassKLALogoutPolicy;
    }
}

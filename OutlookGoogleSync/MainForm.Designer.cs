﻿/*
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
         if (disposing)
         {
            if (components != null)
            {
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
         this.clearUserPropertiesBtn = new System.Windows.Forms.Button();
         this.lNextSync = new System.Windows.Forms.Label();
         this.lLastSync = new System.Windows.Forms.Label();
         this.LogBox = new System.Windows.Forms.TextBox();
         this.bSyncNow = new System.Windows.Forms.Button();
         this.tabPage2 = new System.Windows.Forms.TabPage();
         this.groupBox6 = new System.Windows.Forms.GroupBox();
         this.label6 = new System.Windows.Forms.Label();
         this.label5 = new System.Windows.Forms.Label();
         this.outlookAutoLogonPwdTextBox = new System.Windows.Forms.TextBox();
         this.outlookAutoLogonTextBox = new System.Windows.Forms.TextBox();
         this.outlookAutoLogonCheckBox = new System.Windows.Forms.CheckBox();
         this.groupBox5 = new System.Windows.Forms.GroupBox();
         this.cbAddReminders = new System.Windows.Forms.CheckBox();
         this.cbAddAttendees = new System.Windows.Forms.CheckBox();
         this.cbAddDescription = new System.Windows.Forms.CheckBox();
         this.groupBox4 = new System.Windows.Forms.GroupBox();
         this.cbMinimizeToTray = new System.Windows.Forms.CheckBox();
         this.cbStartInTray = new System.Windows.Forms.CheckBox();
         this.cbCreateFiles = new System.Windows.Forms.CheckBox();
         this.groupBox3 = new System.Windows.Forms.GroupBox();
         this.cbShowBubbleTooltips = new System.Windows.Forms.CheckBox();
         this.cbSyncEveryHour = new System.Windows.Forms.CheckBox();
         this.tbMinuteOffsets = new System.Windows.Forms.TextBox();
         this.groupBox2 = new System.Windows.Forms.GroupBox();
         this.label3 = new System.Windows.Forms.Label();
         this.bGetMyCalendars = new System.Windows.Forms.Button();
         this.cbCalendars = new System.Windows.Forms.ComboBox();
         this.bSave = new System.Windows.Forms.Button();
         this.groupBox1 = new System.Windows.Forms.GroupBox();
         this.tbDaysInTheFuture = new System.Windows.Forms.TextBox();
         this.tbDaysInThePast = new System.Windows.Forms.TextBox();
         this.label2 = new System.Windows.Forms.Label();
         this.label1 = new System.Windows.Forms.Label();
         this.tabPage3 = new System.Windows.Forms.TabPage();
         this.pictureBox1 = new System.Windows.Forms.PictureBox();
         this.linkLabel1 = new System.Windows.Forms.LinkLabel();
         this.label4 = new System.Windows.Forms.Label();
         this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
         this.outlookKeepOpenAfterSync = new System.Windows.Forms.CheckBox();
         this.tabControl1.SuspendLayout();
         this.tabPage1.SuspendLayout();
         this.tabPage2.SuspendLayout();
         this.groupBox6.SuspendLayout();
         this.groupBox5.SuspendLayout();
         this.groupBox4.SuspendLayout();
         this.groupBox3.SuspendLayout();
         this.groupBox2.SuspendLayout();
         this.groupBox1.SuspendLayout();
         this.tabPage3.SuspendLayout();
         ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
         this.SuspendLayout();
         // 
         // tabControl1
         // 
         this.tabControl1.Controls.Add(this.tabPage1);
         this.tabControl1.Controls.Add(this.tabPage2);
         this.tabControl1.Controls.Add(this.tabPage3);
         this.tabControl1.Location = new System.Drawing.Point(12, 12);
         this.tabControl1.Name = "tabControl1";
         this.tabControl1.SelectedIndex = 0;
         this.tabControl1.Size = new System.Drawing.Size(495, 505);
         this.tabControl1.TabIndex = 0;
         // 
         // tabPage1
         // 
         this.tabPage1.Controls.Add(this.clearUserPropertiesBtn);
         this.tabPage1.Controls.Add(this.lNextSync);
         this.tabPage1.Controls.Add(this.lLastSync);
         this.tabPage1.Controls.Add(this.LogBox);
         this.tabPage1.Controls.Add(this.bSyncNow);
         this.tabPage1.Location = new System.Drawing.Point(4, 22);
         this.tabPage1.Name = "tabPage1";
         this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
         this.tabPage1.Size = new System.Drawing.Size(487, 479);
         this.tabPage1.TabIndex = 0;
         this.tabPage1.Text = "Sync";
         this.tabPage1.UseVisualStyleBackColor = true;
         // 
         // clearUserPropertiesBtn
         // 
         this.clearUserPropertiesBtn.Location = new System.Drawing.Point(361, 442);
         this.clearUserPropertiesBtn.Name = "clearUserPropertiesBtn";
         this.clearUserPropertiesBtn.Size = new System.Drawing.Size(120, 31);
         this.clearUserPropertiesBtn.TabIndex = 3;
         this.clearUserPropertiesBtn.Text = "Clear Properties (Dev)";
         this.clearUserPropertiesBtn.UseVisualStyleBackColor = true;
         this.clearUserPropertiesBtn.Click += new System.EventHandler(this.clearUserPropertiesBtn_Click);
         // 
         // lNextSync
         // 
         this.lNextSync.Location = new System.Drawing.Point(252, 14);
         this.lNextSync.Name = "lNextSync";
         this.lNextSync.Size = new System.Drawing.Size(232, 31);
         this.lNextSync.TabIndex = 2;
         this.lNextSync.Text = "Next scheduled sync: ";
         // 
         // lLastSync
         // 
         this.lLastSync.Location = new System.Drawing.Point(5, 14);
         this.lLastSync.Name = "lLastSync";
         this.lLastSync.Size = new System.Drawing.Size(251, 31);
         this.lLastSync.TabIndex = 2;
         this.lLastSync.Text = "Last succeded synchro: ";
         // 
         // LogBox
         // 
         this.LogBox.AcceptsTab = true;
         this.LogBox.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
         this.LogBox.Location = new System.Drawing.Point(3, 57);
         this.LogBox.MaxLength = 500000;
         this.LogBox.Multiline = true;
         this.LogBox.Name = "LogBox";
         this.LogBox.ReadOnly = true;
         this.LogBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
         this.LogBox.Size = new System.Drawing.Size(481, 379);
         this.LogBox.TabIndex = 1;
         // 
         // bSyncNow
         // 
         this.bSyncNow.Location = new System.Drawing.Point(4, 442);
         this.bSyncNow.Name = "bSyncNow";
         this.bSyncNow.Size = new System.Drawing.Size(98, 31);
         this.bSyncNow.TabIndex = 0;
         this.bSyncNow.Text = "Sync Now";
         this.bSyncNow.UseVisualStyleBackColor = true;
         this.bSyncNow.Click += new System.EventHandler(this.SyncNow_Click);
         // 
         // tabPage2
         // 
         this.tabPage2.Controls.Add(this.groupBox6);
         this.tabPage2.Controls.Add(this.groupBox5);
         this.tabPage2.Controls.Add(this.groupBox4);
         this.tabPage2.Controls.Add(this.groupBox3);
         this.tabPage2.Controls.Add(this.groupBox2);
         this.tabPage2.Controls.Add(this.bSave);
         this.tabPage2.Controls.Add(this.groupBox1);
         this.tabPage2.Location = new System.Drawing.Point(4, 22);
         this.tabPage2.Name = "tabPage2";
         this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
         this.tabPage2.Size = new System.Drawing.Size(487, 479);
         this.tabPage2.TabIndex = 1;
         this.tabPage2.Text = "Settings";
         this.tabPage2.UseVisualStyleBackColor = true;
         // 
         // groupBox6
         // 
         this.groupBox6.Controls.Add(this.outlookKeepOpenAfterSync);
         this.groupBox6.Controls.Add(this.label6);
         this.groupBox6.Controls.Add(this.label5);
         this.groupBox6.Controls.Add(this.outlookAutoLogonPwdTextBox);
         this.groupBox6.Controls.Add(this.outlookAutoLogonTextBox);
         this.groupBox6.Controls.Add(this.outlookAutoLogonCheckBox);
         this.groupBox6.Location = new System.Drawing.Point(6, 80);
         this.groupBox6.Name = "groupBox6";
         this.groupBox6.Size = new System.Drawing.Size(475, 67);
         this.groupBox6.TabIndex = 13;
         this.groupBox6.TabStop = false;
         this.groupBox6.Text = "Outlook Calendar";
         // 
         // label6
         // 
         this.label6.AutoSize = true;
         this.label6.Location = new System.Drawing.Point(167, 43);
         this.label6.Name = "label6";
         this.label6.Size = new System.Drawing.Size(56, 13);
         this.label6.TabIndex = 4;
         this.label6.Text = "Password:";
         // 
         // label5
         // 
         this.label5.AutoSize = true;
         this.label5.Location = new System.Drawing.Point(184, 16);
         this.label5.Name = "label5";
         this.label5.Size = new System.Drawing.Size(39, 13);
         this.label5.TabIndex = 3;
         this.label5.Text = "Profile:";
         // 
         // outlookAutoLogonPwdTextBox
         // 
         this.outlookAutoLogonPwdTextBox.Location = new System.Drawing.Point(228, 39);
         this.outlookAutoLogonPwdTextBox.Name = "outlookAutoLogonPwdTextBox";
         this.outlookAutoLogonPwdTextBox.PasswordChar = '*';
         this.outlookAutoLogonPwdTextBox.ReadOnly = true;
         this.outlookAutoLogonPwdTextBox.Size = new System.Drawing.Size(240, 20);
         this.outlookAutoLogonPwdTextBox.TabIndex = 2;
         this.outlookAutoLogonPwdTextBox.UseSystemPasswordChar = true;
         // 
         // outlookAutoLogonTextBox
         // 
         this.outlookAutoLogonTextBox.Location = new System.Drawing.Point(228, 12);
         this.outlookAutoLogonTextBox.Name = "outlookAutoLogonTextBox";
         this.outlookAutoLogonTextBox.ReadOnly = true;
         this.outlookAutoLogonTextBox.Size = new System.Drawing.Size(240, 20);
         this.outlookAutoLogonTextBox.TabIndex = 1;
         // 
         // outlookAutoLogonCheckBox
         // 
         this.outlookAutoLogonCheckBox.AutoSize = true;
         this.outlookAutoLogonCheckBox.Location = new System.Drawing.Point(12, 17);
         this.outlookAutoLogonCheckBox.Name = "outlookAutoLogonCheckBox";
         this.outlookAutoLogonCheckBox.Size = new System.Drawing.Size(81, 17);
         this.outlookAutoLogonCheckBox.TabIndex = 0;
         this.outlookAutoLogonCheckBox.Text = "Auto Logon";
         this.outlookAutoLogonCheckBox.UseVisualStyleBackColor = true;
         this.outlookAutoLogonCheckBox.CheckedChanged += new System.EventHandler(this.outlookAutoLogonCheckBox_CheckedChanged);
         // 
         // groupBox5
         // 
         this.groupBox5.Controls.Add(this.cbAddReminders);
         this.groupBox5.Controls.Add(this.cbAddAttendees);
         this.groupBox5.Controls.Add(this.cbAddDescription);
         this.groupBox5.Location = new System.Drawing.Point(7, 245);
         this.groupBox5.Name = "groupBox5";
         this.groupBox5.Size = new System.Drawing.Size(475, 95);
         this.groupBox5.TabIndex = 12;
         this.groupBox5.TabStop = false;
         this.groupBox5.Text = "When creating Google Calendar Entries...   ";
         // 
         // cbAddReminders
         // 
         this.cbAddReminders.Location = new System.Drawing.Point(12, 68);
         this.cbAddReminders.Name = "cbAddReminders";
         this.cbAddReminders.Size = new System.Drawing.Size(139, 24);
         this.cbAddReminders.TabIndex = 8;
         this.cbAddReminders.Text = "Add Reminders";
         this.cbAddReminders.UseVisualStyleBackColor = true;
         this.cbAddReminders.CheckedChanged += new System.EventHandler(this.CbAddRemindersCheckedChanged);
         // 
         // cbAddAttendees
         // 
         this.cbAddAttendees.Location = new System.Drawing.Point(12, 41);
         this.cbAddAttendees.Name = "cbAddAttendees";
         this.cbAddAttendees.Size = new System.Drawing.Size(235, 24);
         this.cbAddAttendees.TabIndex = 6;
         this.cbAddAttendees.Text = "Add Attendees at the end of the Description";
         this.cbAddAttendees.UseVisualStyleBackColor = true;
         this.cbAddAttendees.CheckedChanged += new System.EventHandler(this.cbAddAttendees_CheckedChanged);
         // 
         // cbAddDescription
         // 
         this.cbAddDescription.Location = new System.Drawing.Point(12, 14);
         this.cbAddDescription.Name = "cbAddDescription";
         this.cbAddDescription.Size = new System.Drawing.Size(209, 24);
         this.cbAddDescription.TabIndex = 7;
         this.cbAddDescription.Text = "Add Description";
         this.cbAddDescription.UseVisualStyleBackColor = true;
         this.cbAddDescription.CheckedChanged += new System.EventHandler(this.CbAddDescriptionCheckedChanged);
         // 
         // groupBox4
         // 
         this.groupBox4.Controls.Add(this.cbMinimizeToTray);
         this.groupBox4.Controls.Add(this.cbStartInTray);
         this.groupBox4.Controls.Add(this.cbCreateFiles);
         this.groupBox4.Location = new System.Drawing.Point(6, 344);
         this.groupBox4.Name = "groupBox4";
         this.groupBox4.Size = new System.Drawing.Size(475, 94);
         this.groupBox4.TabIndex = 11;
         this.groupBox4.TabStop = false;
         this.groupBox4.Text = "Options";
         // 
         // cbMinimizeToTray
         // 
         this.cbMinimizeToTray.Location = new System.Drawing.Point(12, 41);
         this.cbMinimizeToTray.Name = "cbMinimizeToTray";
         this.cbMinimizeToTray.Size = new System.Drawing.Size(104, 24);
         this.cbMinimizeToTray.TabIndex = 0;
         this.cbMinimizeToTray.Text = "Minimize to Tray";
         this.cbMinimizeToTray.UseVisualStyleBackColor = true;
         this.cbMinimizeToTray.CheckedChanged += new System.EventHandler(this.CbMinimizeToTrayCheckedChanged);
         // 
         // cbStartInTray
         // 
         this.cbStartInTray.Location = new System.Drawing.Point(12, 14);
         this.cbStartInTray.Name = "cbStartInTray";
         this.cbStartInTray.Size = new System.Drawing.Size(104, 24);
         this.cbStartInTray.TabIndex = 1;
         this.cbStartInTray.Text = "Start in Tray";
         this.cbStartInTray.UseVisualStyleBackColor = true;
         this.cbStartInTray.CheckedChanged += new System.EventHandler(this.CbStartInTrayCheckedChanged);
         // 
         // cbCreateFiles
         // 
         this.cbCreateFiles.Location = new System.Drawing.Point(12, 68);
         this.cbCreateFiles.Name = "cbCreateFiles";
         this.cbCreateFiles.Size = new System.Drawing.Size(235, 24);
         this.cbCreateFiles.TabIndex = 7;
         this.cbCreateFiles.Text = "Create text files with found/identified entries";
         this.cbCreateFiles.UseVisualStyleBackColor = true;
         this.cbCreateFiles.CheckedChanged += new System.EventHandler(this.cbCreateFiles_CheckedChanged);
         // 
         // groupBox3
         // 
         this.groupBox3.Controls.Add(this.cbShowBubbleTooltips);
         this.groupBox3.Controls.Add(this.cbSyncEveryHour);
         this.groupBox3.Controls.Add(this.tbMinuteOffsets);
         this.groupBox3.Location = new System.Drawing.Point(177, 153);
         this.groupBox3.Name = "groupBox3";
         this.groupBox3.Size = new System.Drawing.Size(304, 85);
         this.groupBox3.TabIndex = 10;
         this.groupBox3.TabStop = false;
         this.groupBox3.Text = "Sync Regularly";
         // 
         // cbShowBubbleTooltips
         // 
         this.cbShowBubbleTooltips.Location = new System.Drawing.Point(6, 49);
         this.cbShowBubbleTooltips.Name = "cbShowBubbleTooltips";
         this.cbShowBubbleTooltips.Size = new System.Drawing.Size(259, 24);
         this.cbShowBubbleTooltips.TabIndex = 7;
         this.cbShowBubbleTooltips.Text = "Show Bubble Tooltip in Taskbar when Syncing";
         this.cbShowBubbleTooltips.UseVisualStyleBackColor = true;
         this.cbShowBubbleTooltips.CheckedChanged += new System.EventHandler(this.CbShowBubbleTooltipsCheckedChanged);
         // 
         // cbSyncEveryHour
         // 
         this.cbSyncEveryHour.Location = new System.Drawing.Point(6, 19);
         this.cbSyncEveryHour.Name = "cbSyncEveryHour";
         this.cbSyncEveryHour.Size = new System.Drawing.Size(180, 24);
         this.cbSyncEveryHour.TabIndex = 6;
         this.cbSyncEveryHour.Text = "Delay between sync (in minutes)";
         this.cbSyncEveryHour.UseVisualStyleBackColor = true;
         this.cbSyncEveryHour.CheckedChanged += new System.EventHandler(this.CbSyncEveryHourCheckedChanged);
         // 
         // tbMinuteOffsets
         // 
         this.tbMinuteOffsets.Location = new System.Drawing.Point(192, 21);
         this.tbMinuteOffsets.Name = "tbMinuteOffsets";
         this.tbMinuteOffsets.Size = new System.Drawing.Size(106, 20);
         this.tbMinuteOffsets.TabIndex = 5;
         this.tbMinuteOffsets.TextChanged += new System.EventHandler(this.TbMinuteOffsetsTextChanged);
         this.tbMinuteOffsets.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NumericOnlyKeyPress);
         // 
         // groupBox2
         // 
         this.groupBox2.Controls.Add(this.label3);
         this.groupBox2.Controls.Add(this.bGetMyCalendars);
         this.groupBox2.Controls.Add(this.cbCalendars);
         this.groupBox2.Location = new System.Drawing.Point(6, 6);
         this.groupBox2.Name = "groupBox2";
         this.groupBox2.Size = new System.Drawing.Size(475, 68);
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
         this.bGetMyCalendars.Location = new System.Drawing.Point(363, 19);
         this.bGetMyCalendars.Name = "bGetMyCalendars";
         this.bGetMyCalendars.Size = new System.Drawing.Size(106, 40);
         this.bGetMyCalendars.TabIndex = 2;
         this.bGetMyCalendars.Text = "Get My\r\nGoogle Calendars";
         this.bGetMyCalendars.UseVisualStyleBackColor = true;
         this.bGetMyCalendars.Click += new System.EventHandler(this.GetMyGoogleCalendars_Click);
         // 
         // cbCalendars
         // 
         this.cbCalendars.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
         this.cbCalendars.FormattingEnabled = true;
         this.cbCalendars.Location = new System.Drawing.Point(124, 30);
         this.cbCalendars.Name = "cbCalendars";
         this.cbCalendars.Size = new System.Drawing.Size(225, 21);
         this.cbCalendars.TabIndex = 1;
         this.cbCalendars.SelectedIndexChanged += new System.EventHandler(this.ComboBox1SelectedIndexChanged);
         // 
         // bSave
         // 
         this.bSave.Location = new System.Drawing.Point(6, 442);
         this.bSave.Name = "bSave";
         this.bSave.Size = new System.Drawing.Size(75, 31);
         this.bSave.TabIndex = 8;
         this.bSave.Text = "Save";
         this.bSave.UseVisualStyleBackColor = true;
         this.bSave.Click += new System.EventHandler(this.Save_Click);
         // 
         // groupBox1
         // 
         this.groupBox1.Controls.Add(this.tbDaysInTheFuture);
         this.groupBox1.Controls.Add(this.tbDaysInThePast);
         this.groupBox1.Controls.Add(this.label2);
         this.groupBox1.Controls.Add(this.label1);
         this.groupBox1.Location = new System.Drawing.Point(6, 153);
         this.groupBox1.Name = "groupBox1";
         this.groupBox1.Size = new System.Drawing.Size(165, 85);
         this.groupBox1.TabIndex = 0;
         this.groupBox1.TabStop = false;
         this.groupBox1.Text = "Sync Date Range";
         // 
         // tbDaysInTheFuture
         // 
         this.tbDaysInTheFuture.Location = new System.Drawing.Point(112, 51);
         this.tbDaysInTheFuture.Name = "tbDaysInTheFuture";
         this.tbDaysInTheFuture.Size = new System.Drawing.Size(39, 20);
         this.tbDaysInTheFuture.TabIndex = 4;
         this.tbDaysInTheFuture.TextChanged += new System.EventHandler(this.TbDaysInTheFutureTextChanged);
         this.tbDaysInTheFuture.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NumericOnlyKeyPress);
         // 
         // tbDaysInThePast
         // 
         this.tbDaysInThePast.Location = new System.Drawing.Point(112, 21);
         this.tbDaysInThePast.Name = "tbDaysInThePast";
         this.tbDaysInThePast.Size = new System.Drawing.Size(39, 20);
         this.tbDaysInThePast.TabIndex = 3;
         this.tbDaysInThePast.TextChanged += new System.EventHandler(this.TbDaysInThePastTextChanged);
         this.tbDaysInThePast.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NumericOnlyKeyPress);
         // 
         // label2
         // 
         this.label2.Location = new System.Drawing.Point(6, 54);
         this.label2.Name = "label2";
         this.label2.Size = new System.Drawing.Size(100, 23);
         this.label2.TabIndex = 0;
         this.label2.Text = "Days in the Future";
         // 
         // label1
         // 
         this.label1.Location = new System.Drawing.Point(6, 24);
         this.label1.Name = "label1";
         this.label1.Size = new System.Drawing.Size(100, 23);
         this.label1.TabIndex = 0;
         this.label1.Text = "Days in the Past";
         // 
         // tabPage3
         // 
         this.tabPage3.Controls.Add(this.pictureBox1);
         this.tabPage3.Controls.Add(this.linkLabel1);
         this.tabPage3.Controls.Add(this.label4);
         this.tabPage3.Location = new System.Drawing.Point(4, 22);
         this.tabPage3.Name = "tabPage3";
         this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
         this.tabPage3.Size = new System.Drawing.Size(487, 479);
         this.tabPage3.TabIndex = 2;
         this.tabPage3.Text = "About";
         this.tabPage3.UseVisualStyleBackColor = true;
         // 
         // pictureBox1
         // 
         this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
         this.pictureBox1.Location = new System.Drawing.Point(117, 221);
         this.pictureBox1.Name = "pictureBox1";
         this.pictureBox1.Size = new System.Drawing.Size(256, 256);
         this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
         this.pictureBox1.TabIndex = 3;
         this.pictureBox1.TabStop = false;
         // 
         // linkLabel1
         // 
         this.linkLabel1.Location = new System.Drawing.Point(6, 196);
         this.linkLabel1.Name = "linkLabel1";
         this.linkLabel1.Size = new System.Drawing.Size(475, 22);
         this.linkLabel1.TabIndex = 2;
         this.linkLabel1.TabStop = true;
         this.linkLabel1.Text = "https://outlookgooglesyncbidirectional.codeplex.com/";
         this.linkLabel1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
         this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel1LinkClicked);
         // 
         // label4
         // 
         this.label4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
         this.label4.Location = new System.Drawing.Point(3, 32);
         this.label4.Name = "label4";
         this.label4.Size = new System.Drawing.Size(481, 164);
         this.label4.TabIndex = 1;
         this.label4.Text = "OutlookGoogleSync\r\n\r\nVersion {version}\r\n\r\nprogrammed 2012-2013 by\r\nZissis Siantid" +
    "is\r\n\r\nenhanced in 2014 by\r\n~ Bycrobe ~\r\n\r\nenhanced in 2014 by\r\nDeception666";
         this.label4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
         // 
         // notifyIcon1
         // 
         this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
         this.notifyIcon1.Text = "OutlookGoogleSync";
         this.notifyIcon1.Click += new System.EventHandler(this.NotifyIcon1Click);
         // 
         // outlookKeepOpenAfterSync
         // 
         this.outlookKeepOpenAfterSync.AutoSize = true;
         this.outlookKeepOpenAfterSync.Location = new System.Drawing.Point(12, 41);
         this.outlookKeepOpenAfterSync.Name = "outlookKeepOpenAfterSync";
         this.outlookKeepOpenAfterSync.Size = new System.Drawing.Size(132, 17);
         this.outlookKeepOpenAfterSync.TabIndex = 5;
         this.outlookKeepOpenAfterSync.Text = "Keep Open After Sync";
         this.outlookKeepOpenAfterSync.UseVisualStyleBackColor = true;
         this.outlookKeepOpenAfterSync.CheckedChanged += new System.EventHandler(this.outlookKeepOpenAfterSync_CheckedChanged);
         // 
         // MainForm
         // 
         this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
         this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
         this.ClientSize = new System.Drawing.Size(519, 529);
         this.Controls.Add(this.tabControl1);
         this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
         this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
         this.Name = "MainForm";
         this.Text = "OutlookGoogleSync";
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
         this.tabPage3.ResumeLayout(false);
         this.tabPage3.PerformLayout();
         ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
         this.ResumeLayout(false);

      }
      private System.Windows.Forms.CheckBox cbAddReminders;
      private System.Windows.Forms.CheckBox cbAddDescription;
      private System.Windows.Forms.CheckBox cbShowBubbleTooltips;
      private System.Windows.Forms.CheckBox cbSyncEveryHour;
      private System.Windows.Forms.CheckBox cbMinimizeToTray;
      private System.Windows.Forms.CheckBox cbStartInTray;
      private System.Windows.Forms.GroupBox groupBox4;
      private System.Windows.Forms.GroupBox groupBox5;
      private System.Windows.Forms.LinkLabel linkLabel1;
      private System.Windows.Forms.TabPage tabPage3;
      private System.Windows.Forms.TextBox tbMinuteOffsets;
      private System.Windows.Forms.GroupBox groupBox3;
      private System.Windows.Forms.NotifyIcon notifyIcon1;
      private System.Windows.Forms.Label label4;
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
      private System.Windows.Forms.Label lLastSync;
      private System.Windows.Forms.Label lNextSync;
      private System.Windows.Forms.PictureBox pictureBox1;
      private System.Windows.Forms.GroupBox groupBox6;
      private System.Windows.Forms.TextBox outlookAutoLogonTextBox;
      private System.Windows.Forms.CheckBox outlookAutoLogonCheckBox;
      private System.Windows.Forms.Label label6;
      private System.Windows.Forms.Label label5;
      private System.Windows.Forms.TextBox outlookAutoLogonPwdTextBox;
      private System.Windows.Forms.Button clearUserPropertiesBtn;
      private System.Windows.Forms.CheckBox outlookKeepOpenAfterSync;
   }
}

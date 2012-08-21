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
			this.bSyncNow = new System.Windows.Forms.Button();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.tbMinuteOffsets = new System.Windows.Forms.TextBox();
			this.label6 = new System.Windows.Forms.Label();
			this.cbCreateFiles = new System.Windows.Forms.CheckBox();
			this.cbAddAttendees = new System.Windows.Forms.CheckBox();
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
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.label4 = new System.Windows.Forms.Label();
			this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			this.tabPage2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.tabPage3.SuspendLayout();
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
			// LogBox
			// 
			this.LogBox.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.LogBox.Location = new System.Drawing.Point(7, 6);
			this.LogBox.Multiline = true;
			this.LogBox.Name = "LogBox";
			this.LogBox.Size = new System.Drawing.Size(477, 430);
			this.LogBox.TabIndex = 1;
			// 
			// bSyncNow
			// 
			this.bSyncNow.Location = new System.Drawing.Point(4, 442);
			this.bSyncNow.Name = "bSyncNow";
			this.bSyncNow.Size = new System.Drawing.Size(98, 31);
			this.bSyncNow.TabIndex = 0;
			this.bSyncNow.Text = "Sync now";
			this.bSyncNow.UseVisualStyleBackColor = true;
			this.bSyncNow.Click += new System.EventHandler(this.SyncNow_Click);
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.groupBox3);
			this.tabPage2.Controls.Add(this.cbCreateFiles);
			this.tabPage2.Controls.Add(this.cbAddAttendees);
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
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.tbMinuteOffsets);
			this.groupBox3.Controls.Add(this.label6);
			this.groupBox3.Location = new System.Drawing.Point(177, 112);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(304, 85);
			this.groupBox3.TabIndex = 10;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Sync Schedule";
			// 
			// tbMinuteOffsets
			// 
			this.tbMinuteOffsets.Location = new System.Drawing.Point(159, 27);
			this.tbMinuteOffsets.Name = "tbMinuteOffsets";
			this.tbMinuteOffsets.Size = new System.Drawing.Size(139, 20);
			this.tbMinuteOffsets.TabIndex = 5;
			this.tbMinuteOffsets.TextChanged += new System.EventHandler(this.TbMinuteOffsetsTextChanged);
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(6, 30);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(155, 23);
			this.label6.TabIndex = 0;
			this.label6.Text = "Sync at these Minute Offset(s)";
			// 
			// cbCreateFiles
			// 
			this.cbCreateFiles.Location = new System.Drawing.Point(12, 246);
			this.cbCreateFiles.Name = "cbCreateFiles";
			this.cbCreateFiles.Size = new System.Drawing.Size(275, 24);
			this.cbCreateFiles.TabIndex = 7;
			this.cbCreateFiles.Text = "Create text files with found/identified entries";
			this.cbCreateFiles.UseVisualStyleBackColor = true;
			this.cbCreateFiles.CheckedChanged += new System.EventHandler(this.CheckBox2CheckedChanged);
			// 
			// cbAddAttendees
			// 
			this.cbAddAttendees.Location = new System.Drawing.Point(12, 216);
			this.cbAddAttendees.Name = "cbAddAttendees";
			this.cbAddAttendees.Size = new System.Drawing.Size(286, 24);
			this.cbAddAttendees.TabIndex = 6;
			this.cbAddAttendees.Text = "Add Attendees at the end of the description";
			this.cbAddAttendees.UseVisualStyleBackColor = true;
			this.cbAddAttendees.CheckedChanged += new System.EventHandler(this.CheckBox1CheckedChanged);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.bGetMyCalendars);
			this.groupBox2.Controls.Add(this.cbCalendars);
			this.groupBox2.Location = new System.Drawing.Point(6, 26);
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
			this.bGetMyCalendars.Location = new System.Drawing.Point(355, 19);
			this.bGetMyCalendars.Name = "bGetMyCalendars";
			this.bGetMyCalendars.Size = new System.Drawing.Size(114, 40);
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
			this.groupBox1.Location = new System.Drawing.Point(6, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(165, 85);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Sync Date Range";
			// 
			// tbDaysInTheFuture
			// 
			this.tbDaysInTheFuture.Location = new System.Drawing.Point(112, 50);
			this.tbDaysInTheFuture.Name = "tbDaysInTheFuture";
			this.tbDaysInTheFuture.Size = new System.Drawing.Size(39, 20);
			this.tbDaysInTheFuture.TabIndex = 4;
			this.tbDaysInTheFuture.TextChanged += new System.EventHandler(this.TbDaysInTheFutureTextChanged);
			// 
			// tbDaysInThePast
			// 
			this.tbDaysInThePast.Location = new System.Drawing.Point(112, 27);
			this.tbDaysInThePast.Name = "tbDaysInThePast";
			this.tbDaysInThePast.Size = new System.Drawing.Size(39, 20);
			this.tbDaysInThePast.TabIndex = 3;
			this.tbDaysInThePast.TextChanged += new System.EventHandler(this.TbDaysInThePastTextChanged);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(6, 53);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(100, 23);
			this.label2.TabIndex = 0;
			this.label2.Text = "Days in the Future";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(6, 30);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(100, 23);
			this.label1.TabIndex = 0;
			this.label1.Text = "Days in the Past";
			// 
			// tabPage3
			// 
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
			// linkLabel1
			// 
			this.linkLabel1.Location = new System.Drawing.Point(6, 138);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(475, 23);
			this.linkLabel1.TabIndex = 2;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "http://outlookgooglesync.codeplex.com/";
			this.linkLabel1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabel1LinkClicked);
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(3, 32);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(481, 96);
			this.label4.TabIndex = 1;
			this.label4.Text = "OutlookGoogleSync\r\n\r\nVersion {version}\r\n\r\nprogrammed 2012 by\r\nZissis Siantidis\r\n";
			this.label4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// notifyIcon1
			// 
			this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
			this.notifyIcon1.Text = "OutlookGoogleSync";
			this.notifyIcon1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.NotifyIcon1MouseDoubleClick);
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
			this.groupBox3.ResumeLayout(false);
			this.groupBox3.PerformLayout();
			this.groupBox2.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.tabPage3.ResumeLayout(false);
			this.ResumeLayout(false);
		}
		private System.Windows.Forms.LinkLabel linkLabel1;
		private System.Windows.Forms.TabPage tabPage3;
		private System.Windows.Forms.Label label6;
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
	}
}

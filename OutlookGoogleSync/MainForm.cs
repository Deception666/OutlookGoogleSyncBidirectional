//TODO: consider description updates?
//TODO: optimize comparison algorithms
using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Net;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleSync
{
   /// <summary>
   /// Description of MainForm.
   /// </summary>
   public partial class MainForm : Form
   {
      public static MainForm Instance;

      public const string FILENAME = "settings.xml";
      public const string VERSION = "1.1.0";

      public Timer ogstimer;
      public DateTime oldtime;
      public List<int> MinuteOffsets = new List<int>();
      DateTime lastSyncDate;
      int currentTimerInterval = 0;

      public MainForm()
      {
         InitializeComponent();
         label4.Text = label4.Text.Replace("{version}", VERSION);

         Instance = this;

         //set system proxy
         WebProxy wp = (WebProxy)System.Net.GlobalProxySelection.Select;
         wp.UseDefaultCredentials = true;
         System.Net.WebRequest.DefaultWebProxy = wp;

         //load settings/create settings file
         if (File.Exists(FILENAME))
         {
            Settings.Instance = XMLManager.import<Settings>(FILENAME);
         }
         else
         {
            XMLManager.export(Settings.Instance, FILENAME);
         }

         //create the timer for the autosynchro
         ogstimer = new Timer();
         ogstimer.Tick += new EventHandler(ogstimer_Tick);

         //update GUI from Settings
         tbDaysInThePast.Text = Settings.Instance.DaysInThePast.ToString();
         tbDaysInTheFuture.Text = Settings.Instance.DaysInTheFuture.ToString();
         tbMinuteOffsets.Text = Settings.Instance.MinuteOffsets;
         lastSyncDate = Settings.Instance.LastSyncDate;
         cbCalendars.Items.Add(Settings.Instance.UseGoogleCalendar);
         cbCalendars.SelectedIndex = 0;
         cbSyncEveryHour.Checked = Settings.Instance.SyncEveryHour;
         cbShowBubbleTooltips.Checked = Settings.Instance.ShowBubbleTooltipWhenSyncing;
         cbStartInTray.Checked = Settings.Instance.StartInTray;
         cbMinimizeToTray.Checked = Settings.Instance.MinimizeToTray;
         cbAddDescription.Checked = Settings.Instance.AddDescription;
         cbAddAttendees.Checked = Settings.Instance.AddAttendeesToDescription;
         cbAddReminders.Checked = Settings.Instance.AddReminders;
         cbCreateFiles.Checked = Settings.Instance.CreateTextFiles;

         // init the calendar service if the user is populated
         if (Settings.Instance.UseGoogleCalendar.User != "")
         {
            if (!GoogleCalendar.Instance.InitCalendarService(Settings.Instance.UseGoogleCalendar.User))
            {
               cbCalendars.Items.Clear();

               logboxout("Unable to initialize Google calendar service for the following Google user: " +
                         Settings.Instance.UseGoogleCalendar.User);
            }
            else
            {
               logboxout("Initializing Google calendar service for the following Google user: " +
                         Settings.Instance.UseGoogleCalendar.User + " (" +
                         Settings.Instance.UseGoogleCalendar.Name + ")");
            }
         }

         //Start in tray?
         if (cbStartInTray.Checked)
         {
            this.WindowState = FormWindowState.Minimized;
            notifyIcon1.Visible = true;
            this.Hide();
            this.ShowInTaskbar = false;
         }

         //set up tooltips for some controls
         ToolTip toolTip1 = new ToolTip();
         toolTip1.AutoPopDelay = 10000;
         toolTip1.InitialDelay = 500;
         toolTip1.ReshowDelay = 200;
         toolTip1.ShowAlways = true;
         toolTip1.SetToolTip(cbCalendars, "The Google Calendar to synchonize with.");
         toolTip1.SetToolTip(tbMinuteOffsets,
             "One ore more Minute Offsets at which the sync is automatically started each hour. \n" +
             "Separate by comma (e.g. 5,15,25).");
         toolTip1.SetToolTip(cbAddAttendees,
             "While Outlook has fields for Organizer, RequiredAttendees and OptionalAttendees, Google has not.\n" +
             "If checked, this data is added at the end of the description as text.");
         toolTip1.SetToolTip(cbAddReminders,
             "If checked, the reminder set in outlook will be carried over to the Google Calendar entry (as a popup reminder).");
         toolTip1.SetToolTip(cbCreateFiles,
             "If checked, all entries found in Outlook/Google and identified for creation/deletion will be exported \n" +
             "to 4 separate text files in the application's directory (named \"export_*.txt\"). \n" +
             "Only for debug/diagnostic purposes.");
         toolTip1.SetToolTip(cbAddDescription,
             "The description may contain email addresses, which Outlook may complain about (PopUp-Message: \"Allow Access?\" etc.). \n" +
             "Turning this off allows OutlookGoogleSync to run without intervention in this case.");

         //Refresh synchronizations (last and next)
         lLastSync.Text = "Last succeded synchro:\n     " + lastSyncDate.ToLongDateString() + " - " + lastSyncDate.ToLongTimeString();
         setNextSync(getResyncInterval());
      }

      int getResyncInterval()
      {
         int min = 0;
         int.TryParse(tbMinuteOffsets.Text, out min);
         if (min < 1) { min = 60; }
         return min;
      }

      void ogstimer_Tick(object sender, EventArgs e)
      {
         SyncNow_Click(null, null);
      }

      void setNextSync(int delay)
      {
         if (cbSyncEveryHour.Checked)
         {
            DateTime nextSyncDate = lastSyncDate.AddMinutes(delay);
            if (currentTimerInterval != delay)
            {
               ogstimer.Stop();
               DateTime now = DateTime.Now;
               TimeSpan diff = nextSyncDate - now;
               currentTimerInterval = diff.Minutes;
               if (currentTimerInterval < 1) { currentTimerInterval = 1; nextSyncDate = now.AddMinutes(currentTimerInterval); }
               ogstimer.Interval = currentTimerInterval * 60000;
               ogstimer.Start();
            }
            lNextSync.Text = "Next scheduled sync:\n     " + nextSyncDate.ToLongDateString() + " - " + nextSyncDate.ToLongTimeString();
         }
         else
         {
            lNextSync.Text = "Next scheduled sync:\n     Inactive";
         }
      }

      void GetMyGoogleCalendars_Click(object sender, EventArgs e)
      {
         bGetMyCalendars.Enabled = false;
         cbCalendars.Enabled = false;
         List<MyCalendarListEntry> calendars = null;

         UserAccountForm user_account_form = new UserAccountForm(Settings.Instance.UseGoogleCalendar.User);

         if (user_account_form.ShowDialog(this) == DialogResult.OK)
         {
            try
            {
               if (!GoogleCalendar.Instance.InitCalendarService(user_account_form.UserAccount))
               {
                  cbCalendars.Items.Clear();

                  logboxout("Unable to initialize Google calendar service for the following Google user: " + user_account_form.UserAccount);
               }
               else
               {
                  logboxout("Initializing Google calendar service for the following Google user: " + user_account_form.UserAccount);

                  calendars = GoogleCalendar.Instance.getCalendars();
               }
            }
            catch (System.Exception ex)
            {
               logboxout("Unable to get the list of Google Calendars. The folowing error occurs:");
               logboxout(ex.Message + "\r\n => Check your network connection.");
            }
            if (calendars != null)
            {
               cbCalendars.Items.Clear();
               foreach (MyCalendarListEntry mcle in calendars)
               {
                  cbCalendars.Items.Add(mcle);
               }
               MainForm.Instance.cbCalendars.SelectedIndex = 0;
            }

            cbCalendars.Enabled = true;
         }

         bGetMyCalendars.Enabled = true;
      }

      void SyncNow_Click(object sender, EventArgs e)
      {
         // update the property key value before doing the sync
         UpdateEventPropertyKey();

         bSyncNow.Enabled = false;

         lNextSync.Text = "Next scheduled sync:\n     In progress...";

         LogBox.Clear();

         DateTime SyncStarted = DateTime.Now;

         logboxout("Sync started at " + SyncStarted.ToString());
         logboxout("--------------------------------------------------");

         Boolean syncOk = synchronize();
         logboxout("--------------------------------------------------");
         logboxout(syncOk ? "Sync finished with success !" : "Operation aborted !");

         if (syncOk)
         {
            lastSyncDate = SyncStarted;
            Settings.Instance.LastSyncDate = lastSyncDate;
            XMLManager.export(Settings.Instance, FILENAME);
            lLastSync.Text = "Last succeded synchro:\n     " + SyncStarted.ToLongDateString() + " - " + SyncStarted.ToLongTimeString();
            setNextSync(getResyncInterval());
         }
         else
         {
            setNextSync(5);
         }

         if (!cbSyncEveryHour.Checked)
         {
            // close the outlook calendar instance so not to always be logged in
            OutlookCalendar.Instance.Release();
         }

         bSyncNow.Enabled = true;
      }

      void synchronize( List< AppointmentItem > outlook_items, List< Event > google_items )
      {
         // indicates the number of entries added / updated / removed....
         uint google_entries_added = 0;
         uint google_entries_updated = 0;
         uint google_entries_removed = 0;
         uint outlook_entries_added = 0;
         uint outlook_entries_updated = 0;
         uint outlook_entries_removed = 0;

         // first synchronize outlook -> google

         // run over all the office items and add or update on google
         for (int i = outlook_items.Count - 1; i >= 0; --i)
         {
            // obtain a reference to the outlook item...
            var oitem = outlook_items[i];

            // determine first if this outlook item is associated with a google calendar item...
            UserProperty oitem_google_prop = oitem.UserProperties.Find(EventPropertyKey);

            if (oitem_google_prop == null)
            {
               // give some indication of what will take place
               logboxout("Creating Google event: " + oitem.Subject);

               // the property does not exist... this mean that google calendar
               // does not have this outlook entry...  after this call, the outlook item
               // and the google item should be tied together by the use of properties...
               var gitem = GoogleCalendar.Instance.addEntry(oitem, cbAddDescription.Checked, cbAddReminders.Checked, cbAddAttendees.Checked);

               // updated the stats
               ++google_entries_added;
            }
            else
            {
               // the property exists, so we need to determine if the event should be updated
               // first, find the event in the list of google items...
               Event gitem = null;
               foreach (var g in google_items)
               {
                  if (oitem_google_prop.Value == g.Id)
                  {
                     gitem = g; break;
                  }
               }

               if (gitem == null)
               {
                  // give some indication of what will take place
                  logboxout("Removing Outlook event: " + oitem.Subject);

                  // the item does not exist, so it was removed from google calendar
                  // since it was removed from google, remove it from outlook
                  OutlookCalendar.Instance.deleteCalendarEntry(oitem);

                  // outlook item can no longer be used after it is deleted
                  outlook_items.RemoveAt(i);

                  // update the stats
                  ++outlook_entries_removed;
               }
               else
               {
                  // the item does exist...
                  // determine if the event should be updated...
                  if (signature(oitem) != signature(gitem) && oitem.LastModificationTime > gitem.Updated)
                  {
                     // give some indication of what will take place
                     logboxout("Updating Google event: " + gitem.Summary);

                     // update the event based on the outlook item
                     GoogleCalendar.Instance.updateEntry(gitem, oitem, cbAddDescription.Checked, cbAddReminders.Checked, cbAddAttendees.Checked);

                     // this google item has been processed
                     google_items.Remove(gitem);

                     // update the status
                     ++google_entries_updated;
                  }
               }
            }
         }

         // second synchronize google -> outlook

         // run over all the google items and add or update on outlook
         for (int i = google_items.Count - 1; i >= 0; --i)
         {
            // obtain a reference to the google item...
            var gitem = google_items[i];

            // determine first if this google item is associated with an outlook calendar item...
            string outlook_id = null;
            if (gitem.ExtendedProperties != null &&
                gitem.ExtendedProperties.Private != null &&
                gitem.ExtendedProperties.Private.ContainsKey(EventPropertyKey))
            {
               outlook_id = gitem.ExtendedProperties.Private[EventPropertyKey];
            }

            if (outlook_id == null)
            {
               // give some indication of what will take place
               logboxout("Creating Outlook event: " + gitem.Summary);

               // the property does not exist... this means that the outlook calendar
               // does not have this google entry...  after this call, the outlook item
               // and the google item should be tied together by the use of properties...
               var oitem = OutlookCalendar.Instance.addEntry(gitem, cbAddDescription.Checked, cbAddReminders.Checked, cbAddAttendees.Checked);

               // the google currently does not have an updated id... this needs to be reflected on the server...
               GoogleCalendar.Instance.updateEntry(gitem);

               // update the stats
               ++outlook_entries_added;
            }
            else
            {
               // the property exists, so we need to determine if the event should be updated
               // first, find the event in the list of outlook items...
               AppointmentItem oitem = null;
               foreach (var o in outlook_items)
               {
                  if (outlook_id == o.GlobalAppointmentID)
                  {
                     oitem = o; break;
                  }
               }

               if (oitem == null)
               {
                  // give some indication of what will take place
                  logboxout("Removing Google event: " + gitem.Summary);

                  // the item does not exist, so it was removed from outlook calendar
                  // since it was removed from outlook, remove it from google
                  GoogleCalendar.Instance.deleteCalendarEntry(gitem);

                  // update the stats
                  ++google_entries_removed;
               }
               else
               {
                  // this outlook item is about to be processed... no matter the action
                  // taken, we can remove it from the list of items, so as not
                  // to need to look at it again in future iterations...
                  outlook_items.Remove(oitem);

                  // the item does exist...
                  // determine if the event should be updated...
                  if (signature(gitem) != signature(oitem) && gitem.Updated > oitem.LastModificationTime)
                  {
                     // give some indication of what will take place
                     logboxout("Updating Outlook event: " + oitem.Subject);

                     // update the event based on the google item
                     OutlookCalendar.Instance.updateEntry(oitem, gitem, cbAddDescription.Checked, cbAddReminders.Checked, cbAddAttendees.Checked);

                     // update the status
                     ++outlook_entries_updated;
                  }
               }
            }
         }

         // clear out both lists... these items have been processed...
         outlook_items.Clear();
         google_items.Clear();

         // summarize the changes made
         logboxout("--------------------------------------------------");
         logboxout("Google entries added: " + google_entries_added);
         logboxout("Google entries updated: " + google_entries_updated);
         logboxout("Google entries removed: " + google_entries_removed);
         logboxout("");
         logboxout("Outlook entries added: " + outlook_entries_added);
         logboxout("Outlook entries updated: " + outlook_entries_updated);
         logboxout("Outlook entries removed: " + outlook_entries_removed);
      }

      Boolean synchronize()
      {
         if (Settings.Instance.UseGoogleCalendar.Id == "")
         {
            MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
            return false;
         }

         logboxout("Reading Outlook Calendar Entries...");
         List<AppointmentItem> OutlookEntries = null;
         try
         {
            OutlookEntries = OutlookCalendar.Instance.getCalendarEntriesInRange();
         }
         catch (System.Exception ex)
         {
            logboxout("Unable to access to the Outlook Calendar. The folowing error occurs:");
            logboxout(ex.Message + "\r\n => Retry later.");
            OutlookCalendar.Instance.Reset();
            return false;
         }
         if (cbCreateFiles.Checked)
         {
            TextWriter tw = new StreamWriter("export_found_in_outlook.txt");
            foreach (AppointmentItem ai in OutlookEntries)
            {
               tw.WriteLine(signature(ai));
            }
            tw.Close();
         }
         logboxout("Found " + OutlookEntries.Count + " Outlook Calendar Entries.");
         logboxout("--------------------------------------------------");



         logboxout("Reading Google Calendar Entries...");
         List<Event> GoogleEntries = null;
         try
         {
            GoogleEntries = GoogleCalendar.Instance.getCalendarEntriesInRange();
         }
         catch (System.Exception ex)
         {
            logboxout("Unable to connect to the Google Calendar. The folowing error occurs:");
            logboxout(ex.Message + "\r\n => Check your network connection.");
            return false;
         }

         if (cbCreateFiles.Checked)
         {
            TextWriter tw = new StreamWriter("export_found_in_google.txt");
            foreach (Event ev in GoogleEntries)
            {
               tw.WriteLine(signature(ev));
            }
            tw.Close();
         }
         logboxout("Found " + GoogleEntries.Count + " Google Calendar Entries.");
         logboxout("--------------------------------------------------");

         // synchronize both outlook and google calendars
         synchronize(OutlookEntries, GoogleEntries);

         return true;
      }

      //creates a standardized summary string with the key attributes of a calendar entry for comparison
      public string signature(AppointmentItem ai)
      {
         return (ai.Start + ";" + ai.End + ";" + ai.Subject + ";" + ai.Location).Trim();
      }
      public string signature(Event ev)
      {
         string start_time = ev.Start.Date != null ?
                             DateTime.Parse(ev.Start.Date).ToString() :
                             ev.Start.DateTime.ToString();

         string end_time = ev.End.Date != null ?
                           DateTime.Parse(ev.End.Date).ToString() :
                           ev.End.DateTime.ToString();

         return (start_time + ";" + end_time + ";" + ev.Summary + ";" + ev.Location).Trim();
      }

      void logboxout(string s)
      {
         LogBox.AppendText(s + Environment.NewLine);
      }

      void Save_Click(object sender, EventArgs e)
      {
         XMLManager.export(Settings.Instance, FILENAME);
      }

      void ComboBox1SelectedIndexChanged(object sender, EventArgs e)
      {
         Settings.Instance.UseGoogleCalendar = (MyCalendarListEntry)cbCalendars.SelectedItem;
      }

      void TbDaysInThePastTextChanged(object sender, EventArgs e)
      {
         Settings.Instance.DaysInThePast = int.Parse(tbDaysInThePast.Text);
      }

      void TbDaysInTheFutureTextChanged(object sender, EventArgs e)
      {
         Settings.Instance.DaysInTheFuture = int.Parse(tbDaysInTheFuture.Text);
      }

      void TbMinuteOffsetsTextChanged(object sender, EventArgs e)
      {
         Settings.Instance.MinuteOffsets = tbMinuteOffsets.Text;
         setNextSync(getResyncInterval());
      }

      void CbSyncEveryHourCheckedChanged(object sender, System.EventArgs e)
      {
         Settings.Instance.SyncEveryHour = cbSyncEveryHour.Checked;
         setNextSync(getResyncInterval());
      }

      void CbShowBubbleTooltipsCheckedChanged(object sender, System.EventArgs e)
      {
         Settings.Instance.ShowBubbleTooltipWhenSyncing = cbShowBubbleTooltips.Checked;
      }

      void CbStartInTrayCheckedChanged(object sender, System.EventArgs e)
      {
         Settings.Instance.StartInTray = cbStartInTray.Checked;
      }

      void CbMinimizeToTrayCheckedChanged(object sender, System.EventArgs e)
      {
         Settings.Instance.MinimizeToTray = cbMinimizeToTray.Checked;
      }

      void CbAddDescriptionCheckedChanged(object sender, EventArgs e)
      {
         Settings.Instance.AddDescription = cbAddDescription.Checked;
      }

      void CbAddRemindersCheckedChanged(object sender, EventArgs e)
      {
         Settings.Instance.AddReminders = cbAddReminders.Checked;
      }

      void cbAddAttendees_CheckedChanged(object sender, EventArgs e)
      {
         Settings.Instance.AddAttendeesToDescription = cbAddAttendees.Checked;
      }

      void cbCreateFiles_CheckedChanged(object sender, EventArgs e)
      {
         Settings.Instance.CreateTextFiles = cbCreateFiles.Checked;
      }

      void NotifyIcon1Click(object sender, EventArgs e)
      {
         this.Show();
         this.WindowState = FormWindowState.Normal;
      }

      void MainFormResize(object sender, EventArgs e)
      {
         if (!cbMinimizeToTray.Checked) return;
         if (this.WindowState == FormWindowState.Minimized)
         {
            notifyIcon1.Visible = true;
            this.Hide();
            this.ShowInTaskbar = false;
         }
         else if (this.WindowState == FormWindowState.Normal)
         {
            notifyIcon1.Visible = false;
            this.Show();
            this.ShowInTaskbar = true;
         }
      }

      public void HandleException(System.Exception ex)
      {
         MessageBox.Show(ex.ToString(), "Exception!", MessageBoxButtons.OK, MessageBoxIcon.Error);
         TextWriter tw = new StreamWriter("exception.txt");
         tw.WriteLine(ex.ToString());
         tw.Close();

         this.Close();
         System.Environment.Exit(-1);
         System.Windows.Forms.Application.Exit();
      }

      void LinkLabel1LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
      {
         System.Diagnostics.Process.Start(linkLabel1.Text);
      }

      private string EventPropertyKey { get; set; }

      private void UpdateEventPropertyKey( )
      {
         if (Settings.Instance.UseGoogleCalendar.Id == null)
         {
            throw new System.Exception("Unable to set EventPropertyKey!!!  Obtain Google calendars first!!!");
         }
         else
         {
            // the string cannot have '[', ']', '_', or '#'
            string key = Settings.Instance.UseGoogleCalendar.Id;
            key = key.Replace("[", "").Replace("]", "").Replace("_", "").Replace("#", "");

            // the string cannot exceed 45 characters in length (google restriction)
            key = key.Substring(0, key.Length > 44 ? 44 : key.Length);

            // set the event properties for all parties
            EventPropertyKey = key;
            GoogleCalendar.Instance.EventPropertyKey = key;
            OutlookCalendar.Instance.EventPropertyKey = key;
         }
      }
   }
}

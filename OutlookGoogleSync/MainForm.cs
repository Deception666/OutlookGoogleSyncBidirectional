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

      public static readonly string LOCAL_APP_DATA = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData) + "\\OutlookGoogleSync";
      public static readonly string FILENAME = LOCAL_APP_DATA + "\\settings.xml";
      public static readonly string VERSION = System.Windows.Forms.Application.ProductVersion;

      public Timer ogstimer;
      public DateTime oldtime;
      public List<int> MinuteOffsets = new List<int>();
      DateTime lastSyncDate;
      int currentTimerInterval = 0;

      public MainForm()
      {
         if (System.IO.Directory.Exists(LOCAL_APP_DATA) == false)
         {
            System.IO.Directory.CreateDirectory(LOCAL_APP_DATA);
         }

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
         outlookAutoLogonCheckBox.Checked = Settings.Instance.OutlookAutoLogonEnabled;
         outlookAutoLogonTextBox.Text = Settings.Instance.OutlookAutoLogonProfileName;
         outlookAutoLogonPwdTextBox.Text = Settings.Instance.OutlookAutoLogonProfilePassword;

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
                  logboxout("Initializing Google calendar service for the following Google user: " +
                             Settings.Instance.UseGoogleCalendar.User + " (" +
                             Settings.Instance.UseGoogleCalendar.Name + ")");

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

      bool ObtainOutlookEntries( out List< AppointmentItem > outlook_items )
      {
         bool obtained = true;
         outlook_items = new List< AppointmentItem >();

         try
         {
            outlook_items = OutlookCalendar.Instance.getCalendarEntriesInRange();
         }
         catch (System.Exception ex)
         {
            logboxout("Unable to access to the Outlook Calendar. The folowing error occurs:");
            logboxout(ex.Message + Environment.NewLine + "=> Retry later.");

            OutlookCalendar.Instance.Release();
            obtained = false;
         }

         return obtained;
      }

      bool ObtainGoogleEntries( out List< Event > google_items )
      {
         bool obtained = true;
         google_items = new List< Event >();

         try
         {
            google_items = GoogleCalendar.Instance.getCalendarEntriesInRange();
         }
         catch (System.Exception ex)
         {
            logboxout("Unable to connect to the Google Calendar. The folowing error occurs:");
            logboxout(ex.Message + Environment.NewLine + " => Check your network connection.");

            obtained = false;
         }

         return obtained;
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

         // close the outlook calendar instance so not to always be logged in
         OutlookCalendar.Instance.Release();

         bSyncNow.Enabled = true;
      }

      void synchronize( List< AppointmentItem > outlook_items, List< Event > google_items )
      {
         // TODO: outlook to google recurrence 90% complete
         //       getting weird error message from outlook: "The operation cannot be performed because the message has changed"
         //       need to determine what happens when deleting just the master recurrence object
         // TODO: google to outlook recurrence 0% complete... need to read up on http://www.ietf.org/rfc/rfc2445
         // TODO: testing on a larger scale...
         // TODO: refactor this function, as it is getting out of control...

         // indicates the number of entries added / updated / removed....
         uint google_entries_added = 0;
         uint google_entries_updated = 0;
         uint google_entries_removed = 0;
         uint outlook_entries_added = 0;
         uint outlook_entries_updated = 0;
         uint outlook_entries_removed = 0;
         uint bound_entries_found = 0;

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
               // run across the google events to see if there is one that matches
               Event gitem = null;
               foreach (var g in google_items)
               {
                  // lowercase match the signature and
                  // the property is not set or the property matches the oitem, match found...
                  // the user has to type it the same in both cases to make a match...
                  // is there something better here?
                  if (signature(g).ToLower() == signature(oitem).ToLower() &&
                     (g.ExtendedProperties == null ||
                      g.ExtendedProperties.Private == null ||
                      g.ExtendedProperties.Private.ContainsKey(EventPropertyKey) == false ||
                      g.ExtendedProperties.Private[EventPropertyKey] == OutlookCalendar.FormatEventID(oitem)))
                  {
                     gitem = g; break;
                  }
               }

               if (gitem != null)
               {
                  // give some indication of what will take place
                  logboxout("Binding Outlook and Google event: " + oitem.Subject + " (" + oitem.Start + ")");

                  // bind the properties together
                  OutlookCalendar.Instance.Bind(oitem, gitem);

                  // update the instance on the google calendar
                  GoogleCalendar.Instance.updateEntry(gitem);

                  // remove the outlook item and the google item from the lists...
                  // do not need to process them again for future iterations...
                  google_items.Remove(gitem);
                  outlook_items.Remove(oitem);

                  // save and close the outlook item
                  ((_AppointmentItem)oitem).Close(OlInspectorClose.olSave);

                  // remove the com reference
                  System.Runtime.InteropServices.Marshal.ReleaseComObject(oitem);

                  // update the stats
                  ++bound_entries_found;  
               }
               else
               {
                  // give some indication of what will take place
                  logboxout("Creating Google event: " + oitem.Subject);

                  // the property does not exist... this mean that google calendar
                  // does not have this outlook entry...  after this call, the outlook item
                  // and the google item should be tied together by the use of properties...
                  gitem = GoogleCalendar.Instance.addEntry(oitem, cbAddDescription.Checked, cbAddReminders.Checked, cbAddAttendees.Checked);

                  // updated the stats
                  ++google_entries_added;
               }
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

                  // need to release the com reference
                  System.Runtime.InteropServices.Marshal.ReleaseComObject(oitem);

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

               // need to release the com reference
               System.Runtime.InteropServices.Marshal.ReleaseComObject(oitem_google_prop);
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
               // run across the outlook events to see if there is one that matches...
               // this may not need to be done, as the google pass should have found them all...
               // lets be on the safe side here and do it anyway...
               AppointmentItem oitem = null;
               foreach (var o in outlook_items)
               {
                  // lowercase match the signature and
                  // the property is not set or the property matches the gitem, match found...
                  // the user has to type it the same in both cases to make a match...
                  // is there something better here?
                  if (signature(o).ToLower() == signature(gitem).ToLower() &&
                     (oitem.UserProperties == null ||
                      oitem.UserProperties.Find(EventPropertyKey) == null ||
                      oitem.UserProperties.Find(EventPropertyKey).Value == gitem.Id))
                  {
                     oitem = o; break;
                  }
               }

               if (oitem != null)
               {
                  // give some indication of what will take place
                  logboxout("Binding Outlook and Google event: " + gitem.Summary + " (" + gitem.Start + ")");

                  // bind the properties together
                  OutlookCalendar.Instance.Bind(oitem, gitem);

                  // update the instance on the google calendar
                  GoogleCalendar.Instance.updateEntry(gitem);

                  // remove the outlook item and the google item from the lists...
                  // do not need to process them again for future iterations...
                  google_items.Remove(gitem);
                  outlook_items.Remove(oitem);

                  // save and close the outlook item
                  ((_AppointmentItem)oitem).Close(OlInspectorClose.olSave);

                  // need to release the com reference
                  System.Runtime.InteropServices.Marshal.ReleaseComObject(oitem);

                  // update the stats
                  ++bound_entries_found;  
               }
               else
               {
                  // give some indication of what will take place
                  logboxout("Creating Outlook event: " + gitem.Summary);

                  // the property does not exist... this means that the outlook calendar
                  // does not have this google entry...  after this call, the outlook item
                  // and the google item should be tied together by the use of properties...
                  oitem = OutlookCalendar.Instance.addEntry(gitem, cbAddDescription.Checked, cbAddReminders.Checked, cbAddAttendees.Checked);

                  // google currently does not have an updated id... this needs to be reflected on the server...
                  GoogleCalendar.Instance.updateEntry(gitem);

                  // need to release the com reference
                  System.Runtime.InteropServices.Marshal.ReleaseComObject(oitem);

                  // update the stats
                  ++outlook_entries_added;
               }
            }
            else
            {
               // the property exists, so we need to determine if the event should be updated
               // first, find the event in the list of outlook items...
               AppointmentItem oitem = null;
               foreach (var o in outlook_items)
               {
                  if (outlook_id == OutlookCalendar.FormatEventID(o))
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

                  // save and close the outlook item
                  ((_AppointmentItem)oitem).Close(OlInspectorClose.olSave);

                  // need to release the com reference
                  System.Runtime.InteropServices.Marshal.ReleaseComObject(oitem);
               }
            }
         }

         // close all the outlook items
         foreach (var o in outlook_items)
         {
            ((_AppointmentItem)o).Close(OlInspectorClose.olSave);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
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
         logboxout("");
         logboxout("Bound entries found: " + bound_entries_found);
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
         if (!ObtainOutlookEntries(out OutlookEntries)) return false;
         
         if (cbCreateFiles.Checked)
         {
            TextWriter tw = new StreamWriter(LOCAL_APP_DATA + "\\export_found_in_outlook.txt");
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
         if (!ObtainGoogleEntries(out GoogleEntries)) return false;

         if (cbCreateFiles.Checked)
         {
            TextWriter tw = new StreamWriter(LOCAL_APP_DATA + "\\export_found_in_google.txt");
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
         Settings.Instance.OutlookAutoLogonProfileName = outlookAutoLogonTextBox.Text;
         Settings.Instance.OutlookAutoLogonProfilePassword = outlookAutoLogonPwdTextBox.Text;

         XMLManager.export(Settings.Instance, FILENAME);
      }

      void ComboBox1SelectedIndexChanged(object sender, EventArgs e)
      {
         Settings.Instance.UseGoogleCalendar = (MyCalendarListEntry)cbCalendars.SelectedItem;
      }

      void TbDaysInThePastTextChanged(object sender, EventArgs e)
      {
         Settings.Instance.DaysInThePast = 0;

         if (tbDaysInThePast.Text.Length != 0)
         {
            Settings.Instance.DaysInThePast = int.Parse(tbDaysInThePast.Text);
         }
      }

      void TbDaysInTheFutureTextChanged(object sender, EventArgs e)
      {
         Settings.Instance.DaysInTheFuture = 0;

         if (tbDaysInTheFuture.Text.Length != 0)
         {
            Settings.Instance.DaysInTheFuture = int.Parse(tbDaysInTheFuture.Text);
         }
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
         TextWriter tw = new StreamWriter(LOCAL_APP_DATA + "\\exception.txt");
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

            // the string cannot exceed 44 characters in length (google restriction)
            key = key.Substring(0, key.Length > 44 ? 44 : key.Length);

            // set the event properties for all parties
            EventPropertyKey = key;
            GoogleCalendar.Instance.EventPropertyKey = key;
            OutlookCalendar.Instance.EventPropertyKey = key;
         }
      }

      private void outlookAutoLogonCheckBox_CheckedChanged(object sender, EventArgs e)
      {
         if (outlookAutoLogonCheckBox.Checked)
         {
            outlookAutoLogonTextBox.ReadOnly = false;
            outlookAutoLogonPwdTextBox.ReadOnly = false;
            Settings.Instance.OutlookAutoLogonEnabled = true;
            if (OutlookCalendar.IsLoggedIn()) OutlookCalendar.Instance.Release();
         }
         else
         {
            outlookAutoLogonTextBox.ReadOnly = true;
            outlookAutoLogonPwdTextBox.ReadOnly = true;
            Settings.Instance.OutlookAutoLogonEnabled = false;
            if (OutlookCalendar.IsLoggedIn()) OutlookCalendar.Instance.Release();
         }

      }

      private void clearUserPropertiesBtn_Click(object sender, EventArgs ea)
      {
         // do not allow user to press buttons
         bSyncNow.Enabled = false;
         clearUserPropertiesBtn.Enabled = false;

         // this is a developer action... it may be required to start fresh, so clear out the bindings...
         var result = MessageBox.Show(this,
                                      "Do you want to clear the bindings between Outlook and Google events?",
                                      "Clear Bindings (Developer Action)",
                                      MessageBoxButtons.YesNo);

         if (result == DialogResult.Yes)
         {
            // indicate start
            LogBox.Text = "";
            logboxout("Clear properties started at " + DateTime.Now);
            logboxout("");

            // update the property key value before doing the sync
            UpdateEventPropertyKey();

            // obtain the outlook items
            List< AppointmentItem > outlook_items = null;
            if (!ObtainOutlookEntries(out outlook_items)) return;

            // remove the user property for the currently logged in outlook user...
            foreach (var o in outlook_items)
            {
               if (o != null && o.UserProperties != null)
               {
                  UserProperty oitem_property = o.UserProperties.Find(OutlookCalendar.Instance.EventPropertyKey);

                  if (oitem_property != null)
                  {
                     logboxout("Clearing Outlook property: " + o.Subject);

                     oitem_property.Delete();

                     System.Runtime.InteropServices.Marshal.ReleaseComObject(oitem_property);
                  }

                  ((_AppointmentItem)o).Close(OlInspectorClose.olSave);

                  System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
               }
            }

            // obtain the google items
            List< Event > google_items = null;
            if (!ObtainGoogleEntries(out google_items)) return;

            // remove the user property for the currently logged in google user...
            foreach (var e in google_items)
            {
               if (e != null &&
                   e.ExtendedProperties != null &&
                   e.ExtendedProperties.Private != null &&
                   e.ExtendedProperties.Private.ContainsKey(GoogleCalendar.Instance.EventPropertyKey))
               {
                  logboxout("Clearing Google property: " + e.Summary);

                  e.ExtendedProperties.Private.Remove(GoogleCalendar.Instance.EventPropertyKey);

                  GoogleCalendar.Instance.updateEntry(e);
               }
            }

            // close the outlook calendar instance so not to always be logged in
            OutlookCalendar.Instance.Release();

            // indicate finish
            logboxout("");
            logboxout("Clear properties ended at " + DateTime.Now);
         }

         // restore button presses...
         bSyncNow.Enabled = true;
         clearUserPropertiesBtn.Enabled = true;
      }

      private void NumericOnlyKeyPress(object sender, KeyPressEventArgs e)
      {
         string allowed_chars = "0123456789\b";

         if (allowed_chars.IndexOf(e.KeyChar) == -1)
         {
            e.Handled = true;
         }
      }
   }
}

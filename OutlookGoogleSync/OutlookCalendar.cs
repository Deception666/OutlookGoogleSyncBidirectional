﻿using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3.Data;


namespace OutlookGoogleSync
{
   /// <summary>
   /// Description of OutlookCalendar.
   /// </summary>
   public class OutlookCalendar
   {
      private static OutlookCalendar instance;

      public static OutlookCalendar Instance
      {
         get
         {
            if (instance == null) instance = new OutlookCalendar();
            return instance;
         }
      }

      public static bool IsLoggedIn( )
      {
         return instance != null;
      }

      private Application OutlookApplication;
      private NameSpace OutlookNamespace;
      private MAPIFolder OutlookFolder;

      public OutlookCalendar()
      {
         // set the default google reminder time
         GoogleDefaultReminderMinutesBeforeStart = 15;

         // Create the Outlook application.
         OutlookApplication = new Application();

         // Get the NameSpace and Logon information.
         // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
         OutlookNamespace = OutlookApplication.GetNamespace("mapi");

         if (!Settings.Instance.OutlookAutoLogonEnabled)
         {
            //Log on by using a dialog box to choose the profile.
            OutlookNamespace.Logon("", "", true, true);
         }
         else
         {
            // Log on by using the profile name given...
            OutlookNamespace.Logon(Settings.Instance.OutlookAutoLogonProfileName, Settings.Instance.GetOutlookAutoLogonProfilePassword(), false, true);
         }

         //Alternate logon method that uses a specific profile.
         // If you use this logon method, 
         // change the profile name to an appropriate value.
         //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

         // Get the Calendar folder.
         OutlookFolder = OutlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);


         //Show the item to pause.
         //oAppt.Display(true);

         // Done. Log off.
         OutlookNamespace.Logoff();
      }

      public void Release( )
      {
         // quit the application
         ((_Application)OutlookApplication).Quit();

         // release the instance
         OutlookApplication = null;
         OutlookNamespace = null;
         OutlookFolder = null;
         instance = null;

         // http://msdn.microsoft.com/en-us/library/aa679807%28office.11%29.aspx#officeinteroperabilitych2_part2_gc
         GC.Collect();
         GC.WaitForPendingFinalizers();
         GC.Collect();
         GC.WaitForPendingFinalizers();

         // another good reference:
         // http://blogs.msdn.com/b/mstehle/archive/2007/12/07/oom-net-part-2-outlook-item-leaks.aspx
      }

      public List<AppointmentItem> getCalendarEntries()
      {
         Items OutlookItems = OutlookFolder.Items;
         if (OutlookItems != null)
         {
            List<AppointmentItem> result = new List<AppointmentItem>();
            foreach (AppointmentItem ai in OutlookItems)
            {
               result.Add(ai);
            }
            return result;
         }
         return null;
      }

      public List<AppointmentItem> getCalendarEntriesInRange()
      {
         List<AppointmentItem> result = new List<AppointmentItem>();

         Items OutlookItems = OutlookFolder.Items;
         OutlookItems.Sort("[Start]", Type.Missing);
         OutlookItems.IncludeRecurrences = true;

         if (OutlookItems != null)
         {
            DateTime min = DateTime.Now.AddDays(-Settings.Instance.DaysInThePast);
            DateTime max = DateTime.Now.AddDays(+Settings.Instance.DaysInTheFuture + 1);

            //initial version: did not work in all non-German environments
            //string filter = "[End] >= '" + min.ToString("dd.MM.yyyy HH:mm") + "' AND [Start] < '" + max.ToString("dd.MM.yyyy HH:mm") + "'";

            //proposed by WolverineFan, included here for future reference
            //string filter = "[End] >= '" + min.ToString("dd.MM.yyyy HH:mm") + "' AND [Start] < '" + max.ToString("dd.MM.yyyy HH:mm") + "'";

            //trying this instead, also proposed by WolverineFan, thanks!!! 
            string filter = "[End] >= '" + min.ToString("g") + "' AND [Start] < '" + max.ToString("g") + "'";


            foreach (AppointmentItem ai in OutlookItems.Restrict(filter))
            {
               result.Add(ai);
            }
         }
         return result;
      }

      public bool deleteCalendarEntry( AppointmentItem ai )
      {
         ai.Delete();

         return true;
      }

      public AppointmentItem addEntry( Event e, bool add_description, bool add_reminders, bool add_attendees )
      {
         AppointmentItem result = null;

         try
         {
            result = (AppointmentItem)OutlookApplication.CreateItem(OlItemType.olAppointmentItem);

            ModifyEvent(result, e, add_description, add_reminders, add_attendees);

            result.Save();
         }
         catch (System.Exception ex)
         {
            throw ex;
         }

         return result;
      }

      public AppointmentItem updateEntry( AppointmentItem ai, Event e, bool add_description, bool add_reminders, bool add_attendees )
      {
         ModifyEvent(ai, e, add_description, add_reminders, add_attendees);

         try
         {
            // save the outlook event
            ai.Save();
         }
         catch (System.Exception)
         {
            // unable to save the outlook item...
            // should report an error, but this will get handled by the main form...
         }

         return ai;
      }

      private void ModifyEvent( AppointmentItem ai, Event e, bool add_description, bool add_reminders, bool add_attendees )
      {
         ai.Start = new DateTime();
         ai.End = new DateTime();

         if (e.Start.Date != null && e.End.Date != null)
         {
            ai.AllDayEvent = true;
            ai.Start = DateTime.Parse(e.Start.Date);
            ai.End = DateTime.Parse(e.End.Date);
         }
         else
         {
            ai.AllDayEvent = false;
            ai.Start = e.Start.DateTime.Value;
            ai.End = e.End.DateTime.Value;
         }

         ai.Subject = e.Summary;
         if (add_description) ai.Body = OutlookGoogleSync.Utilities.ObtainUserBodyData(e.Description);
         ai.Location = e.Location;

         // determine how to set the status based on the google event
         OutlookGoogleSync.Utilities.SetEventStatus(ai, e);

         // consider the reminder set in google
         if (add_reminders)
         {
            if (e.Reminders != null)
            {
               if (e.Reminders.Overrides != null)
               {
                  if (e.Reminders.Overrides.Count > 0)
                  {
                     ai.ReminderMinutesBeforeStart = e.Reminders.Overrides[0].Minutes.Value;
                  }
               }
               else if (e.Reminders.UseDefault != null && (bool)e.Reminders.UseDefault)
               {
                  ai.ReminderMinutesBeforeStart = GoogleDefaultReminderMinutesBeforeStart;
               }
            }
         }

         if (add_attendees)
         {
            ai.Body += Environment.NewLine + Environment.NewLine;
            ai.Body += OutlookGoogleSync.Utilities.BODY_SEPARATOR;
            ai.Body += Environment.NewLine;
            ai.Body += Environment.NewLine + "==============================================";
            ai.Body += Environment.NewLine + "Added by OutlookGoogleSync Bidirectional:";
            ai.Body += Environment.NewLine + "ORGANIZER: " + Environment.NewLine + e.Organizer.DisplayName;
            ai.Body += Environment.NewLine + "REQUIRED: " + Environment.NewLine + splitAttendees(e.Attendees, true);
            ai.Body += Environment.NewLine + "OPTIONAL: " + Environment.NewLine + splitAttendees(e.Attendees, false);
            ai.Body += Environment.NewLine + "==============================================";
         }

         // bind the two events together...
         Bind(ai, e);
      }

      public void Bind( AppointmentItem ai, Event e )
      {
         OutlookGoogleSync.Utilities.BindEvents(ai, e, EventPropertyKey);
      }

      // one attendee per line
      public string splitAttendees( IList< EventAttendee > attendees, bool process_required )
      {
         string attendees_per_line = "";

         if (attendees != null)
         {
            foreach (var attendee in attendees)
            {
               if (process_required && !attendee.Optional.Value || !process_required && attendee.Optional.Value)
               {
                  attendees_per_line += attendee.DisplayName.Trim() + Environment.NewLine;
               }
            }
         }

         return attendees_per_line;
      }

      // defines the property name for associating outlook events to google events
      public string EventPropertyKey { get; set; }

      // formats the event proplery id value
      public static string FormatEventID( AppointmentItem ai )
      {
         string id = ai.GlobalAppointmentID;

         if (ai.IsRecurring)
         {
            id += " - RECURRENCE - " + ai.Start.ToString();
         }

         return id;
      }

      // defines the default google reminder trigger time in minutes
      public int GoogleDefaultReminderMinutesBeforeStart { get; set; }

   }
}

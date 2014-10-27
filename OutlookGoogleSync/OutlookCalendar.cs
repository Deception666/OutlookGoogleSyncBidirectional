using System;
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
            OutlookNamespace.Logon(Settings.Instance.OutlookAutoLogonProfileName, Settings.Instance.OutlookAutoLogonProfilePassword, false, true);
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


      public void Reset()
      {
         Release();

         instance = new OutlookCalendar();
      }

      public void Release( )
      {
         // quit the application
         ((_Application)OutlookApplication).Quit();
         // release the instance
         System.Runtime.InteropServices.Marshal.FinalReleaseComObject(OutlookFolder);
         System.Runtime.InteropServices.Marshal.FinalReleaseComObject(OutlookNamespace);
         System.Runtime.InteropServices.Marshal.FinalReleaseComObject(OutlookApplication.Session);
         System.Runtime.InteropServices.Marshal.FinalReleaseComObject(OutlookApplication);

         OutlookApplication = null;
         OutlookNamespace = null;
         OutlookFolder = null;
         instance = null;

         // http://msdn.microsoft.com/en-us/library/aa679807%28office.11%29.aspx#officeinteroperabilitych2_part2_gc
         GC.Collect();
         GC.WaitForPendingFinalizers();
         GC.Collect();
         GC.WaitForPendingFinalizers(); 
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

         ai.Save();

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
         if (add_description) ai.Body = e.Description;
         ai.Location = e.Location;

         // consider the reminder set in google
         if (add_reminders && e.Reminders != null && e.Reminders.Overrides != null)
         {
            if (e.Reminders.Overrides.Count > 0)
            {
               ai.ReminderMinutesBeforeStart = e.Reminders.Overrides[0].Minutes.Value;
            }
         }

         if (add_attendees)
         {
            ai.Body += Environment.NewLine;
            ai.Body += Environment.NewLine + "==============================================";
            ai.Body += Environment.NewLine + "Added by OutlookGoogleSync:";
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
         // TODO: move to a utility function, as it matches that of google

         // make sure to tag the user property of the google id
         UserProperty oitem_google_prop = ai.UserProperties.Find(EventPropertyKey);

         if (oitem_google_prop != null)
         {
            oitem_google_prop.Value = e.Id;
         }
         else
         {
            ai.UserProperties.Add(EventPropertyKey, OlUserPropertyType.olText).Value = e.Id;
         }

         // make sure to tag the private property of the outlook id
         if (e.ExtendedProperties == null)
         {
            e.ExtendedProperties = new Event.ExtendedPropertiesData();
            e.ExtendedProperties.Private = new Dictionary< string, string >();
         }

         e.ExtendedProperties.Private[EventPropertyKey] = OutlookCalendar.FormatEventID(ai);
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

   }
}

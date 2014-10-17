using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using DotNetOpenAuth.OAuth2;
using Google.Apis.Authentication;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util;
using Microsoft.Office.Interop.Outlook;


namespace OutlookGoogleSync
{
   /// <summary>
   /// Description of GoogleCalendar.
   /// </summary>
   public class GoogleCalendar
   {

      private static GoogleCalendar instance;

      public static GoogleCalendar Instance
      {
         get
         {
            if (instance == null) instance = new GoogleCalendar();
            return instance;
         }
      }

      private CalendarService service;

      public GoogleCalendar() { }

      public bool InitCalendarService( string user )
      {
         try
         {
            // obtain the user's credentials
            Google.Apis.Auth.OAuth2.UserCredential user_credential =
               Google.Apis.Auth.OAuth2.GoogleWebAuthorizationBroker.AuthorizeAsync(
                  new Google.Apis.Auth.OAuth2.ClientSecrets
                  {
                     ClientId = "857287834886-tau9nuakoamjctkm3plbqeojjrlmn140.apps.googleusercontent.com",
                     ClientSecret = "Gb5q1srU74At1erFyLQBKxXm"
                  },
                  new [] { Google.Apis.Calendar.v3.CalendarService.Scope.Calendar },
                  user,
                  System.Threading.CancellationToken.None).Result;

            if (user_credential != null)
            {
               // init the calendar service
               service =
                  new Google.Apis.Calendar.v3.CalendarService(
                     new Google.Apis.Services.BaseClientService.Initializer()
                     {
                        HttpClientInitializer = user_credential,
                        ApplicationName = "GoogleOutlookCalendarSync"
                     });

               // save the refresh token
               // this is not needed anymore, as the optional data store is a file store located
               // at Google.Apis.Auth under the user's Environment.SpecialFolder.ApplicationData
               // keep this for backwards compatibility, as it may be needed in the future
               Settings.Instance.RefreshToken = user_credential.Token.RefreshToken;

               // export the settings to keep backwards compatibility
               XMLManager.export(Settings.Instance, MainForm.FILENAME);
            }
         }
         catch (System.Exception)
         {
            // something went very wrong so the service is dead
            service = null;
         }

         // indicate if the service was initialized
         return service != null ? true : false;
      }

      public List<MyCalendarListEntry> getCalendars()
      {
         CalendarList request = null;
         try
         {
            request = service.CalendarList.List().Execute();
         }
         catch (System.Exception ex)
         {
            //MainForm.Instance.HandleException(ex);
            throw ex;
         }

         if (request != null)
         {

            List<MyCalendarListEntry> result = new List<MyCalendarListEntry>();
            foreach (CalendarListEntry cle in request.Items)
            {
               result.Add(new MyCalendarListEntry(((Google.Apis.Auth.OAuth2.UserCredential)service.HttpClientInitializer).UderId, cle));
            }
            return result;
         }
         return null;
      }



      public List<Event> getCalendarEntriesInRange()
      {
         List<Event> result = new List<Event>();
         Events request = null;

         try
         {
            EventsResource.ListRequest lr = service.Events.List(Settings.Instance.UseGoogleCalendar.Id);

            lr.TimeMin = DateTime.Now.AddDays(-Settings.Instance.DaysInThePast);
            lr.TimeMax = DateTime.Now.AddDays(+Settings.Instance.DaysInTheFuture + 1);

            do
            {
               // request the current page of information
               request = lr.Execute();

               // add to the results the current items
               result.AddRange(request.Items);

               // request the next page
               lr.PageToken = request.NextPageToken;
            } while (lr.PageToken != null);
         }
         catch (System.Exception ex)
         {
            //MainForm.Instance.HandleException(ex);
            throw ex;
         }

         return result;
      }

      public void deleteCalendarEntry(Event e)
      {
         string request;

         try
         {
            request = service.Events.Delete(Settings.Instance.UseGoogleCalendar.Id, e.Id).Execute();
         }
         catch (System.Exception ex)
         {
            //MainForm.Instance.HandleException(ex);
            throw ex;
         }
      }

      public Event addEntry(Event e)
      {
         Event result = null;

         try
         {
            result = service.Events.Insert(e, Settings.Instance.UseGoogleCalendar.Id).Execute();
         }
         catch (System.Exception ex)
         {
            //MainForm.Instance.HandleException(ex);
            throw ex;
         }

         return result;
      }

      public Event updateEntry( Event e )
      {
         Event result = null;

         try
         {
            result = service.Events.Update(e, Settings.Instance.UseGoogleCalendar.Id, e.Id).Execute();
         }
         catch (System.Exception ex)
         {
            throw ex;
         }

         return result;
      }

      public Event updateEntry( Event e, AppointmentItem ai, bool add_description, bool add_reminders, bool add_attendees )
      {
         Event result = null;

         try
         {
            ModifyEvent(e, ai, add_description, add_reminders, add_attendees);

            result = updateEntry(e);
         }
         catch (System.Exception ex)
         {
            throw ex;
         }

         return result;
      }

      public Event addEntry( AppointmentItem ai, bool add_description, bool add_reminders, bool add_attendees )
      {
         Event result = null;

         try
         {
            Event e = new Event();

            ModifyEvent(e, ai, add_description, add_reminders, add_attendees);

            result = addEntry(e);

            // since this is a new event, we need to update the ids again so that outlook has it
            UpdatePropertyIDs(result, ai);
         }
         catch (System.Exception ex)
         {
            throw ex;
         }

         return result;
      }

      private void ModifyEvent( Event e, AppointmentItem ai, bool add_description, bool add_reminders, bool add_attendees )
      {
         e.Start = new EventDateTime();
         e.End = new EventDateTime();

         if (ai.AllDayEvent)
         {
            e.Start.Date = ai.Start.ToString("yyyy-MM-dd");
            e.End.Date = ai.End.ToString("yyyy-MM-dd");

            // nullify the DateTime; otherwise, google rejects
            e.Start.DateTime = null;
            e.Start.DateTime = null;
         }
         else
         {
            e.Start.DateTime = ai.Start;
            e.End.DateTime = ai.End;
         }

         e.Summary = ai.Subject;
         if (add_description) e.Description = ai.Body;
         e.Location = ai.Location;

         // consider the reminder set in Outlook
         if (add_reminders && ai.ReminderSet)
         {
            e.Reminders = new Event.RemindersData();
            e.Reminders.UseDefault = false;
            EventReminder reminder = new EventReminder();
            reminder.Method = "popup";
            reminder.Minutes = ai.ReminderMinutesBeforeStart;
            e.Reminders.Overrides = new List<EventReminder>();
            e.Reminders.Overrides.Add(reminder);
         }

         if (add_attendees)
         {
            e.Description += Environment.NewLine;
            e.Description += Environment.NewLine + "==============================================";
            e.Description += Environment.NewLine + "Added by OutlookGoogleSync:" + Environment.NewLine;
            e.Description += Environment.NewLine + "ORGANIZER: " + Environment.NewLine + ai.Organizer + Environment.NewLine;
            e.Description += Environment.NewLine + "REQUIRED: " + Environment.NewLine + splitAttendees(ai.RequiredAttendees) + Environment.NewLine;
            e.Description += Environment.NewLine + "OPTIONAL: " + Environment.NewLine + splitAttendees(ai.OptionalAttendees);
            e.Description += Environment.NewLine + "==============================================";
         }

         // update the property ids that tie the two events together
         UpdatePropertyIDs(e, ai);
      }

      private void UpdatePropertyIDs( Event e, AppointmentItem ai )
      {
         // make sure to tag the private property of the outlook id
         if (e.ExtendedProperties == null)
         {
            e.ExtendedProperties = new Event.ExtendedPropertiesData();
            e.ExtendedProperties.Private = new Dictionary< string, string >();
         }

         e.ExtendedProperties.Private[EventPropertyKey] = ai.GlobalAppointmentID;

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

         // save the outlook event
         ai.Save();
      }

      // returns the Google Time Format String of a given .Net DateTime value
      // Google Time Format = "2012-08-20T00:00:00+02:00"
      public string GoogleTimeFrom(DateTime dt)
      {
         string timezone = TimeZoneInfo.Local.GetUtcOffset(dt).ToString();
         if (timezone[0] != '-') timezone = '+' + timezone;
         timezone = timezone.Substring(0, 6);

         string result = dt.GetDateTimeFormats('s')[0] + timezone;
         return result;
      }

      // one attendee per line
      public string splitAttendees(string attendees)
      {
         if (attendees == null) return "";
         string[] tmp1 = attendees.Split(';');
         for (int i = 0; i < tmp1.Length; i++) tmp1[i] = tmp1[i].Trim();
         return String.Join(Environment.NewLine, tmp1);
      }

      // defines the property name for associating google events to outlook events
      public string EventPropertyKey { get; set; }
   }

}

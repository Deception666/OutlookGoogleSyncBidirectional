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
         catch (Exception ex)
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

            // request = lr.Fetch();
         }
         catch (Exception ex)
         {
            //MainForm.Instance.HandleException(ex);
            throw ex;
         }

         if (request != null)
         {
            if (request.Items != null) result.AddRange(request.Items);
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
         catch (Exception ex)
         {
            //MainForm.Instance.HandleException(ex);
            throw ex;
         }
      }

      public void addEntry(Event e)
      {
         try
         {
            var result = service.Events.Insert(e, Settings.Instance.UseGoogleCalendar.Id).Execute();
         }
         catch (Exception ex)
         {
            //MainForm.Instance.HandleException(ex);
            throw ex;
         }
      }


      //returns the Google Time Format String of a given .Net DateTime value
      //Google Time Format = "2012-08-20T00:00:00+02:00"
      public string GoogleTimeFrom(DateTime dt)
      {
         string timezone = TimeZoneInfo.Local.GetUtcOffset(dt).ToString();
         if (timezone[0] != '-') timezone = '+' + timezone;
         timezone = timezone.Substring(0, 6);

         string result = dt.GetDateTimeFormats('s')[0] + timezone;
         return result;
      }


   }
}

using System;
using System.Collections.Generic;
using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleSync
{

class Utilities
{
   // defines the body separator, separating the user information from what this application adds
   public static readonly string BODY_SEPARATOR =
      "- - - - - - - - - - - - - - - - - - - - Outlook Google Sync Bidirectional - - - - - - - - - - - - - - - - - - - -";

   // obtains the text from the body / description up until the body separator
   public static string ObtainUserBodyData( string body )
   {
      if (body != null)
      {
         int body_sep_loc = body.IndexOf(BODY_SEPARATOR);

         return body_sep_loc != -1 ? body.Substring(0, body_sep_loc).TrimEnd((" " + Environment.NewLine).ToCharArray()) : body;
      }

      return "";
   }

   // binds the two events together
   public static void BindEvents( AppointmentItem oitem, Event gitem, string event_property_key )
   {
      // make sure to tag the user property of the google id
      UserProperty oitem_google_prop = oitem.UserProperties.Find(event_property_key);

      if (oitem_google_prop != null)
      {
         oitem_google_prop.Value = gitem.Id;

         System.Runtime.InteropServices.Marshal.ReleaseComObject(oitem_google_prop);
      }
      else
      {
         oitem.UserProperties.Add(event_property_key, OlUserPropertyType.olText).Value = gitem.Id;
      }

      try
      {
         // save the outlook event
         oitem.Save();
      }
      catch (System.Exception)
      {
         // unable to save the outlook item...
         // should report an error, but this will get handled by the main form...
      }

      // make sure to tag the private property of the outlook id
      if (gitem.ExtendedProperties == null)
      {
         gitem.ExtendedProperties = new Event.ExtendedPropertiesData();
         gitem.ExtendedProperties.Shared = new Dictionary< string, string >();
      }
      else if (gitem.ExtendedProperties.Shared == null)
      {
         gitem.ExtendedProperties.Shared = new Dictionary< string, string >();
      }

      gitem.ExtendedProperties.Shared[event_property_key] = OutlookCalendar.FormatEventID(oitem);
   }

   // event status constants
   public static readonly string EVENT_STATUS_FREE = "Free";
   public static readonly string EVENT_STATUS_BUSY = "Busy";
   public static readonly string EVENT_STATUS_TENTATIVE = "Tentative";

   // determines the status of the event (busy / free / tentative)
   public static string EventStatus( AppointmentItem ai )
   {
      string status = EVENT_STATUS_BUSY;

      switch (ai.BusyStatus)
      {
         case OlBusyStatus.olFree: status = EVENT_STATUS_FREE; break;
         case OlBusyStatus.olTentative: status = EVENT_STATUS_TENTATIVE; break;

         case OlBusyStatus.olBusy:
         case OlBusyStatus.olOutOfOffice:
         case OlBusyStatus.olWorkingElsewhere:
         default:

            status = EVENT_STATUS_BUSY;

            break;
      }

      return status;
   }

   // sets the correct event status for the outlook event based on the google event
   public static void SetEventStatus( AppointmentItem ai, Event ev )
   {
      string google_status = EventStatus(ev);

      if (google_status == EVENT_STATUS_FREE)
      {
         ai.BusyStatus = OlBusyStatus.olFree;
      }
      else if (google_status == EVENT_STATUS_TENTATIVE)
      {
         ai.BusyStatus = OlBusyStatus.olTentative;
      }
      else
      {
         ai.BusyStatus = OlBusyStatus.olBusy;
      }
   }

   // determines the status of the event (busy / free / tentative)
   public static string EventStatus( Event ev )
   {
      string status = EVENT_STATUS_BUSY;

      if (ev.Status != null)
      {
         if (ev.Status == "confirmed" && (ev.Transparency == null || ev.Transparency == "opaque"))
         {
            status = EVENT_STATUS_BUSY;
         }
         else if (ev.Status == "confirmed" && ev.Transparency != null && ev.Transparency == "transparent")
         {
            status = EVENT_STATUS_FREE;
         }
         else if (ev.Status == "tentative")
         {
            status = EVENT_STATUS_TENTATIVE;
         }
      }

      return status;
   }

   // sets the correct event status for the google event based on the outlook event
   public static void SetEventStatus( Event ev, AppointmentItem ai )
   {
      string outlook_status = EventStatus(ai);

      if (outlook_status == EVENT_STATUS_FREE)
      {
         ev.Status = "confirmed";
         ev.Transparency = "transparent";
      }
      else if (outlook_status == EVENT_STATUS_TENTATIVE)
      {
         ev.Status = "tentative";
         ev.Transparency = "transparent";
      }
      else
      {
         ev.Status = "confirmed";
         ev.Transparency = "opaque";
      }
   }
};

} // namespace OutlookGoogleSync

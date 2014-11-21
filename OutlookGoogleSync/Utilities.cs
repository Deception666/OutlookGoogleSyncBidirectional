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

      // save the outlook event
      oitem.Save();

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
};

} // namespace OutlookGoogleSync

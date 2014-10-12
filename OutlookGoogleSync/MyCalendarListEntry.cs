using System;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;


namespace OutlookGoogleSync
{
   /// <summary>
   /// Description of MyCalendarListEntry.
   /// </summary>
   public class MyCalendarListEntry
   {
      public string Id = "";
      public string Name = "";
      public string User = "";

      public MyCalendarListEntry()
      {
      }

      public MyCalendarListEntry(string user, CalendarListEntry init)
      {
         Id = init.Id;
         Name = init.Summary;
         User = user;
      }

      public override string ToString()
      {
         return Name;
      }


   }
}

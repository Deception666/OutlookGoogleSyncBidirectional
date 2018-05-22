
using System;
using System.Drawing;

namespace OutlookGoogleSync
{
   /// <summary>
   /// Description of Settings.
   /// </summary>
   public class Settings
   {
      private static Settings instance;

      public static Settings Instance
      {
         get
         {
            if (instance == null) instance = new Settings();
            return instance;
         }
         set
         {
            instance = value;
         }

      }


      public string RefreshToken = "";
      public string MinuteOffsets = "";
      public DateTime LastSyncDate = new DateTime(0);
      public int DaysInThePast = 1;
      public int DaysInTheFuture = 60;
      public MyCalendarListEntry UseGoogleCalendar = new MyCalendarListEntry();

      public bool SyncEveryHour = false;
      public bool ShowBubbleTooltipWhenSyncing = false;
      public bool StartInTray = false;
      public bool MinimizeToTray = false;

      public bool AddDescription = true;
      public bool AddReminders = false;
      public bool AddAttendeesToDescription = true;
      public bool CreateTextFiles = true;

      public string OutlookAutoLogonProfileName = "";
      public byte[] OutlookAutoLogonProfilePassword = null;
      public bool OutlookAutoLogonEnabled = false;
      public bool OutlookKeepOpenAfterSync = false;

      public string GetOutlookAutoLogonProfilePassword( )
      {
         string password = "";

         if (OutlookAutoLogonProfilePassword != null && OutlookAutoLogonProfilePassword.Length != 0)
         {
            password = System.Text.Encoding.ASCII.GetString(
               System.Security.Cryptography.ProtectedData.Unprotect(OutlookAutoLogonProfilePassword,
                                                                    null,
                                                                    System.Security.Cryptography.DataProtectionScope.CurrentUser));
         }

         return password;
      }

      public void SetOutlookAutoLogonProfilePassword( string password )
      {
         OutlookAutoLogonProfilePassword =
            System.Security.Cryptography.ProtectedData.Protect(System.Text.Encoding.ASCII.GetBytes(password),
                                                               null,
                                                               System.Security.Cryptography.DataProtectionScope.CurrentUser);
      }

   }
}

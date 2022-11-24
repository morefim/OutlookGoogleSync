using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using MAPIFolder = Microsoft.Office.Interop.Outlook.MAPIFolder;

//using Redemption;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of OutlookCalendar.
    /// </summary>
    public class OutlookCalendar
    {
	    private static OutlookCalendar _instance;

        public static OutlookCalendar Instance
        {
            get 
            {
                try
                {
                    if (_instance == null || !Marshal.IsComObject(_instance.UseOutlookCalendar))
                        _instance = null;

                    if (_instance == null)
                        _instance = new OutlookCalendar();

                    return _instance;
                }
                finally
                {
                    _instance = null;
                }
            }
        }
        
        public MAPIFolder UseOutlookCalendar;

        public OutlookCalendar()
        {
            // Create the Outlook application.
            var oApp = new Application();

            // Get the NameSpace and Logon information.
            // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
            var oNs = oApp.GetNamespace("mapi");

            //Log on by using a dialog box to choose the profile.
            oNs.Logon(OgsSettings.Instance.User, OgsSettings.Instance.OutlookPassword, true, true);

            //Alternate logon method that uses a specific profile.
            // If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 
			
            // Get the Calendar folder.
            UseOutlookCalendar = oNs.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            
            //Show the item to pause.
            //oAppt.Display(true);

            // Done. Log off.
            oNs.Logoff();
        }

        public List<AppointmentItem> GetCalendarEntries()
        {
            var outlookItems = UseOutlookCalendar.Items;
            if (outlookItems != null)
            {
                var result = new List<AppointmentItem>();
                foreach (AppointmentItem ai in outlookItems)
                {
                    result.Add(ai);
                }
                return result;
            }
            return null;
        }
        
        public List<AppointmentItem> GetCalendarEntriesInRange()
        {
            var result = new List<AppointmentItem>();
            
            var outlookItems = UseOutlookCalendar.Items;
            outlookItems.Sort("[Start]",Type.Missing);
            outlookItems.IncludeRecurrences = true;
            
            if (outlookItems != null)
            {
                var min = DateTime.Now.AddDays(-OgsSettings.Instance.DaysInThePast);
                var max = DateTime.Now.AddDays(+OgsSettings.Instance.DaysInTheFuture+1);

                //initial version: did not work in all non-German environments
                //string filter = "[End] >= '" + min.ToString("dd.MM.yyyy HH:mm") + "' AND [Start] < '" + max.ToString("dd.MM.yyyy HH:mm") + "'";
                
                //proposed by WolverineFan, included here for future reference
                //string filter = "[End] >= '" + min.ToString("dd.MM.yyyy HH:mm") + "' AND [Start] < '" + max.ToString("dd.MM.yyyy HH:mm") + "'";

                //trying this instead, also proposed by WolverineFan, thanks!!! 
                var filter = "[End] >= '" + min.ToString("g") + "' AND [Start] < '" + max.ToString("g") + "'";
                
                
                foreach(AppointmentItem ai in outlookItems.Restrict(filter))
                {
                    result.Add(ai);
                }
            }
            return result;
        }
        
        
        
    }
}

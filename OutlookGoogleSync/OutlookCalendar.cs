using System;
using System.Collections.Generic;
//using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
//using Redemption;

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
                try
                {
                    if (instance == null || !Marshal.IsComObject(instance.UseOutlookCalendar))
                        instance = null;

                    if (instance == null)
                        instance = new OutlookCalendar();

                    return instance;
                }
                finally
                {
                    instance = null;
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
            var oNS = oApp.GetNamespace("mapi");

            //Log on by using a dialog box to choose the profile.
            oNS.Logon(OGSSettings.Instance.User, OGSSettings.Instance.OutlookPassword, true, true);

            //Alternate logon method that uses a specific profile.
            // If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 
			
            // Get the Calendar folder.
            UseOutlookCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

            
            //Show the item to pause.
            //oAppt.Display(true);

            // Done. Log off.
            oNS.Logoff();
        }
        
        
        public List<AppointmentItem> getCalendarEntries()
        {
            var OutlookItems = UseOutlookCalendar.Items;
            if (OutlookItems != null)
            {
                var result = new List<AppointmentItem>();
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
            var result = new List<AppointmentItem>();
            
            var OutlookItems = UseOutlookCalendar.Items;
            OutlookItems.Sort("[Start]",Type.Missing);
            OutlookItems.IncludeRecurrences = true;
            
            if (OutlookItems != null)
            {
                var min = DateTime.Now.AddDays(-OGSSettings.Instance.DaysInThePast);
                var max = DateTime.Now.AddDays(+OGSSettings.Instance.DaysInTheFuture+1);

                //initial version: did not work in all non-German environments
                //string filter = "[End] >= '" + min.ToString("dd.MM.yyyy HH:mm") + "' AND [Start] < '" + max.ToString("dd.MM.yyyy HH:mm") + "'";
                
                //proposed by WolverineFan, included here for future reference
                //string filter = "[End] >= '" + min.ToString("dd.MM.yyyy HH:mm") + "' AND [Start] < '" + max.ToString("dd.MM.yyyy HH:mm") + "'";

                //trying this instead, also proposed by WolverineFan, thanks!!! 
                var filter = "[End] >= '" + min.ToString("g") + "' AND [Start] < '" + max.ToString("g") + "'";
                
                
                foreach(AppointmentItem ai in OutlookItems.Restrict(filter))
                {
                    result.Add(ai);
                }
            }
            return result;
        }
        
        
        
    }
}

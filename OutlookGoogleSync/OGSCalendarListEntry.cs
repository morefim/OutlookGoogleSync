using System;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of MyCalendarListEntry.
    /// </summary>
    public class OGSCalendarListEntry
    {
        public string ID = "";

        public string Name = "";

        public bool IsEmpty { get { return ID == ""; } }

        public string CalendarID { get { return Coder.Decrypt(ID); } }

        public OGSCalendarListEntry()
        {
        }
        
        public OGSCalendarListEntry(CalendarListEntry init)
        {
            ID = Coder.Encrypt(init.Id);
            Name = Coder.Encrypt(init.Summary);
        }
        
        public override string ToString()
		{
            return Coder.Decrypt(Name);
		}        
    }
}

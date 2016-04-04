using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of MyCalendarListEntry.
    /// </summary>
    public class OgsCalendarListEntry
    {
        public string Id = "";

        public string Name = "";

        public bool IsEmpty { get { return Id == ""; } }

        public string CalendarId { get { return Coder.Decrypt(Id); } }

        public OgsCalendarListEntry()
        {
        }
        
        public OgsCalendarListEntry(CalendarListEntry init)
        {
            Id = Coder.Encrypt(init.Id);
            Name = Coder.Encrypt(init.Summary);
        }
        
        public override string ToString()
		{
            return Coder.Decrypt(Name);
		}        
    }
}

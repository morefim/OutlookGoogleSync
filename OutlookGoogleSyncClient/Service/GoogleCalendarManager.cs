using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookGoogleSyncClient.Service
{
    public interface IGoogleCalendarManager
    {
        IEnumerable<GoogleCalendar> GetCalendars();
    }

    public class GoogleCalendarManager : IGoogleCalendarManager
    {
        public IEnumerable<GoogleCalendar> GetCalendars()
        {
            return new[] { new GoogleCalendar { Name =  "First" }, new GoogleCalendar { Name = "Second" } };
        }
    }
}

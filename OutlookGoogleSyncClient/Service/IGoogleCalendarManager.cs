using System.Collections.Generic;

namespace OutlookGoogleSyncClient.Service
{
    public interface IGoogleCalendarManager
    {
        IEnumerable<GoogleCalendar> GetCalendars();
    }
}
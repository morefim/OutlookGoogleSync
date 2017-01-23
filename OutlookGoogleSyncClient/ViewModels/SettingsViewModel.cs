using System.Linq;
using Caliburn.Micro;
using OutlookGoogleSyncClient.Service;

namespace OutlookGoogleSyncClient.ViewModels
{
    public class SettingsViewModel : Screen
    {
        public BindableCollection<GoogleCalendar> GoogleCalendars { get; set; }

        public GoogleCalendar SelectedGoogleCalendar { get; set; }

        public SettingsViewModel(IGoogleCalendarManager googleCalendarManager)
        {
            GoogleCalendars = new BindableCollection<GoogleCalendar>();

            var calendars = googleCalendarManager.GetCalendars();
            GoogleCalendars.AddRange(calendars);

            SelectedGoogleCalendar = GoogleCalendars.First();
        }
    }
}

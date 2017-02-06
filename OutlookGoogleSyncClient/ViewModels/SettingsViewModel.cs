using System.Linq;
using Caliburn.Micro;
using OutlookGoogleSyncClient.Service;

namespace OutlookGoogleSyncClient.ViewModels
{
    public class SettingsViewModel : Screen
    {
        public BindableCollection<GoogleCalendar> GoogleCalendars { get; set; }

        private readonly IOutlookGoogleSyncSettings _settings;
        
        public GoogleCalendar SelectedGoogleCalendar
        {
            get
            {
                return _settings.GoogleCalendar == null ? GoogleCalendars.First() : GoogleCalendars.FirstOrDefault(o => o.Id == _settings.GoogleCalendar.Id);
            }

            set
            {
                _settings.GoogleCalendar = new OutlookGoogleSyncSettings.GoogleCalendarEntry { Id = value.Id, Name = value.Name };
                _settings.Save();
                NotifyOfPropertyChange();
            }
        }

        public SettingsViewModel(IGoogleCalendarManager googleCalendarManager, IOutlookGoogleSyncSettings settings)
        {
            _settings = settings;
            GoogleCalendars = new BindableCollection<GoogleCalendar>();

            var calendars = googleCalendarManager.GetCalendars();
            GoogleCalendars.AddRange(calendars);

            NotifyOfPropertyChange(() => SelectedGoogleCalendar);
        }
    }
}

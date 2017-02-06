using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;


namespace OutlookGoogleSyncClient.Service
{
    /// <summary>
    /// Description of GoogleCalendar.
    /// </summary>
    public class GoogleCalendarManager : IGoogleCalendarManager
    {
        private const string ApplicationName = "Outlook Google Calendar Sync Engine";
        /// <summary>
        /// From Google Developer console https://console.developers.google.com
        /// </summary>
        private const string ClientId = "662204240419.apps.googleusercontent.com";
        /// <summary>
        /// From Google Developer console https://console.developers.google.com
        /// </summary>
        private const string ClientSecret = "4nJPnk5fE8yJM_HNUNQEEvjU";
        /// <summary>
        /// A string used to identify a user.
        /// </summary>
        private const string UserName = "user";

        private readonly CalendarService _service;

        private readonly IOutlookGoogleSyncSettings _settings;

        private static readonly string[] Scopes =
        {
            CalendarService.Scope.Calendar, // Manage your calendars
            CalendarService.Scope.CalendarReadonly // View your Calendars 
        };

        public GoogleCalendarManager(IOutlookGoogleSyncSettings settings)
        {
            _settings = settings;

            // here is where we Request the user to give us access, or use the Refresh Token that was previously stored in %AppData%
            var clientSecrets = new ClientSecrets
            {
                ClientId = ClientId,
                ClientSecret = ClientSecret
            };
            var fileDataStore = new FileDataStore("GoogleSyncEngine.Auth.Store");

            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, Scopes, UserName, CancellationToken.None, fileDataStore).ConfigureAwait(false).GetAwaiter().GetResult();

            // Create the service.
            _service = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        public IEnumerable<GoogleCalendar> GetCalendars()
        {
            var request = _service.CalendarList.List().Execute();
            return request?.Items.Select(cle => new GoogleCalendar { Name = cle.Summary, Id = cle.Id });
        }

        public List<Event> GetCalendarEntriesInRange()
        {
            var result = new List<Event>();
            var lr = _service.Events.List(_settings.GoogleCalendar.Id);

            lr.TimeMin = DateTime.Now.AddDays(-_settings.DaysInThePast);
            lr.TimeMax = DateTime.Now.AddDays(+_settings.DaysInTheFuture + 1);

            var request = lr.Execute();
            if (request?.Items != null)
            {
                result.AddRange(request.Items);
            }
            return result;
        }

        public void DeleteCalendarEntry(Event e)
        {
            _service.Events.Delete(_settings.GoogleCalendar.Id, e.Id).Execute();
        }

        public void AddEntry(Event e)
        {
            try
            {
                _service.Events.Insert(e, _settings.GoogleCalendar.Id).Execute();
            }
            catch (GoogleApiException ex)
            {
                const string error403Signature = @"Calendar usage limits exceeded.";
                if (ex.Error.Code == 403 && ex.Error.Message == error403Signature)
                    MoveAttendeesToDescriptionAndRetry(e);
                else
                    throw;
            }
        }

        private void MoveAttendeesToDescriptionAndRetry(Event e)
        {
            e.Description += "Attendees:\r\n";
            foreach (var attendee in e.Attendees)
                e.Description += attendee.DisplayName + "[" + attendee.Email + "]\r\n";

            e.Attendees.Clear();
            e.Attendees = null;

            _service.Events.Insert(e, _settings.GoogleCalendar.Id).Execute();
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reactive.Subjects;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Polly;
using Polly.Retry;


namespace OutlookGoogleSync
{
	/// <summary>
	/// Description of GoogleCalendar.
	/// </summary>
	public class GoogleCalendar
    {
        private const string ClientId = "1050291135187-5c9ilee5s2398e5glsejtqghetbi114d.apps.googleusercontent.com";
        private const string ClientSecret = "uQiR7G6ExPMxXxXgrZ9eDpoj";


        private const string ApplicationName = "Outlook Google Calendar Sync Engine";
        /// <summary>
        /// From Google Developer console https://console.developers.google.com
        /// </summary>
        //private const string ClientId = "662204240419.apps.googleusercontent.com";
        /// <summary>
        /// From Google Developer console https://console.developers.google.com
        /// </summary>
        //private const string ClientSecret = "4nJPnk5fE8yJM_HNUNQEEvjU";
        /// <summary>
        /// A string used to identify a user.
        /// </summary>
        private const string UserName = "user";

	    readonly CalendarService _service;

	    private static GoogleCalendar _instance;

        public static GoogleCalendar Instance
        {
            get { return _instance ?? (_instance = new GoogleCalendar()); }
        }

        private static readonly string[] Scopes = 
        { 
            CalendarService.Scope.Calendar, // Manage your calendars
            CalendarService.Scope.CalendarReadonly // View your Calendars 
        };

        public ISubject<Exception> OnException = new Subject<Exception>();

        public ISubject<string> OnLog = new Subject<string>();

        private readonly AsyncRetryPolicy _policy;

        public GoogleCalendar()
        {
            NEVER_EAT_POISON_Disable_CertificateValidation();

            // here is where we Request the user to give us access, or use the Refresh Token that was previously stored in %AppData%
            var clientSecrets = new ClientSecrets
            {
                ClientId = ClientId,
                ClientSecret = ClientSecret
            };
            var fileDataStore = new FileDataStore("GoogleSyncEngine.Auth.Store");

            var credential = AsyncHelpers.RunSync(() => GoogleWebAuthorizationBroker.AuthorizeAsync(clientSecrets, Scopes, UserName, CancellationToken.None, fileDataStore));

            // Create the service.
            _service = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            var waitTime = TimeSpan.FromSeconds(10);

            _policy = Policy
                .Handle<Exception>()
                .WaitAndRetryAsync(3, _ => waitTime,
                    (exception, timeSpan, context) =>
                    {
                        if (context.TryGetValue("Event", out var eventValue))
                        {
                            if (eventValue is Event e)
                            {
                                MoveAttendeesToDescription(e);
                            }
                        }

                        var ex = new Exception($"Failed to add Google Calendar Event. Retry in {timeSpan}", exception);
                        OnException?.OnNext(ex);
                    });
        }

        //[Obsolete("Do not use this in Production code!!!", true)]
        static void NEVER_EAT_POISON_Disable_CertificateValidation()
        {
            // Disabling certificate validation can expose you to a man-in-the-middle attack
            // which may allow your encrypted message to be read by an attacker
            // https://stackoverflow.com/a/14907718/740639
            ServicePointManager.ServerCertificateValidationCallback = (s, certificate, chain, sslPolicyErrors) => true;
        }

        public async Task<List<OgsCalendarListEntry>> GetCalendars()
        {
            CalendarList request = await _service.CalendarList.List().ExecuteAsync();          
            return request != null ? request.Items.Select(cle => new OgsCalendarListEntry(cle)).ToList() : null;
        }
		
        public async Task<List<Event>> GetCalendarEntriesInRange()
        {
            var result = new List<Event>();
            var lr = _service.Events.List(OgsSettings.Instance.UseGoogleCalendar.CalendarId);

            lr.TimeMin = DateTime.Now.AddDays(-OgsSettings.Instance.DaysInThePast);
            lr.TimeMax = DateTime.Now.AddDays(+OgsSettings.Instance.DaysInTheFuture + 1);

            var request = await lr.ExecuteAsync();
            if (request != null && request.Items != null)
            {
                result.AddRange(request.Items);
            }
            return result;
        }

        public async Task<string> DeleteCalendarEntry(Event e)
        {
            return await _service.Events.Delete(OgsSettings.Instance.UseGoogleCalendar.CalendarId, e.Id).ExecuteAsync();
        }

        public async Task<Event> AddEntry(Event e)
        {
            //try
            //{
                if (e.Attendees?.Count > 15)
                    MoveAttendeesToDescription(e);

                var context = new Dictionary<string, object> { {"Event", e} };

                var result = await _policy.ExecuteAndCaptureAsync(async (c) =>
                    await _service.Events.Insert(e, OgsSettings.Instance.UseGoogleCalendar.CalendarId).ExecuteAsync(), context);

                if (result.FinalException == null)
                    return result.Result;

                throw result.FinalException;
            //}
            /*catch (GoogleApiException ex)
            {
                const string error403Signature = @"Calendar usage limits exceeded.";
                if (ex.Error.Code == 403 && ex.Error.Message.Contains(error403Signature))
                {
                    MoveAttendeesToDescription(e);
                }

                throw;
            }
            catch (Exception ex)
            {
                OnException?.OnNext(ex);
                throw;
            }*/
        }

	    private void MoveAttendeesToDescription(Event e)
	    {
            e.Description += "Attendees:\r\n";
            foreach (var attendee in e.Attendees)
                e.Description += attendee.DisplayName + "[" + attendee.Email + "]\r\n";

            e.Attendees?.Clear();
	        e.Attendees = null;
	    }				
	}
}

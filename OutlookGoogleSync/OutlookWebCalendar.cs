using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Task = System.Threading.Tasks.Task;
using System.Configuration;
using OutlookGoogleSync.Properties;

namespace OutlookGoogleSync
{
    public class OutlookWebCalendar
    {
        private static OutlookWebCalendar _instance;

        public static OutlookWebCalendar Instance
        {
            get
            {
                try
                {
                    return _instance ?? (_instance = new OutlookWebCalendar());
                }
                finally
                {
                    _instance = null;
                }
            }
        }

        private async Task<ExchangeService> ConnectToService()
        {
            var sss = new MSgraphV2();
            var token = await sss.CallGraphButton_Click(1, null);

            throw new Exception(token);

            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1)
            {
                Credentials = new WebCredentials(OgsSettings.Instance.User, OgsSettings.Instance.OutlookPassword),
                Url = new Uri("https://outlook.office365.com/ews/exchange.asmx")
            };

            //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            //service.AutodiscoverUrl(OgsSettings.Instance.User, RedirectionUrlValidationCallback);
            //service.UseDefaultCredentials = true;

            //ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            //service.Credentials = new WebCredentials(OgsSettings.Instance.User, OgsSettings.Instance.OutlookPassword);
            //service.AutodiscoverUrl(OgsSettings.Instance.User, RedirectionUrlValidationCallback);

            return await Task.FromResult(service);
        }

        private async Task<ExchangeService> ConnectToService1()
        {
            // Using Microsoft.Identity.Client 4.22.0
            var cca = ConfidentialClientApplicationBuilder
                .Create(MainSettings.Default.appId)
                .WithClientSecret(MainSettings.Default.clientSecret)
                .WithTenantId(MainSettings.Default.tenantId)
                .Build();

            var ewsScopes = new[] { "https://outlook.office365.com/.default" };

            var authResult = await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync();

            // Configure the ExchangeService with the access token
            var ewsClient = new ExchangeService
            {
                Credentials = new OAuthCredentials(authResult.AccessToken),
                Url = new Uri("https://outlook.office365.com/ews/exchange.asmx"),
                //ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "morefim_hotmail.com#EXT#@morefimhotmail.onmicrosoft.com")
            };

            //Include x-anchormailbox header
            /*ewsClient.HttpHeaders.Add("X-AnchorMailbox", "morefim_hotmail.com#EXT#@morefimhotmail.onmicrosoft.com");

            // Make an EWS call
            var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));
            foreach (var folder in folders)
            {
                Console.WriteLine($@"Folder: {folder.DisplayName}");
            }*/

            return ewsClient;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        public async Task<List<WebAppointmentItem>> GetCalendarEntriesInRange()
        {
            var appointmentsList = new List<WebAppointmentItem>();
            ExchangeService service = await ConnectToService();

            const int maxItemsReturned = 50;

            // Initialize the calendar folder object with only the folder ID.
            CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar/*, new PropertySet()*/);

            // get the start and end times
            var startDate = DateTime.Now.AddDays(-OgsSettings.Instance.DaysInThePast);
            var endDate = DateTime.Now.AddDays(+OgsSettings.Instance.DaysInTheFuture + 1);

            // Set the start and end time and number of appointments to retrieve.
            CalendarView calendarView = new CalendarView(startDate, endDate, maxItemsReturned)
            {
                // Limit the properties returned to the appointment's subject, start time, and end time.
                PropertySet = new PropertySet(AppointmentSchema.Start,
                    AppointmentSchema.End, AppointmentSchema.IsRecurring, AppointmentSchema.AppointmentType)
            };

            FindItemsResults<Appointment> appointments = service.FindAppointments(calendar.Id, calendarView);
            if (appointments.Items.Count > 0)
            {
                foreach (Appointment appoint in appointments)
                {
                    appoint.Load();
                    var ai = new WebAppointmentItem
                    {
                        AllDayEvent = appoint.IsAllDayEvent,
                        Start = appoint.Start,
                        End = appoint.End,
                        Subject = appoint.Subject,
                        Location = appoint.Location,
                        Body = appoint.Body,
                        Organizer = appoint.Organizer.ToString(),
                        ReminderSet = appoint.IsReminderSet,
                        ReminderMinutesBeforeStart = appoint.ReminderMinutesBeforeStart,
                        RequiredAttendees = string.Join(";", appoint.RequiredAttendees),
                        OptionalAttendees = string.Join(";", appoint.OptionalAttendees),
                    };
                    appointmentsList.Add(ai);
                }
            }
            return appointmentsList;
        }
    }

    public class WebAppointmentItem
    {
        public bool AllDayEvent { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string Subject { get; set; }
        public string Location { get; set; }
        public string Body { get; set; }
        public string Organizer { get; set; }
        public bool ReminderSet { get; set; }
        public int? ReminderMinutesBeforeStart { get; set; }
        public string RequiredAttendees { get; set; }
        public string OptionalAttendees { get; set; }
    }
}

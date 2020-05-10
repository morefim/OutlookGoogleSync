using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;

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

        private ExchangeService ConnectToService()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1)
            {
                Credentials = new WebCredentials(OgsSettings.Instance.User, OgsSettings.Instance.OutlookPassword),
                Url = new Uri("https://outlook.office365.com/ews/exchange.asmx")
            };
            //service.AutodiscoverUrl(OgsSettings.Instance.User, RedirectionUrlValidationCallback);
            //service.UseDefaultCredentials = true;

            return service;
        }

        public List<WebAppointmentItem> GetCalendarEntriesInRange()
        {
            var appointmentsList = new List<WebAppointmentItem>();
            ExchangeService service = ConnectToService();

            const int maxItemsReturned = 50;

            // Initialize the calendar folder object with only the folder ID.
            CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar, new PropertySet());

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

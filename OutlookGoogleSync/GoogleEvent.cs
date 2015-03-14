using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleSync
{
    class GoogleEvent
    {
        private Event _event;

        public Event Event { get { return _event; } }

        public GoogleEvent(Event evt)
        {
            _event = evt;

            FixTime();
        }

        public GoogleEvent(AppointmentItem ai)
        {
            _event = new Event();

            _event.Start = new EventDateTime();
            _event.End = new EventDateTime();

            if (ai.AllDayEvent)
            {
                FixFullDayTime(ai.Start, ai.End);
            }
            else
            {
                _event.Start.DateTime = GoogleCalendar.Instance.GoogleTimeFrom(ai.Start);
                _event.End.DateTime = GoogleCalendar.Instance.GoogleTimeFrom(ai.End);
            }

            _event.Summary = ai.Subject;
            _event.Location = ai.Location;

            if (OGSSettings.Instance.AddDescription)
            {
                _event.Description = ai.Body;
            }

            if (OGSSettings.Instance.AddReminders)
            {
                //consider the reminder set in Outlook
                _event.Reminders = new Event.RemindersData();
                _event.Reminders.UseDefault = false;
                EventReminder reminder = new EventReminder();
                reminder.Method = "popup";
                reminder.Minutes = ai.ReminderMinutesBeforeStart;
                _event.Reminders.Overrides = new List<EventReminder>();
                _event.Reminders.Overrides.Add(reminder);
            }

            if (ai.Organizer != null)
            {
                if (_event.Description != null)
                    _event.Description += Environment.NewLine;

                _event.Description += "ORGANIZER: " + ai.Organizer;
                var organizer = EmailFromName(ai.Organizer);
                if (organizer != null)
                    _event.Description += " [" + organizer + "]";
            }

            if (OGSSettings.Instance.AddAttendeesToDescription)
            {
                if (_event.Description != null)
                    _event.Description += Environment.NewLine;
                if (ai.RequiredAttendees != null)
                {
                    _event.Description += Environment.NewLine + "REQUIRED: " + Environment.NewLine + splitAttendees(ai.RequiredAttendees);
                }
                if (ai.OptionalAttendees != null)
                {
                    _event.Description += Environment.NewLine + "OPTIONAL: " + Environment.NewLine + splitAttendees(ai.OptionalAttendees);
                }
            }

            if (ai.RequiredAttendees != null)
            {
                _event.Attendees = AtendeesToList(ai.RequiredAttendees.Split(';'), false);
            }

            if (ai.OptionalAttendees != null)
            {
                var attendees = AtendeesToList(ai.OptionalAttendees.Split(';'), true);
                if (_event.Attendees == null)
                    _event.Attendees = attendees;
                else
                {
                    foreach (var attendee in attendees)
                        _event.Attendees.Add(attendee);
                }
            }

            // set ID
            //_event.ICalUID = ai.EntryID;

            FixTime();
        }

        public IList<EventAttendee> AtendeesToList(string[] attendees, bool optional)
        {
            List<EventAttendee> attendeesList = new List<EventAttendee>();
            foreach (var attendee in attendees)
            {
                string email = EmailFromName(attendee);
                if (email == null)
                    continue;
                attendeesList.Add(new EventAttendee() { DisplayName = attendee, Email = email, Optional = optional });
            }
            if (attendeesList.Count > 0)
                return attendeesList;

            return null;
        }

        public string EmailFromName(string attendee)
        {
            attendee = attendee.Trim();
            RegexUtilities regex = new RegexUtilities();
            if (regex.IsValidEmail(attendee))
                return attendee;

            string[] name = attendee.Split(',');
            if (name.Count() == 2)
            {
                var email = name[1].ToLower().Trim() + "." + name[0].ToLower().Trim() + "@kla-tencor.com";                
                return regex.IsValidEmail(email) ? email : null;
            }
            return null;
        }

        //one attendee per line
        public string splitAttendees(string attendees)
        {
            if (attendees == null) return "";
            string[] tmp1 = attendees.Split(';');
            for (int i = 0; i < tmp1.Length; i++) tmp1[i] = tmp1[i].Trim();
            return String.Join(Environment.NewLine, tmp1);
        }

        public override string ToString()
        {
            string outString = string.Empty;

            if (_event == null)
                return outString;

            if (_event.Start != null && _event.Start.DateTime != null)
                outString += _event.Start.DateTime.Trim() + "; ";

            if (_event.End != null && _event.End.DateTime != null)
                outString += _event.End.DateTime.Trim() + "; ";

            if (_event.Summary != null)
                outString += _event.Summary.Trim() + "; ";

            if (_event.Location != null)
                outString += _event.Location.Trim() + "; ";

            if (OGSSettings.Instance.AddAttendeesToDescription)
            {
                if (_event.Description != null)
                    outString += _event.Description.Trim() + "; ";
            }

            return outString.Replace(Environment.NewLine, "; ").ToLower();
        }

        public void FixTime()
        {
            if (_event.Start.DateTime == null)
                _event.Start.DateTime = GoogleCalendar.Instance.GoogleTimeFrom(DateTime.Parse(_event.Start.Date));

            if (_event.End.DateTime == null)
                _event.End.DateTime = GoogleCalendar.Instance.GoogleTimeFrom(DateTime.Parse(_event.End.Date));
        }

        public void FixFullDayTime(DateTime start, DateTime end)
        {
            var curTimeZone = TimeZone.CurrentTimeZone;
            var dateStart = new DateTimeOffset(start, curTimeZone.GetUtcOffset(start));
            var dateEnd = new DateTimeOffset(end, curTimeZone.GetUtcOffset(end));

            var startTimeString = dateStart.ToString("o");
            var endTimeString = dateEnd.ToString("o");

            const string fractPart = ".0000000";
            endTimeString = endTimeString.Replace(fractPart, string.Empty);
            startTimeString = startTimeString.Replace(fractPart, string.Empty);

            _event.Start = new EventDateTime()
            {
                DateTime = startTimeString
            };

            _event.End = new EventDateTime()
            {
                DateTime = endTimeString
            };
        }

        public override bool Equals(object obj)
        {
            // If parameter cannot be cast to ThreeDPoint return false:
            GoogleEvent p = obj as GoogleEvent;
            if ((object)p == null)
            {
                return false;
            }

            // Return true if the fields match:
            if (!string.Equals(_event.Start.DateTime, p._event.Start.DateTime, StringComparison.InvariantCultureIgnoreCase))
                return false;

            if (!string.Equals(_event.End.DateTime, p._event.End.DateTime, StringComparison.InvariantCultureIgnoreCase))
                return false;

            if (!string.Equals(_event.Summary, p._event.Summary, StringComparison.InvariantCultureIgnoreCase))
                return false;

            if (!string.Equals(_event.Location, p._event.Location, StringComparison.InvariantCultureIgnoreCase))
                return false;

            if (OGSSettings.Instance.AddAttendeesToDescription)
            {
                if (!string.Equals(_event.Description, p._event.Description, StringComparison.InvariantCultureIgnoreCase))
                    return false;
            }

            return true;
        }

        public bool Equals(GoogleEvent p)
        {
            // Return true if the fields match:
            return Equals((object)p);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public static bool operator ==(GoogleEvent a, GoogleEvent b)
        {
            // If both are null, or both are same instance, return true.
            if (System.Object.ReferenceEquals(a, b))
            {
                return true;
            }

            // If one is null, but not both, return false.
            if (((object)a == null) || ((object)b == null))
            {
                return false;
            }

            // Return true if the fields match:
            return a.Equals(b);
        }

        public static bool operator !=(GoogleEvent a, GoogleEvent b)
        {
            return !(a == b);
        }
    }

    class GoogleEventsList : List<GoogleEvent>
    {
        public GoogleEventsList(List<Event> inList)
        {
            foreach (Event item in inList)
                Add(new GoogleEvent(item));
        }

        public GoogleEventsList(List<AppointmentItem> inList)
        {
            foreach (AppointmentItem item in inList)
                Add(new GoogleEvent(item));
        }
    }
}

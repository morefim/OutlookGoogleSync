﻿using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;
using TimeZone = System.TimeZone;

namespace OutlookGoogleSync
{
    public class GoogleEvent
    {
        public Event Event { get; }

        public GoogleEvent(Event evt)
        {
            Event = evt;

            FixTime();
        }

        public GoogleEvent(_AppointmentItem ai)
        {
            Event = new Event {Start = new EventDateTime(), End = new EventDateTime()};

            if (ai.AllDayEvent)
            {
                FixFullDayTime(ai.Start, ai.End);
            }
            else
            {
                Event.Start.DateTime = ai.Start;
                Event.End.DateTime = ai.End;
            }

            Event.Summary = ai.Subject;
            Event.Location = ai.Location;

            if (OgsSettings.Instance.AddDescription)
            {
                Event.Description = ai.Body;
            }

            if (OgsSettings.Instance.AddReminders)
            {
                //consider the reminder set in Outlook
                Event.Reminders = new Event.RemindersData {UseDefault = false};
                var reminder = new EventReminder {Method = "popup", Minutes = ai.ReminderSet ? ai.ReminderMinutesBeforeStart : 0};
                Event.Reminders.Overrides = new List<EventReminder> {reminder};
            }

            if (ai.Organizer != null)
            {
                if (Event.Description != null)
                    Event.Description += Environment.NewLine;

                Event.Description += "ORGANIZER: " + ai.Organizer;
                var organizer = EmailFromName(ai.Organizer);
                if (organizer != null)
                    Event.Description += " [" + organizer + "]";

                Event.Organizer = new Event.OrganizerData { Email = organizer, DisplayName = ai.Organizer };
            }

            if (OgsSettings.Instance.AddAttendeesToDescription)
            {
                if (Event.Description != null)
                    Event.Description += Environment.NewLine;
                if (ai.RequiredAttendees != null)
                {
                    Event.Description += Environment.NewLine + "REQUIRED:" + Environment.NewLine + SplitAttendees(ai.RequiredAttendees);
                }
                if (ai.OptionalAttendees != null)
                {
                    Event.Description += Environment.NewLine + "OPTIONAL:" + Environment.NewLine + SplitAttendees(ai.OptionalAttendees);
                }
            }

            if (ai.RequiredAttendees != null)
            {
                var attendees = AtendeesToList(ai.RequiredAttendees.Split(';'), false);
                Event.Attendees = attendees.HavingEmail;

                if (attendees.HavingNoEmail.Any())
                    Event.Description += Environment.NewLine + "REQUIRED(No Email):" + Environment.NewLine + attendees.HavingNoEmail.Aggregate((i, j) => i + Environment.NewLine + j);
            }

            if (ai.OptionalAttendees != null)
            {
                var attendees = AtendeesToList(ai.OptionalAttendees.Split(';'), true);
                if (Event.Attendees == null)
                    Event.Attendees = attendees.HavingEmail;
                else
                {
                    foreach (var attendee in attendees.HavingEmail)
                        Event.Attendees.Add(attendee);
                }

                if (attendees.HavingNoEmail.Any())
                    Event.Description += Environment.NewLine + "OPTIONAL(No Email):" + Environment.NewLine + attendees.HavingNoEmail.Aggregate((i, j) => i + Environment.NewLine + j);
            }

            // set ID
            //_event.ICalUID = ai.EntryID;

            FixTime();
        }

        public GoogleEvent(WebAppointmentItem ai)
        {
            Event = new Event { Start = new EventDateTime(), End = new EventDateTime() };

            if (ai.AllDayEvent)
            {
                FixFullDayTime(ai.Start, ai.End);
            }
            else
            {
                Event.Start.DateTime = ai.Start;
                Event.End.DateTime = ai.End;
            }

            Event.Summary = ai.Subject;
            Event.Location = ai.Location;

            if (OgsSettings.Instance.AddDescription)
            {
                Event.Description = ai.Body;
            }

            if (OgsSettings.Instance.AddReminders)
            {
                //consider the reminder set in Outlook
                Event.Reminders = new Event.RemindersData { UseDefault = false };
                var reminder = new EventReminder { Method = "popup", Minutes = ai.ReminderSet ? ai.ReminderMinutesBeforeStart : 0 };
                Event.Reminders.Overrides = new List<EventReminder> { reminder };
            }

            if (ai.Organizer != null)
            {
                if (Event.Description != null)
                    Event.Description += Environment.NewLine;

                Event.Description += "ORGANIZER: " + ai.Organizer;
                var organizer = EmailFromName(ai.Organizer);
                if (organizer != null)
                    Event.Description += " [" + organizer + "]";

                Event.Organizer = new Event.OrganizerData { Email = organizer, DisplayName = ai.Organizer };
            }

            if (OgsSettings.Instance.AddAttendeesToDescription)
            {
                if (Event.Description != null)
                    Event.Description += Environment.NewLine;
                if (ai.RequiredAttendees != null)
                {
                    Event.Description += Environment.NewLine + "REQUIRED:" + Environment.NewLine + SplitAttendees(ai.RequiredAttendees);
                }
                if (ai.OptionalAttendees != null)
                {
                    Event.Description += Environment.NewLine + "OPTIONAL:" + Environment.NewLine + SplitAttendees(ai.OptionalAttendees);
                }
            }

            if (ai.RequiredAttendees != null)
            {
                var attendees = AtendeesToList(ai.RequiredAttendees.Split(';'), false);
                Event.Attendees = attendees.HavingEmail;

                if (attendees.HavingNoEmail.Any())
                    Event.Description += Environment.NewLine + "REQUIRED(No Email):" + Environment.NewLine + attendees.HavingNoEmail.Aggregate((i, j) => i + Environment.NewLine + j);
            }

            if (ai.OptionalAttendees != null)
            {
                var attendees = AtendeesToList(ai.OptionalAttendees.Split(';'), true);
                if (Event.Attendees == null)
                    Event.Attendees = attendees.HavingEmail;
                else
                {
                    foreach (var attendee in attendees.HavingEmail)
                        Event.Attendees.Add(attendee);
                }

                if (attendees.HavingNoEmail.Any())
                    Event.Description += Environment.NewLine + "OPTIONAL(No Email):" + Environment.NewLine + attendees.HavingNoEmail.Aggregate((i, j) => i + Environment.NewLine + j);
            }

            FixTime();
        }

        public class Attendees
        {
            public IList<EventAttendee> HavingEmail = new List<EventAttendee>();
            public IList<string> HavingNoEmail = new List<string>();
        }

        public Attendees AtendeesToList(string[] attendees, bool optional)
        {
            var attendeesList = new Attendees();
            foreach (var attendee in attendees)
            {
                var email = EmailFromName(attendee);
                if (email == null)
                    attendeesList.HavingNoEmail.Add(attendee);
                else
                    attendeesList.HavingEmail.Add(new EventAttendee { DisplayName = attendee, Email = email, Optional = optional });
            }
            if (attendeesList.HavingEmail.Count > 0 && attendeesList.HavingEmail.Count < 100)
                return attendeesList;

            return attendeesList;
        }

        public string EmailFromName(string attendee)
        {
            attendee = attendee.Trim();
            var regex = new RegexUtilities();
            if (regex.IsValidEmail(attendee))
                return attendee;

            var name = attendee.Split(',');
            if (name.Length != 2) return null;
            var email = name[1].ToLower().Trim() + "." + name[0].ToLower().Trim() + "@kla-tencor.com";                
            return regex.IsValidEmail(email) ? email : null;
        }

        //one attendee per line
        public string SplitAttendees(string attendees)
        {
            if (attendees == null) return "";
            var tmp1 = attendees.Split(';');
            for (var i = 0; i < tmp1.Length; i++) tmp1[i] = tmp1[i].Trim();
            return string.Join(Environment.NewLine, tmp1);
        }

        public override string ToString()
        {
            var outString = string.Empty;

            if (Event == null)
                return outString;

            if (Event.Start != null && Event.Start.DateTime != null)
                outString += Event.Start.DateTime + "; ";

            if (Event.End != null && Event.End.DateTime != null)
                outString += Event.End.DateTime + "; ";

            if (Event.Summary != null)
                outString += Event.Summary.Trim() + "; ";

            if (Event.Location != null)
                outString += Event.Location.Trim() + "; ";

            if (OgsSettings.Instance.AddAttendeesToDescription)
            {
                if (Event.Description != null)
                    outString += Event.Description.Trim() + "; ";
            }

            return outString.Replace(Environment.NewLine, "; ").ToLower();
        }

        public void FixTime()
        {
            if (Event.Start.DateTime == null)
                Event.Start.DateTime = DateTime.Parse(Event.Start.Date);

            if (Event.End.DateTime == null)
                Event.End.DateTime = DateTime.Parse(Event.End.Date);
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

            Event.Start = new EventDateTime
            {
                DateTime = DateTime.Parse(startTimeString)
            };

            Event.End = new EventDateTime
            {
                DateTime = DateTime.Parse(endTimeString)
            };
        }

        public override bool Equals(object obj)
        {
            // If parameter cannot be cast to ThreeDPoint return false:
            var p = obj as GoogleEvent;
            if ((object)p == null)
            {
                return false;
            }

            // Return true if the fields match:
            if (!Event.Start.DateTime.Equals(p.Event.Start.DateTime))
                return false;

            if (!Event.End.DateTime.Equals(p.Event.End.DateTime))
                return false;

            if (!string.Equals(Event.Summary, p.Event.Summary, StringComparison.InvariantCultureIgnoreCase))
                return false;

            if (!string.Equals(Event.Location, p.Event.Location, StringComparison.InvariantCultureIgnoreCase))
                return false;

            if (OgsSettings.Instance.AddAttendeesToDescription)
            {
                if (!string.Equals(Event.Description, p.Event.Description, StringComparison.InvariantCultureIgnoreCase))
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
            // ReSharper disable once BaseObjectGetHashCodeCallInGetHashCode
            return base.GetHashCode();
        }

        public static bool operator ==(GoogleEvent a, GoogleEvent b)
        {
            // If both are null, or both are same instance, return true.
            if (ReferenceEquals(a, b))
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

    internal class GoogleEventsList : List<GoogleEvent>
    {
        public GoogleEventsList(IEnumerable<Event> inList)
        {
            foreach (var item in inList)
                Add(new GoogleEvent(item));
        }

        public GoogleEventsList(IEnumerable<AppointmentItem> inList)
        {
            using (new WarningSuppressor())
            {
                foreach (var item in inList)
                    Add(new GoogleEvent(item));
            }
        }

        public GoogleEventsList(List<WebAppointmentItem> inList)
        {
            using (new WarningSuppressor())
            {
                foreach (var item in inList)
                    Add(new GoogleEvent(item));
            }
        }
    }
}

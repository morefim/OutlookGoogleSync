﻿using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;
using TimeZone = System.TimeZone;

namespace OutlookGoogleSync
{
    class GoogleEvent
    {
        private readonly Event _event;

        public Event Event { get { return _event; } }

        public GoogleEvent(Event evt)
        {
            _event = evt;

            FixTime();
        }

        public GoogleEvent(_AppointmentItem ai)
        {
            _event = new Event {Start = new EventDateTime(), End = new EventDateTime()};

            if (ai.AllDayEvent)
            {
                FixFullDayTime(ai.Start, ai.End);
            }
            else
            {
                _event.Start.DateTime = ai.Start;
                _event.End.DateTime = ai.End;
            }

            _event.Summary = ai.Subject;
            _event.Location = ai.Location;

            if (OgsSettings.Instance.AddDescription)
            {
                _event.Description = ai.Body;
            }

            if (OgsSettings.Instance.AddReminders)
            {
                //consider the reminder set in Outlook
                _event.Reminders = new Event.RemindersData {UseDefault = false};
                var reminder = new EventReminder {Method = "popup", Minutes = ai.ReminderSet ? ai.ReminderMinutesBeforeStart : 0};
                _event.Reminders.Overrides = new List<EventReminder> {reminder};
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

            if (OgsSettings.Instance.AddAttendeesToDescription)
            {
                if (_event.Description != null)
                    _event.Description += Environment.NewLine;
                if (ai.RequiredAttendees != null)
                {
                    _event.Description += Environment.NewLine + "REQUIRED:" + Environment.NewLine + SplitAttendees(ai.RequiredAttendees);
                }
                if (ai.OptionalAttendees != null)
                {
                    _event.Description += Environment.NewLine + "OPTIONAL:" + Environment.NewLine + SplitAttendees(ai.OptionalAttendees);
                }
            }

            if (ai.RequiredAttendees != null)
            {
                var attendees = AtendeesToList(ai.RequiredAttendees.Split(';'), false);
                _event.Attendees = attendees.HavingEmail;

                if (attendees.HavingNoEmail.Any())
                    _event.Description += Environment.NewLine + "REQUIRED(No Email):" + Environment.NewLine + attendees.HavingNoEmail.Aggregate((i, j) => i + Environment.NewLine + j);
            }

            if (ai.OptionalAttendees != null)
            {
                var attendees = AtendeesToList(ai.OptionalAttendees.Split(';'), true);
                if (_event.Attendees == null)
                    _event.Attendees = attendees.HavingEmail;
                else
                {
                    foreach (var attendee in attendees.HavingEmail)
                        _event.Attendees.Add(attendee);
                }

                if (attendees.HavingNoEmail.Any())
                    _event.Description += Environment.NewLine + "OPTIONAL(No Email):" + Environment.NewLine + attendees.HavingNoEmail.Aggregate((i, j) => i + Environment.NewLine + j);
            }

            // set ID
            //_event.ICalUID = ai.EntryID;

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
            return String.Join(Environment.NewLine, tmp1);
        }

        public override string ToString()
        {
            var outString = string.Empty;

            if (_event == null)
                return outString;

            if (_event.Start != null && _event.Start.DateTime != null)
                outString += _event.Start.DateTime + "; ";

            if (_event.End != null && _event.End.DateTime != null)
                outString += _event.End.DateTime + "; ";

            if (_event.Summary != null)
                outString += _event.Summary.Trim() + "; ";

            if (_event.Location != null)
                outString += _event.Location.Trim() + "; ";

            if (OgsSettings.Instance.AddAttendeesToDescription)
            {
                if (_event.Description != null)
                    outString += _event.Description.Trim() + "; ";
            }

            return outString.Replace(Environment.NewLine, "; ").ToLower();
        }

        public void FixTime()
        {
            if (_event.Start.DateTime == null)
                _event.Start.DateTime = DateTime.Parse(_event.Start.Date);

            if (_event.End.DateTime == null)
                _event.End.DateTime = DateTime.Parse(_event.End.Date);
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

            _event.Start = new EventDateTime
            {
                DateTime = DateTime.Parse(startTimeString)
            };

            _event.End = new EventDateTime
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
            if (!_event.Start.DateTime.Equals(p._event.Start.DateTime))
                return false;

            if (!_event.End.DateTime.Equals(p._event.End.DateTime))
                return false;

            if (!string.Equals(_event.Summary, p._event.Summary, StringComparison.InvariantCultureIgnoreCase))
                return false;

            if (!string.Equals(_event.Location, p._event.Location, StringComparison.InvariantCultureIgnoreCase))
                return false;

            if (OgsSettings.Instance.AddAttendeesToDescription)
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
            foreach (var item in inList)
                Add(new GoogleEvent(item));
        }
    }
}

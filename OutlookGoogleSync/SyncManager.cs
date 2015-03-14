using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using System.IO;
using System.Threading;
using System.Diagnostics;

namespace OutlookGoogleSync
{
    class SyncManager
    {
        public delegate void SyncDoneDelegate(int deleted, int created);
        public delegate void LogboxOutDelegate(string text);
        public delegate void HandleExceptionDelegate(System.Exception e);

        public event LogboxOutDelegate OnLogboxOutDelegate;
        public event HandleExceptionDelegate OnExceptionDelegate;
        public event SyncDoneDelegate OnSyncDoneDelegate;

        public void DoWork()
        {
            ThreadPool.QueueUserWorkItem(DoWorkCallback, this);
        }

        private void DoWorkCallback(Object threadContext)
        {
            int max_attempts = 3;
            for (int attempt = 0; attempt < max_attempts; attempt++)
            {
                try
                {
                    PrivateDoWorkCallback(threadContext);
                    break;
                }
                catch (System.Exception e)
                {
                    if (OnExceptionDelegate != null && attempt + 1 >= max_attempts)
                        OnExceptionDelegate(e);
                }
            }
        }

        private void PrivateDoWorkCallback(Object threadContext)
        {
            DateTime SyncStarted = DateTime.Now;
            bool cbCreateFilesChecked = OGSSettings.Instance.CreateTextFiles;
            bool cbAddDescriptionChecked = OGSSettings.Instance.AddDescription;
            bool cbAddRemindersChecked = OGSSettings.Instance.AddReminders;
            bool cbAddAttendeesChecked = OGSSettings.Instance.AddAttendeesToDescription;

            OnLogboxOutDelegate("Sync started at " + SyncStarted.ToString());
            OnLogboxOutDelegate("--------------------------------------------------");

            OnLogboxOutDelegate("Reading Outlook Calendar Entries...");
            GoogleEventsList OutlookEntries = new GoogleEventsList(OutlookCalendar.Instance.getCalendarEntriesInRange());
            if (cbCreateFilesChecked)
            {
                using (TextWriter tw = new StreamWriter("export_found_in_outlook.txt"))
                {
                    foreach (GoogleEvent ai in OutlookEntries)
                    {
                        tw.WriteLine(ai.ToString());
                    }
                    tw.Close();
                }
            }
            OnLogboxOutDelegate("Found " + OutlookEntries.Count + " Outlook Calendar Entries.");
            OnLogboxOutDelegate("--------------------------------------------------");

            OnLogboxOutDelegate("Reading Google Calendar Entries...");
            GoogleEventsList GoogleEntries = new GoogleEventsList(GoogleCalendar.Instance.getCalendarEntriesInRange());
            if (cbCreateFilesChecked)
            {
                using (TextWriter tw = new StreamWriter("export_found_in_google.txt"))
                {
                    foreach (GoogleEvent ev in GoogleEntries)
                    {
                        tw.WriteLine(ev.ToString());
                    }
                    tw.Close();
                }
            }
            OnLogboxOutDelegate("Found " + GoogleEntries.Count + " Google Calendar Entries.");
            OnLogboxOutDelegate("--------------------------------------------------");

            List<GoogleEvent> GoogleEntriesToBeDeleted = IdentifyGoogleEntriesToBeDeleted(OutlookEntries, GoogleEntries);
            if (cbCreateFilesChecked)
            {
                using (TextWriter tw = new StreamWriter("export_to_be_deleted.txt"))
                {
                    foreach (GoogleEvent ev in GoogleEntriesToBeDeleted)
                    {
                        tw.WriteLine(ev.ToString());
                    }
                    tw.Close();
                }
            }
            OnLogboxOutDelegate(GoogleEntriesToBeDeleted.Count + " Google Calendar Entries to be deleted.");

            //OutlookEntriesToBeCreated ...in Google!
            List<GoogleEvent> OutlookEntriesToBeCreated = IdentifyOutlookEntriesToBeCreated(OutlookEntries, GoogleEntries);
            if (cbCreateFilesChecked)
            {
                using (TextWriter tw = new StreamWriter("export_to_be_created.txt"))
                {
                    foreach (GoogleEvent ai in OutlookEntriesToBeCreated)
                    {
                        tw.WriteLine(ai.ToString());
                    }
                    tw.Close();
                }
            }
            OnLogboxOutDelegate(OutlookEntriesToBeCreated.Count + " Entries to be created in Google.");
            OnLogboxOutDelegate("--------------------------------------------------");

            if (GoogleEntriesToBeDeleted.Count > 0)
            {
                OnLogboxOutDelegate("Deleting " + GoogleEntriesToBeDeleted.Count + " Google Calendar Entries...");
                foreach (GoogleEvent ev in GoogleEntriesToBeDeleted)
                {
                    OnLogboxOutDelegate(string.Format(" {0}", ev.ToString()));
                    GoogleCalendar.Instance.deleteCalendarEntry(ev.Event);
                }
                OnLogboxOutDelegate("Done.");
                OnLogboxOutDelegate("--------------------------------------------------");
            }

            if (OutlookEntriesToBeCreated.Count > 0)
            {
                OnLogboxOutDelegate("Creating " + OutlookEntriesToBeCreated.Count + " Entries in Google...");
                foreach (GoogleEvent ai in OutlookEntriesToBeCreated)
                {
                    OnLogboxOutDelegate(string.Format(" {0}", ai.ToString()));
                    GoogleCalendar.Instance.addEntry(ai.Event);
                }
                OnLogboxOutDelegate("Done.");
                OnLogboxOutDelegate("--------------------------------------------------");
            }

            DateTime SyncFinished = DateTime.Now;
            TimeSpan Elapsed = SyncFinished - SyncStarted;
            OnLogboxOutDelegate("Sync finished at " + SyncFinished.ToString());
            OnLogboxOutDelegate("Time needed: " + Elapsed.Minutes + " min " + Elapsed.Seconds + " s");
            OnSyncDoneDelegate(GoogleEntriesToBeDeleted.Count, OutlookEntriesToBeCreated.Count);
        }

        public List<GoogleEvent> IdentifyGoogleEntriesToBeDeleted(List<GoogleEvent> outlook, List<GoogleEvent> google)
        {
            List<GoogleEvent> result = new List<GoogleEvent>();
            foreach (GoogleEvent g in google)
            {
                bool found = outlook.Exists(o => g.Equals(o));
                if (!found)
                {
                    //Debug.Assert(false);
                    result.Add(g);
                }

                // find duplicates
                found = google.FindAll(o => g.Equals(o)).Count > 1;
                bool added = result.Exists(r => g.Equals(r));
                if (found & !added)
                {
                    //Debug.Assert(false);
                    result.Add(g);
                }
            }
            return result;
        }

        public List<GoogleEvent> IdentifyOutlookEntriesToBeCreated(List<GoogleEvent> outlook, List<GoogleEvent> google)
        {
            List<GoogleEvent> result = new List<GoogleEvent>();
            foreach (GoogleEvent o in outlook)
            {
                bool found = google.Exists(g => g.Equals(o));

                if (!found)
                {
                    //Debug.Assert(false);
                    result.Add(o);
                }
            }
            return result;
        }
    }
}

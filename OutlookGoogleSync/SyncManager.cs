﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace OutlookGoogleSync
{
    public class SyncManager
    {
        public delegate void SyncDoneDelegate(int deleted, int created);
        public delegate void LogboxOutDelegate(string text);
        public delegate void HandleExceptionDelegate(Exception e);

        public event LogboxOutDelegate OnLogboxOutDelegate;
        public event HandleExceptionDelegate OnExceptionDelegate;
        public event SyncDoneDelegate OnSyncDoneDelegate;

        public SyncManager()
        {
            
        }

        public async Task<bool> DoWork()
        {
            const int maxAttempts = 3;
            for (var attempt = 0; attempt < maxAttempts; attempt++)
            {
                try
                {
                    return await PrivateDoWorkCallback();
                }
                catch (Exception e)
                {
                    if (OnExceptionDelegate != null && attempt + 1 >= maxAttempts)
                        OnExceptionDelegate(e);
                }
            }

            return await Task.FromResult(false);
        }

        private async Task<bool> PrivateDoWorkCallback()
        {
            var syncStarted = DateTime.Now;

            OnLogboxOutDelegate("Sync started at " + syncStarted);
            OnLogboxOutDelegate("--------------------------------------------------");

            OnLogboxOutDelegate("Reading Outlook Calendar Entries...");
            //var outlookEntries = new GoogleEventsList(OutlookCalendar.Instance.GetCalendarEntriesInRange());
            var outlookEntries = new GoogleEventsList(OutlookWebCalendar.Instance.GetCalendarEntriesInRange());
            ExportToLogFile("found_in_outlook", outlookEntries);

            OnLogboxOutDelegate("Found " + outlookEntries.Count + " Outlook Calendar Entries.");
            OnLogboxOutDelegate("--------------------------------------------------");

            OnLogboxOutDelegate("Reading Google Calendar Entries...");
            var googleEntries = new GoogleEventsList(GoogleCalendar.Instance.GetCalendarEntriesInRange());
            ExportToLogFile("found_in_google", googleEntries);

            OnLogboxOutDelegate("Found " + googleEntries.Count + " Google Calendar Entries.");
            OnLogboxOutDelegate("--------------------------------------------------");

            var googleEntriesToBeDeleted = IdentifyGoogleEntriesToBeDeleted(outlookEntries, googleEntries);
            ExportToLogFile("to_be_deleted", googleEntriesToBeDeleted);
            OnLogboxOutDelegate(googleEntriesToBeDeleted.Count + " Google Calendar Entries to be deleted.");

            //OutlookEntriesToBeCreated ...in Google!
            var outlookEntriesToBeCreated = IdentifyOutlookEntriesToBeCreated(outlookEntries, googleEntries);
            ExportToLogFile("to_be_created", outlookEntriesToBeCreated);

            OnLogboxOutDelegate(outlookEntriesToBeCreated.Count + " Entries to be created in Google.");
            OnLogboxOutDelegate("--------------------------------------------------");

            if (googleEntriesToBeDeleted.Count > 0)
            {
                OnLogboxOutDelegate("Deleting " + googleEntriesToBeDeleted.Count + " Google Calendar Entries...");
                foreach (var ev in googleEntriesToBeDeleted)
                {
                    OnLogboxOutDelegate($" {ev}");
                    GoogleCalendar.Instance.DeleteCalendarEntry(ev.Event);
                }
                OnLogboxOutDelegate("Done.");
                OnLogboxOutDelegate("--------------------------------------------------");
            }

            if (outlookEntriesToBeCreated.Count > 0)
            {
                OnLogboxOutDelegate("Creating " + outlookEntriesToBeCreated.Count + " Entries in Google...");
                foreach (var ai in outlookEntriesToBeCreated)
                {
                    OnLogboxOutDelegate($" {ai}");
                    GoogleCalendar.Instance.AddEntry(ai.Event);
                }
                OnLogboxOutDelegate("Done.");
                OnLogboxOutDelegate("--------------------------------------------------");
            }

            var syncFinished = DateTime.Now;
            var elapsed = syncFinished - syncStarted;
            OnLogboxOutDelegate("Sync finished at " + syncFinished);
            OnLogboxOutDelegate("Time needed: " + elapsed.Minutes + " min " + elapsed.Seconds + " s");
            OnSyncDoneDelegate(googleEntriesToBeDeleted.Count, outlookEntriesToBeCreated.Count);

            return await Task.FromResult(true);
        }

        private static void ExportToLogFile(string fileName, List<GoogleEvent> googleEntriesToBeDeleted)
        {
            var cbCreateFilesChecked = OgsSettings.Instance.CreateTextFiles;
            if (!cbCreateFilesChecked) return;

            using (TextWriter tw = new StreamWriter($"{fileName}.txt"))
            {
                foreach (var ev in googleEntriesToBeDeleted)
                {
                    tw.WriteLine(ev.ToString());
                }

                tw.Close();
            }
        }

        public List<GoogleEvent> IdentifyGoogleEntriesToBeDeleted(List<GoogleEvent> outlook, List<GoogleEvent> google)
        {
            var result = new List<GoogleEvent>();
            foreach (var g in google)
            {
                // find obsolete events
                var found = outlook.Exists(o => g.Equals(o));
                if (!found)
                {
                    //Debug.Assert(false);
                    result.Add(g);
                }

                // find duplicates
                found = google.FindAll(o => g.Equals(o)).Count > 1;
                var added = result.Exists(r => g.Equals(r));
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
            return (from o in outlook let found = google.Exists(g => g.Equals(o)) where !found select o).ToList();
        }
    }
}

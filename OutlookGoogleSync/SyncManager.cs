using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace OutlookGoogleSync
{
    class SyncManager
    {
        public delegate void SyncDoneDelegate(int deleted, int created);
        public delegate void LogboxOutDelegate(string text);
        public delegate void HandleExceptionDelegate(Exception e);

        public event LogboxOutDelegate OnLogboxOutDelegate;
        public event HandleExceptionDelegate OnExceptionDelegate;
        public event SyncDoneDelegate OnSyncDoneDelegate;

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
            var cbCreateFilesChecked = OGSSettings.Instance.CreateTextFiles;

            OnLogboxOutDelegate("Sync started at " + syncStarted);
            OnLogboxOutDelegate("--------------------------------------------------");

            OnLogboxOutDelegate("Reading Outlook Calendar Entries...");
            var outlookEntries = new GoogleEventsList(OutlookCalendar.Instance.getCalendarEntriesInRange());
            if (cbCreateFilesChecked)
            {
                using (TextWriter tw = new StreamWriter("export_found_in_outlook.txt"))
                {
                    foreach (var ai in outlookEntries)
                    {
                        tw.WriteLine(ai.ToString());
                    }
                    tw.Close();
                }
            }
            OnLogboxOutDelegate("Found " + outlookEntries.Count + " Outlook Calendar Entries.");
            OnLogboxOutDelegate("--------------------------------------------------");

            OnLogboxOutDelegate("Reading Google Calendar Entries...");
            var googleEntries = new GoogleEventsList(GoogleCalendar.Instance.getCalendarEntriesInRange());
            if (cbCreateFilesChecked)
            {
                using (TextWriter tw = new StreamWriter("export_found_in_google.txt"))
                {
                    foreach (var ev in googleEntries)
                    {
                        tw.WriteLine(ev.ToString());
                    }
                    tw.Close();
                }
            }
            OnLogboxOutDelegate("Found " + googleEntries.Count + " Google Calendar Entries.");
            OnLogboxOutDelegate("--------------------------------------------------");

            var googleEntriesToBeDeleted = IdentifyGoogleEntriesToBeDeleted(outlookEntries, googleEntries);
            if (cbCreateFilesChecked)
            {
                using (TextWriter tw = new StreamWriter("export_to_be_deleted.txt"))
                {
                    foreach (var ev in googleEntriesToBeDeleted)
                    {
                        tw.WriteLine(ev.ToString());
                    }
                    tw.Close();
                }
            }
            OnLogboxOutDelegate(googleEntriesToBeDeleted.Count + " Google Calendar Entries to be deleted.");

            //OutlookEntriesToBeCreated ...in Google!
            var outlookEntriesToBeCreated = IdentifyOutlookEntriesToBeCreated(outlookEntries, googleEntries);
            if (cbCreateFilesChecked)
            {
                using (TextWriter tw = new StreamWriter("export_to_be_created.txt"))
                {
                    foreach (var ai in outlookEntriesToBeCreated)
                    {
                        tw.WriteLine(ai.ToString());
                    }
                    tw.Close();
                }
            }
            OnLogboxOutDelegate(outlookEntriesToBeCreated.Count + " Entries to be created in Google.");
            OnLogboxOutDelegate("--------------------------------------------------");

            if (googleEntriesToBeDeleted.Count > 0)
            {
                OnLogboxOutDelegate("Deleting " + googleEntriesToBeDeleted.Count + " Google Calendar Entries...");
                foreach (var ev in googleEntriesToBeDeleted)
                {
                    OnLogboxOutDelegate(string.Format(" {0}", ev));
                    GoogleCalendar.Instance.deleteCalendarEntry(ev.Event);
                }
                OnLogboxOutDelegate("Done.");
                OnLogboxOutDelegate("--------------------------------------------------");
            }

            if (outlookEntriesToBeCreated.Count > 0)
            {
                OnLogboxOutDelegate("Creating " + outlookEntriesToBeCreated.Count + " Entries in Google...");
                foreach (var ai in outlookEntriesToBeCreated)
                {
                    OnLogboxOutDelegate(string.Format(" {0}", ai));
                    GoogleCalendar.Instance.addEntry(ai.Event);
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

        public List<GoogleEvent> IdentifyGoogleEntriesToBeDeleted(List<GoogleEvent> outlook, List<GoogleEvent> google)
        {
            var result = new List<GoogleEvent>();
            foreach (var g in google)
            {
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

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OutlookGoogleSyncClient.Service
{
    public interface IOutlookGoogleSyncSettings
    {
        OutlookGoogleSyncSettings.GoogleCalendarEntry GoogleCalendar { get; set; }
        int DaysInThePast { get; set; }
        int DaysInTheFuture { get; set; }
        void Save();
    }

    public class OutlookGoogleSyncSettings : IOutlookGoogleSyncSettings, IDisposable
    {
        public class GoogleCalendarEntry
        {
            public string Id { get; set; }
            public string Name { get; set; }
        }

        public GoogleCalendarEntry GoogleCalendar { get; set; }
        public int DaysInThePast { get; set; }
        public int DaysInTheFuture { get; set; }

        static readonly string SettingsFilePath = Path.GetFullPath(typeof(OutlookGoogleSyncSettings).Name + ".json");

        public static OutlookGoogleSyncSettings Load()
        {
            if (!File.Exists(SettingsFilePath))
                new OutlookGoogleSyncSettings().Save();

            using (var file = File.OpenText(SettingsFilePath))
            {
                using (var reader = new JsonTextReader(file))
                {
                    var serializer = new JsonSerializer();
                    return serializer.Deserialize<OutlookGoogleSyncSettings>(reader);
                }
            }
        }

        public void Save()
        {
            using (var file = File.CreateText(SettingsFilePath))
            {
                using (var writer = new JsonTextWriter(file))
                {
                    var serializer = new JsonSerializer();
                    serializer.Serialize(writer, this);
                }
            }
        }

        public void Dispose()
        {
            Save();
        }
    }
}

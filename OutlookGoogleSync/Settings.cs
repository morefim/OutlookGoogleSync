using System;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of Settings.
    /// </summary>
    [Serializable, XmlRoot("Settings")]
    public class OgsSettings
    {
        [XmlIgnore]
        private static OgsSettings _instance;

        [XmlIgnore]
        public static OgsSettings Instance
        {
            get 
            {
                if (_instance == null)
                    _instance = new OgsSettings();
                return _instance;
            }
            set => _instance = value;
        }

        public bool Autostart
        {
            get => Shortcut.IsStartupFolderShortcutExists();

            set
            {
                try
                {
                    if (value)
                        Shortcut.CreateStartupFolderShortcut();
                    else
                        Shortcut.DeleteStartupFolderShortcuts(AppDomain.CurrentDomain.FriendlyName);
                }
                catch { }
            }
        }

        [XmlIgnore]
        public string OutlookPassword => Coder.Decrypt(Password);


        [OptionalField(VersionAdded = 2)]
        public string RefreshToken = "";

        [OptionalField(VersionAdded = 2)]
        public int DaysInThePast = 1;

        [OptionalField(VersionAdded = 2)]
        public int DaysInTheFuture = 7;

        [OptionalField(VersionAdded = 2)]
        public OgsCalendarListEntry UseGoogleCalendar = new OgsCalendarListEntry();

        [OptionalField(VersionAdded = 3)]
        public bool SyncEvery = true;

        [OptionalField(VersionAdded = 3)]
        public int SyncPeriod = 60;

        [OptionalField(VersionAdded = 2)]
        public bool ShowBubbleTooltipWhenSyncing = false;

        [OptionalField(VersionAdded = 2)]
        public bool StartInTray = true;

        [OptionalField(VersionAdded = 2)]
        public bool MinimizeToTray = true;

        [OptionalField(VersionAdded = 2)]
        public bool AddDescription = false;

        [OptionalField(VersionAdded = 2)]
        public bool AddReminders = true;

        [OptionalField(VersionAdded = 2)]
        public bool AddAttendeesToDescription = false;

        [OptionalField(VersionAdded = 2)]
        public bool CreateTextFiles = false;

        [OptionalField(VersionAdded = 2)]
        public string User = "";

        [OptionalField(VersionAdded = 2)]
        public string Password = "";

        [OptionalField(VersionAdded = 4)]
        public bool ByPassKLALogoutPolicy = false;
    }
}

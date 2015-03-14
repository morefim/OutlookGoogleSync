using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of Settings.
    /// </summary>
    [Serializable, XmlRoot("Settings")]
    public class OGSSettings
    {
        [XmlIgnore]
        private static OGSSettings instance;

        [XmlIgnore]
        public static OGSSettings Instance
        {
            get 
            {
                if (instance == null)
                    instance = new OGSSettings();
                return instance;
            }
            set
            {
                instance = value;            
            }          
        }

        public bool Autostart
        {
            get 
            {
                return Shortcut.IsStartupFolderShortcutExists();
            }

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
        public string OutlookPassword 
        { 
            get 
            { 
                return Coder.Decrypt(Password); 
            } 
        }


        [OptionalField(VersionAdded = 2)]
        public string RefreshToken = "";

        [OptionalField(VersionAdded = 2)]
        public int DaysInThePast = 1;

        [OptionalField(VersionAdded = 2)]
        public int DaysInTheFuture = 7;

        [OptionalField(VersionAdded = 2)]
        public OGSCalendarListEntry UseGoogleCalendar = new OGSCalendarListEntry();


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
    }
}

//TODO: consider description updates?
//TODO: optimize comparison algorithms
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of MainForm.
    /// </summary>
    public partial class MainForm : Form
    {
        public static MainForm Instance;

        public const string FILENAME = "settings.xml";
        public const string VERSION = "1.0.8";

        public Timer ogstimer;
        public DateTime oldtime;

        SyncManager _syncManager = new SyncManager();

        public MainForm()
        {
            InitializeComponent();

            Instance = this;

            _syncManager.OnLogboxOutDelegate += logboxout;
            _syncManager.OnExceptionDelegate += HandleException;
            _syncManager.OnSyncDoneDelegate += SyncDone;

            //set system proxy
            System.Net.WebRequest.DefaultWebProxy = null; 

            //load settings/create settings file
            try
            {
                if (File.Exists(FILENAME))
                    OGSSettings.Instance = XMLManager.import<OGSSettings>(FILENAME);
                else
                    XMLManager.export(OGSSettings.Instance, FILENAME);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //update GUI from Settings
            tbDaysInThePast.Text = OGSSettings.Instance.DaysInThePast.ToString();
            tbDaysInTheFuture.Text = OGSSettings.Instance.DaysInTheFuture.ToString();
            cbCalendars.Items.Add(OGSSettings.Instance.UseGoogleCalendar);
            cbCalendars.SelectedIndex = 0;
            cbSyncEvery.Checked = OGSSettings.Instance.SyncEvery;
            tbSyncPeriod.Text = OGSSettings.Instance.SyncPeriod.ToString();
            cbShowBubbleTooltips.Checked = OGSSettings.Instance.ShowBubbleTooltipWhenSyncing;
            cbStartInTray.Checked = OGSSettings.Instance.StartInTray;
            cbMinimizeToTray.Checked = OGSSettings.Instance.MinimizeToTray;
            cbAddDescription.Checked = OGSSettings.Instance.AddDescription;
            cbAddAttendees.Checked = OGSSettings.Instance.AddAttendeesToDescription;
            cbAddReminders.Checked = OGSSettings.Instance.AddReminders;
            cbCreateFiles.Checked = OGSSettings.Instance.CreateTextFiles;
            cbStartWithWindows.Checked = OGSSettings.Instance.Autostart;
            tbOutlookUser.Text = OGSSettings.Instance.User;
            tbOutlookPassword.Text = OGSSettings.Instance.OutlookPassword;
            tbOutlookPassword.PasswordChar = '*';

            //set up timer (every 30s) for checking the minute offsets
            ogstimer = new Timer();
            ogstimer.Interval = (int)TimeSpan.FromSeconds(30).TotalMilliseconds;
            ogstimer.Tick += new EventHandler(ogstimer_Tick);
            ogstimer.Start();
            oldtime = DateTime.Now.RoundDown(OGSSettings.Instance.SyncPeriod);

            //set up tooltips for some controls
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 10000;
            toolTip1.InitialDelay = 500;
            toolTip1.ReshowDelay = 200;
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(cbCalendars,
                "The Google Calendar to synchonize with.");
            toolTip1.SetToolTip(cbAddAttendees,
                "While Outlook has fields for Organizer, RequiredAttendees and OptionalAttendees, Google has not.\n" +
                "If checked, this data is added at the end of the description as text.");
            toolTip1.SetToolTip(cbAddReminders,
                "If checked, the reminder set in outlook will be carried over to the Google Calendar entry (as a popup reminder).");
            toolTip1.SetToolTip(cbCreateFiles,
                "If checked, all entries found in Outlook/Google and identified for creation/deletion will be exported \n" +
                "to 4 separate text files in the application's directory (named \"export_*.txt\"). \n" +
                "Only for debug/diagnostic purposes.");
            toolTip1.SetToolTip(cbAddDescription,
                "The description may contain email addresses, which Outlook may complain about (PopUp-Message: \"Allow Access?\" etc.). \n" +
                "Turning this off allows OutlookGoogleSync to run without intervention in this case.");
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            //Start in tray?
            if (cbStartInTray.Checked)
            {
                this.WindowState = FormWindowState.Minimized;
                notifyIcon.Visible = true;
                this.ShowInTaskbar = false;
                this.Hide();
            }
        }

        void ogstimer_Tick(object sender, EventArgs e)
        {
            if (!cbSyncEvery.Checked) return;

            DateTime newtime = DateTime.Now;
            if ((newtime - oldtime) < new TimeSpan(0, OGSSettings.Instance.SyncPeriod, 0)) 
                return;

            oldtime = newtime;
            SyncNow_Click(null, null);
        }

        void SyncDone(int deleted, int created)
        {
            if (cbShowBubbleTooltips.Checked)
            {
                notifyIcon.BalloonTipTitle = "Outlook Google Sync";
                notifyIcon.BalloonTipText = "Syncronization complete";
                if (created > 0)
                    notifyIcon.BalloonTipText += string.Format(", {0} created", created);
                if (deleted > 0)
                    notifyIcon.BalloonTipText += string.Format(", {0} deleted", deleted);
                notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
                notifyIcon.Text = ("OGS: " + notifyIcon.BalloonTipText).Truncate(63);
                notifyIcon.ShowBalloonTip(500);
            }
        }

        void GetMyGoogleCalendars_Click(object sender, EventArgs e)
        {
            bGetMyCalendars.Enabled = false;
            cbCalendars.Enabled = false;

            List<OGSCalendarListEntry> calendars = GoogleCalendar.Instance.getCalendars();
            if (calendars != null)
            {
                cbCalendars.Items.Clear();
                foreach (OGSCalendarListEntry mcle in calendars)
                {
                    cbCalendars.Items.Add(mcle);
                }
                MainForm.Instance.cbCalendars.SelectedIndex = 0;
            }

            bGetMyCalendars.Enabled = true;
            cbCalendars.Enabled = true;
        }

        void SyncNow_Click(object sender, EventArgs e)
        {
            if (OGSSettings.Instance.UseGoogleCalendar.IsEmpty)
            {
                MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
                return;
            }

            bSyncNow.Enabled = false;

            LogBox.Clear();

            _syncManager.DoWork();

            bSyncNow.Enabled = true;
        }

        void logboxout(string s)
        {
            if (LogBox.InvokeRequired)
                Invoke((Action<string>)logboxout, new object[] { s });
            else
                LogBox.Text += s + Environment.NewLine;
        }

        public void HandleException(Exception ex)
        {
            if (InvokeRequired)
                Invoke((Action<Exception>)HandleException, new object[] { ex });
            else
            {
                try
                {
                    if (cbShowBubbleTooltips.Checked)
                    {
                        notifyIcon.BalloonTipTitle = "Outlook Google Sync";
                        notifyIcon.BalloonTipText = "Exception: " + ex.Message;
                        notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
                        notifyIcon.Text = ("OGS: " + notifyIcon.BalloonTipText).Truncate(63);
                        notifyIcon.ShowBalloonTip(500);
                    }

                    logboxout("Exception: " + ex.Message);
                    using (TextWriter tw = new StreamWriter("exception.txt"))
                    {
                        tw.WriteLine(DateTime.Now.ToString() + "\t" + ex.ToString());
                        tw.Close();
                    }
                }
                catch { }
            }
        }

        void Save_Click(object sender, EventArgs e)
        {
            XMLManager.export(OGSSettings.Instance, FILENAME);
        }

        private void tbSyncPeriod_TextChanged(object sender, EventArgs e)
        {
            int.TryParse(tbSyncPeriod.Text, out OGSSettings.Instance.SyncPeriod);
            if (OGSSettings.Instance.SyncPeriod < 1)
                OGSSettings.Instance.SyncPeriod = 1;
        }

        private void tbOutlookUser_TextChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.User = tbOutlookUser.Text;
        }

        private void tbOutlookPassword_TextChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.Password = Coder.Encrypt(tbOutlookPassword.Text);
        }

        void ComboBox1SelectedIndexChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.UseGoogleCalendar = (OGSCalendarListEntry)cbCalendars.SelectedItem;
        }

        void TbDaysInThePastTextChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.DaysInThePast = int.Parse(tbDaysInThePast.Text);
        }

        void TbDaysInTheFutureTextChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.DaysInTheFuture = int.Parse(tbDaysInTheFuture.Text);
        }

        void CbSyncEveryHourCheckedChanged(object sender, System.EventArgs e)
        {
            OGSSettings.Instance.SyncEvery = cbSyncEvery.Checked;
        }

        void CbShowBubbleTooltipsCheckedChanged(object sender, System.EventArgs e)
        {
            OGSSettings.Instance.ShowBubbleTooltipWhenSyncing = cbShowBubbleTooltips.Checked;
        }

        void CbStartInTrayCheckedChanged(object sender, System.EventArgs e)
        {
            OGSSettings.Instance.StartInTray = cbStartInTray.Checked;
        }

        void CbMinimizeToTrayCheckedChanged(object sender, System.EventArgs e)
        {
            OGSSettings.Instance.MinimizeToTray = cbMinimizeToTray.Checked;
        }

        void CbAddDescriptionCheckedChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.AddDescription = cbAddDescription.Checked;
        }

        void CbAddRemindersCheckedChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.AddReminders = cbAddReminders.Checked;
        }

        void cbAddAttendees_CheckedChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.AddAttendeesToDescription = cbAddAttendees.Checked;
        }

        void cbCreateFiles_CheckedChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.CreateTextFiles = cbCreateFiles.Checked;
        }

        private void cbStartWithWindows_CheckedChanged(object sender, EventArgs e)
        {
            OGSSettings.Instance.Autostart = cbStartWithWindows.Checked;
        }

        void NotifyIconDoubleClick(object sender, EventArgs e)
        {
            showToolStripMenuItem_Click(sender, e);
        }

        void MainFormResize(object sender, EventArgs e)
        {
            showToolStripMenuItem.Text = this.WindowState == FormWindowState.Normal ? "Hide" : "Show...";

            if (!cbMinimizeToTray.Checked)
                return;

            if (this.WindowState == FormWindowState.Minimized)
            {
                notifyIcon.Visible = true;
                // notifyIcon1.ShowBalloonTip(500, "OutlookGoogleSync", "Click to open again.", ToolTipIcon.Info);
                this.ShowInTaskbar = false;
                this.Hide();
            }
            else if (this.WindowState == FormWindowState.Normal)
            {
                this.ShowInTaskbar = true;
                // notifyIcon1.Visible = false;
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void showToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                this.WindowState = FormWindowState.Minimized;
            }
            else
            {
                this.Show();
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void buttonIconize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void sincNowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SyncNow_Click(sender, e);
        }
    }

    public static class StringExt
    {
        public static string Truncate(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }


    public static class DateTimeExt
    {
        public static DateTime RoundDown(this DateTime dateTime, int SyncPeriod)
        {
            return new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, dateTime.Hour, (dateTime.Minute / SyncPeriod) * SyncPeriod, 0);
        }
    }
}

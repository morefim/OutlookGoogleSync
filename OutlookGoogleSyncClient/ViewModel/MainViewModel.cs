using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GalaSoft.MvvmLight;
using OutlookGoogleSync;

namespace OutlookGoogleSyncClient.ViewModel
{
    /// <summary>
    /// This class contains properties that the main View can data bind to.
    /// <para>
    /// Use the <strong>mvvminpc</strong> snippet to add bindable properties to this ViewModel.
    /// </para>
    /// <para>
    /// You can also use Blend to data bind with the tool's support.
    /// </para>
    /// <para>
    /// See http://www.galasoft.ch/mvvm
    /// </para>
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        readonly SyncManager _syncManager = new SyncManager();

        public List<string> Calendars { get; set; }
        public string SelectedCalendar { get; set; }

        /// <summary>
        /// Initializes a new instance of the MainViewModel class.
        /// </summary>
        public MainViewModel()
        {
            if (IsInDesignMode)
            {
                // Code runs in Blend --> create design time data.
            }
            else
            {
                // Code runs "for real"
                //Calendars = new[] { "January", "February" };

                //SelectedCalendar = Calendars.First();

                Calendars = new List<string>();

                _syncManager.OnLogboxOutDelegate += Logboxout;
                _syncManager.OnExceptionDelegate += HandleException;
                _syncManager.OnSyncDoneDelegate += SyncDone;

                Task.Run((Func<Task>)_syncManager.DoWork);
            }
        }

        private void SyncDone(int deleted, int created)
        {
            throw new NotImplementedException();
        }

        private void HandleException(Exception e)
        {
            throw new NotImplementedException();
        }

        private void Logboxout(string text)
        {
            Calendars.Add(text);
        }
    }
}
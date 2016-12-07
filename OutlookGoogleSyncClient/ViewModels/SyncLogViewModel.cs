using Caliburn.Micro;

namespace OutlookGoogleSyncClient.ViewModels
{
    public sealed class SyncLogViewModel : Screen
    {
        private string _log = "There will be log text.....";

        public SyncLogViewModel()
        {
            DisplayName = "Syncronization";

            Log = "Test";
        }

        public string Log
        {
            get { return _log; }
            set
            {
                if (value == _log) return;
                _log = value;
                NotifyOfPropertyChange(() => Log);
            }
        }
    }
}

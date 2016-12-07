using Caliburn.Micro;

namespace OutlookGoogleSyncClient.ViewModels
{
    public sealed class ShellViewModel : Screen
    {
        public SyncLogViewModel SyncLogViewModel { get; }
        public SettingsViewModel SettingsViewModel { get; }

        public ShellViewModel(SyncLogViewModel syncLogViewModel, SettingsViewModel settingsViewModel)
        {
            SyncLogViewModel = syncLogViewModel;
            SettingsViewModel = settingsViewModel;
            DisplayName = "Outlook Google Sync Client";
        }
    }
}

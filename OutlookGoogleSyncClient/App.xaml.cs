using System.Windows;

namespace OutlookGoogleSyncClient
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private AppBootstrapper _appBootstrapper;
        public App()
        {
            _appBootstrapper = new AppBootstrapper();
        }
    }
}

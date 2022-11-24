using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Identity.Client;
using OutlookGoogleSync.Properties;

namespace OutlookGoogleSync
{
	/// <summary>
	/// Class with program entry point.
	/// </summary>
	internal sealed class Program
	{
		/// <summary>
		/// Program entry point.
		/// </summary>
		[STAThread]
		private static async Task Main(string[] args)
		{
            if (!SingleInstance.Start())
            {
                SingleInstance.ShowFirstInstance();
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                //await CreateApplication();
                Application.Run(new MainForm());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            SingleInstance.Stop();
		}

        public static async Task CreateApplication(/*bool useWam, bool useBrokerPreview*/)
        {
            await MSGraph.NextEvents(null);
            
            var builder = PublicClientApplicationBuilder
                .Create(ClientId)
                .WithAuthority($"{Instance}{Tenant}")
                .WithDefaultRedirectUri();

            //Use of Broker Requires redirect URI "ms-appx-web://microsoft.aad.brokerplugin/{client_id}" in app registration
            //if (useWam && !useBrokerPreview)
            //{
            builder.WithWindowsBrokerOptions(new WindowsBrokerOptions());
            //}
            //else if (useWam && useBrokerPreview)
            //{
            //    //builder.WithBrokerPreview(true);
            //}
            PublicClientApp = builder.Build();
            TokenCacheHelper.EnableSerialization(PublicClientApp.UserTokenCache);

            var builder1 = ConfidentialClientApplicationBuilder
                .Create(MainSettings.Default.appId)
                .WithClientSecret(MainSettings.Default.clientSecret)
                .WithTenantId(MainSettings.Default.tenantId);
            ConfidentialClientApp = builder1.Build();
            TokenCacheHelper.EnableSerialization(ConfidentialClientApp.UserTokenCache);
        }

        // Below are the clientId (Application Id) of your app registration and the tenant information. 
        // You have to replace:
        // - the content of ClientID with the Application Id for your app registration
        // - The content of Tenant by the information about the accounts allowed to sign-in in your application:
        //   - For Work or School account in your org, use your tenant ID, or domain
        //   - for any Work or School accounts, use organizations
        //   - for any Work or School accounts, or Microsoft personal account, use common
        //   - for Microsoft Personal account, use consumers
        private static readonly string ClientId = MainSettings.Default.appId;
        //private static readonly string ClientId = "4a1aa1d5-c567-49d0-ad0b-cd957a47f842";

        // Note: Tenant is important for the quickstart.
        //private static readonly string Tenant = MainSettings.Default.tenantId;
        private static readonly string Tenant = "common";
        private static readonly string Instance = "https://login.microsoftonline.com/";

        public static IPublicClientApplication PublicClientApp { get; private set; }
        public static IConfidentialClientApplication ConfidentialClientApp { get; private set; }
    }
}

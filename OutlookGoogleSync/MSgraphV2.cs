using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public class MSgraphV2
    {
        //Set the API Endpoint to Graph 'me' endpoint. 
        // To change from Microsoft public cloud to a national cloud, use another value of graphAPIEndpoint.
        // Reference with Graph endpoints here: https://docs.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints
        string graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me";

        //Set the scope for API call to user.read
        string[] scopes = new string[] { "user.read" };

        /// <summary>
        /// Call AcquireToken - to acquire a token requiring user to sign-in
        /// </summary>
        public async Task<string> CallGraphButton_Click(int howToSignIn, object handle)
        {
            AuthenticationResult authResult = null;

            var app = Program.PublicClientApp;
            //var app = Program.ConfidentialClientApp;

            IAccount firstAccount;

            switch (howToSignIn)
            {
                // 0: Use account used to signed-in in Windows (WAM)
                case 0:
                    // WAM will always get an account in the cache. So if we want
                    // to have a chance to select the accounts interactively, we need to
                    // force the non-account
                    firstAccount = PublicClientApplication.OperatingSystemAccount;
                    break;

                //  1: Use one of the Accounts known by Windows(WAM)
                case 1:
                    // We force WAM to display the dialog with the accounts
                    firstAccount = null;
                    break;

                //  Use any account(Azure AD). It's not using WAM
                default:
                    var accounts = await app.GetAccountsAsync();
                    firstAccount = accounts.FirstOrDefault();
                    break;
            }

            try
            {
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent. 
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(firstAccount)
                        .WithParentActivityOrWindow(handle/*new WindowInteropHelper(this).Handle*/) // optional, used to center the browser on the window
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                    return $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                return $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
            }

            if (authResult != null)
            {
                var text = await GetHttpContentWithToken(graphAPIEndpoint, authResult.AccessToken);
                return DisplayBasicTokenInfo(authResult);
            }

            return "Unexpected";
        }

        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click()
        {
            /*var accounts = await App.PublicClientApp.GetAccountsAsync();
            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    this.ResultText.Text = "User has signed-out";
                    this.CallGraphButton.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                }
                catch (MsalException ex)
                {
                    ResultText.Text = $"Error signing-out user: {ex.Message}";
                }
            }*/
        }

        /// <summary>
        /// Display basic information contained in the token
        /// </summary>
        private string DisplayBasicTokenInfo(AuthenticationResult authResult)
        {
            var text = "";
            if (authResult != null)
            {
                text += $"Username: {authResult.Account.Username}" + Environment.NewLine;
                text += $"Token Expires: {authResult.ExpiresOn.ToLocalTime()}" + Environment.NewLine;
            }

            return text;
        }

        private void UseWam_Changed()
        {
            //SignOutButton_Click(sender, e);
            //App.CreateApplication(howToSignIn.SelectedIndex != 2, Convert.ToBoolean(UseBrokerPreview?.IsChecked)); // Not Azure AD accounts (that is use WAM accounts)
        }

        private void UseBrokerPreview_Changed()
        {
            //SignOutButton_Click(sender, e);
            //App.CreateApplication(howToSignIn.SelectedIndex != 2, Convert.ToBoolean(UseBrokerPreview?.IsChecked));
        }
    }
}


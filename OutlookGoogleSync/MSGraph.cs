using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using OutlookGoogleSync.Properties;

namespace OutlookGoogleSync
{
    class MSGraph
    {
        public static void Test()
        {
            MicrosoftGraphCalendarOAuth2AccessToken();

            // This example requires the Chilkat API to have been previously unlocked.
            // See Global Unlock Sample for sample code.

            var http = new Chilkat.Http();

            // Use your previously obtained access token as shown here:
            //    Get Microsoft Graph OAuth2 Access Token with Calendars.ReadWrite scope.

            var jsonToken = new Chilkat.JsonObject();
            var success = jsonToken.LoadFile("qa_data/tokens/msGraphCalendar.json");
            if (success == false)
            {
                Debug.WriteLine(jsonToken.LastErrorText);
                return;
            }

            http.AuthToken = jsonToken.StringOf("access_token");

            // Create a JSON body for the HTTP POST
            // {
            //   "name": "Work"
            // }
            var json = new Chilkat.JsonObject();
            json.UpdateString("name", "Work");

            // POST the JSON to https://graph.microsoft.com/v1.0/me/calendars
            var resp = http.PostJson3("https://graph.microsoft.com/v1.0/me/calendars", "application/json", json);
            if (http.LastMethodSuccess == false)
            {
                Debug.WriteLine(http.LastErrorText);
                return;
            }

            json.Load(resp.BodyStr);
            json.EmitCompact = false;

            if (resp.StatusCode != 201)
            {
                Debug.WriteLine(json.Emit());
                Debug.WriteLine("Failed, response status code = " + Convert.ToString(resp.StatusCode));

                return;
            }

            Debug.WriteLine(json.Emit());

            // A sample response:
            // {
            //   "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('admin%40chilkat.io')/calendars/$entity",
            //   "id": "AQMkAD...TgAAAA==",
            //   "name": "Work",
            //   "color": "auto",
            //   "changeKey": "5+vF7T...HjDcA==",
            //   "canShare": true,
            //   "canViewPrivateItems": true,
            //   "canEdit": true,
            //   "owner": {
            //     "name": "...",
            //     "address": "outlook_3A33...4CC15@outlook.com"
            //   }
            // }

            // Use this online tool to generate parsing code from sample JSON: 
            // Generate Parsing Code from JSON

            var odataContext = json.StringOf("\"@odata.context\"");
            var id = json.StringOf("id");
            var name = json.StringOf("name");
            var color = json.StringOf("color");
            var changeKey = json.StringOf("changeKey");
            var canShare = json.BoolOf("canShare");
            var canViewPrivateItems = json.BoolOf("canViewPrivateItems");
            var canEdit = json.BoolOf("canEdit");
            var ownerName = json.StringOf("owner.name");
            var ownerAddress = json.StringOf("owner.address");

            Debug.WriteLine("Success.");
        }

        static void MicrosoftGraphCalendarOAuth2AccessToken()
        {
            // This example requires the Chilkat API to have been previously unlocked.
            // See Global Unlock Sample for sample code.

            Chilkat.OAuth2 oauth2 = new Chilkat.OAuth2();
            bool success;

            // This should be the port in the localhost callback URL for your app.  
            // The callback URL would look like "http://localhost:3017/" if the port number is 3017.
            oauth2.ListenPort = 3017;

            oauth2.AuthorizationEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
            oauth2.TokenEndpoint = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

            // Replace these with actual values.
            oauth2.ClientId = "MICROSOFT-GRAPH-CLIENT-ID";
            // This is your app password:
            oauth2.ClientSecret = "MICROSOFT-GRAPH-CLIENT-SECRET";

            oauth2.CodeChallenge = false;
            // Provide a SPACE separated list of scopes.
            // See https://developer.microsoft.com/en-us/graph/docs/authorization/permission_scopes 

            // Important: To get a refresh token in the final response, you must include the "offline_access" scope
            oauth2.Scope = "openid profile offline_access user.readwrite calendars.readwrite files.readwrite";

            // Begin the OAuth2 three-legged flow.  This returns a URL that should be loaded in a browser.
            string url = oauth2.StartAuth();
            if (oauth2.LastMethodSuccess != true)
            {
                Debug.WriteLine(oauth2.LastErrorText);
                return;
            }

            // At this point, your application should load the URL in a browser.
            // For example, 
            // in C#: System.Diagnostics.Process.Start(url);
            // in Java: Desktop.getDesktop().browse(new URI(url));
            // in VBScript: Set wsh=WScript.CreateObject("WScript.Shell")
            //              wsh.Run url
            // in Xojo: ShowURL(url)  (see http://docs.xojo.com/index.php/ShowURL)
            // in Dataflex: Runprogram Background "c:\Program Files\Internet Explorer\iexplore.exe" sUrl        
            // The Microsoft account owner would interactively accept or deny the authorization request.

            // Add the code to load the url in a web browser here...
            // Add the code to load the url in a web browser here...
            // Add the code to load the url in a web browser here...

            // Now wait for the authorization.
            // We'll wait for a max of 30 seconds.
            int numMsWaited = 0;
            while ((numMsWaited < 30000) && (oauth2.AuthFlowState < 3))
            {
                oauth2.SleepMs(100);
                numMsWaited = numMsWaited + 100;
            }

            // If there was no response from the browser within 30 seconds, then 
            // the AuthFlowState will be equal to 1 or 2.
            // 1: Waiting for Redirect. The OAuth2 background thread is waiting to receive the redirect HTTP request from the browser.
            // 2: Waiting for Final Response. The OAuth2 background thread is waiting for the final access token response.
            // In that case, cancel the background task started in the call to StartAuth.
            if (oauth2.AuthFlowState < 3)
            {
                oauth2.Cancel();
                Debug.WriteLine("No response from the browser!");
                return;
            }

            // Check the AuthFlowState to see if authorization was granted, denied, or if some error occurred
            // The possible AuthFlowState values are:
            // 3: Completed with Success. The OAuth2 flow has completed, the background thread exited, and the successful JSON response is available in AccessTokenResponse property.
            // 4: Completed with Access Denied. The OAuth2 flow has completed, the background thread exited, and the error JSON is available in AccessTokenResponse property.
            // 5: Failed Prior to Completion. The OAuth2 flow failed to complete, the background thread exited, and the error information is available in the FailureInfo property.
            if (oauth2.AuthFlowState == 5)
            {
                Debug.WriteLine("OAuth2 failed to complete.");
                Debug.WriteLine(oauth2.FailureInfo);
                return;
            }

            if (oauth2.AuthFlowState == 4)
            {
                Debug.WriteLine("OAuth2 authorization was denied.");
                Debug.WriteLine(oauth2.AccessTokenResponse);
                return;
            }

            if (oauth2.AuthFlowState != 3)
            {
                Debug.WriteLine("Unexpected AuthFlowState:" + Convert.ToString(oauth2.AuthFlowState));
                return;
            }

            Debug.WriteLine("OAuth2 authorization granted!");
            Debug.WriteLine("Access Token = " + oauth2.AccessToken);

            // Get the full JSON response:
            Chilkat.JsonObject json = new Chilkat.JsonObject();
            json.Load(oauth2.AccessTokenResponse);
            json.EmitCompact = false;

            // The JSON response looks like this:

            // {
            //   "token_type": "Bearer",
            //   "scope": "openid profile User.ReadWrite Calendars.ReadWrite Files.ReadWrite User.Read",
            //   "expires_in": 3600,
            //   "ext_expires_in": 0,
            //   "access_token": "EwBAA8l6B...",
            //   "refresh_token": "MCRMdbe...",
            //   "id_token": "eyJ0eXA..."
            // }

            // If an "expires_on" member does not exist, then add the JSON member by
            // getting the current system date/time and adding the "expires_in" seconds.
            // This way we'll know when the token expires.
            if (json.HasMember("expires_on") != true)
            {
                Chilkat.CkDateTime dtExpire = new Chilkat.CkDateTime();
                dtExpire.SetFromCurrentSystemTime();
                dtExpire.AddSeconds(json.IntOf("expires_in"));
                json.AppendString("expires_on", dtExpire.GetAsUnixTimeStr(false));
            }

            Debug.WriteLine(json.Emit());

            // Save the JSON to a file for future requests.
            Chilkat.FileAccess fac = new Chilkat.FileAccess();
            fac.WriteEntireTextFile("qa_data/tokens/msGraphCalendar.json", json.Emit(), "utf-8", false);
        }

        public static async Task NextEvents(IAuthenticationProvider authProvider)
        {
            await Authenticate();

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("startdatetime", "2022-10-15T20:39:57.490Z"),
                new QueryOption("enddatetime", "2022-10-22T20:39:57.490Z")
            };

            var calendarView = await graphClient.Me.CalendarView
                .Request(queryOptions)
                .GetAsync();
        }

        public static async Task Authenticate()
        {
            // The Azure AD tenant ID or a verified domain (e.g. contoso.onmicrosoft.com) 
            var tenantId = MainSettings.Default.tenantId;//"{tenant-id-or-domain-name}";

            // The client ID of the app registered in Azure AD
            var clientId = MainSettings.Default.appId;//"{client-id}";

            // *Never* include client secrets in source code!
            var clientSecret = MainSettings.Default.clientSecret;//await GetClientSecretFromKeyVault(); // Or some other secure place.

            var authority = $@"https://login.microsoftonline.com/${tenantId}/v2.0";

            // Configure the MSAL client as a confidential client
            var confidentialClient = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithAuthority(authority)
                .WithClientSecret(clientSecret)
                .Build();

            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 
            GraphServiceClient graphServiceClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                async requestMessage => await AuthenticateRequestAsyncDelegate(requestMessage, confidentialClient)));

            // Make a Microsoft Graph API query
            var users = await graphServiceClient.Users.Request().GetAsync();
        }

        private static async Task AuthenticateRequestAsyncDelegate(HttpRequestMessage requestMessage, IConfidentialClientApplication confidentialClient)
        {
            // The app registration should be configured to require access to permissions
            // sufficient for the Microsoft Graph API calls the app will be making, and
            // those permissions should be granted by a tenant administrator.
            var scopes = new string[] { "https://graph.microsoft.com/.default" };

            // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
            var authResult = await confidentialClient
                    .AcquireTokenForClient(scopes)
                    .ExecuteAsync();

            // Add the access token in the Authorization header of the API request.
            requestMessage.Headers.Authorization =
                new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DotNetOpenAuth.OAuth2;
using Google;
using Google.Apis.Authentication.OAuth2;
using Google.Apis.Authentication.OAuth2.DotNetOpenAuth;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Util;


namespace OutlookGoogleSync
{
	/// <summary>
	/// Description of GoogleCalendar.
	/// </summary>
	public class GoogleCalendar
	{
	    private static GoogleCalendar instance;

        public static GoogleCalendar Instance
        {
            get 
            {
                if (instance == null) instance = new GoogleCalendar();
                return instance;
            }
        }
        
	    CalendarService service;
	    
		public GoogleCalendar()
		{
            var provider = new NativeApplicationClient(GoogleAuthenticationServer.Description);
            provider.ClientIdentifier = "662204240419.apps.googleusercontent.com";
            provider.ClientSecret = "4nJPnk5fE8yJM_HNUNQEEvjU";
            service = new CalendarService(new OAuth2Authenticator<NativeApplicationClient>(provider, GetAuthentication));
            service.Key = "AIzaSyDRGFSAyMGondZKR8fww1RtRARYtCbBC4k";
		}
				
		private static IAuthorizationState GetAuthentication(NativeApplicationClient arg)
        {
            // Get the auth URL:
            IAuthorizationState state = new AuthorizationState(new[] { CalendarService.Scopes.Calendar.GetStringValue() });
            state.Callback = new Uri(NativeApplicationClient.OutOfBandCallbackUrl);
            state.RefreshToken = Coder.Decrypt(OGSSettings.Instance.RefreshToken);
            var authUri = arg.RequestUserAuthorization(state);
            
            IAuthorizationState result = null;
            
		    if (state.RefreshToken == "")
		    {
                // Request authorization from the user (by opening a browser window):
                Process.Start(authUri.ToString());
                
                var eac = new EnterAuthorizationCode();
                if (eac.ShowDialog(MainForm.Instance) == DialogResult.OK)
                {
                    // Retrieve the access/refresh tokens by using the authorization code:
                    result = arg.ProcessUserAuthorization(eac.authcode, state);
                    
                    //save the refresh token for future use
                    OGSSettings.Instance.RefreshToken = Coder.Encrypt(result.RefreshToken);
                    XMLManager.export(OGSSettings.Instance, MainForm.Filename);
                    
                    return result;
                } 
                else 
                {
                    return null;
                }		        
		    } 
            else 
            {
		        arg.RefreshToken(state, null);
		        result = state;
		        return result;
		    }
        
        }

        public List<OGSCalendarListEntry> getCalendars()
        {
            var request = service.CalendarList.List().Fetch();            
            if (request != null)
            {

                var result = new List<OGSCalendarListEntry>();
                foreach (var cle in request.Items)
                {
                    result.Add(new OGSCalendarListEntry(cle));
                }
                return result;
            }
            return null;
        }
		
        public List<Event> getCalendarEntriesInRange()
        {
            var result = new List<Event>();
            var lr = service.Events.List(OGSSettings.Instance.UseGoogleCalendar.CalendarID);

            lr.TimeMin = GoogleTimeFrom(DateTime.Now.AddDays(-OGSSettings.Instance.DaysInThePast));
            lr.TimeMax = GoogleTimeFrom(DateTime.Now.AddDays(+OGSSettings.Instance.DaysInTheFuture + 1));

            var request = lr.Fetch();
            if (request != null && request.Items != null)
            {
                result.AddRange(request.Items);
            }
            return result;
        }

        public void deleteCalendarEntry(Event e)
        {
            var request = service.Events.Delete(OGSSettings.Instance.UseGoogleCalendar.CalendarID, e.Id).Fetch();
        }

        public void addEntry(Event e)
        {
            try
            {
                var result = service.Events.Insert(e, OGSSettings.Instance.UseGoogleCalendar.CalendarID).Fetch();
            }
            catch (GoogleApiRequestException ex)
            {
                const string error403Signature = @"Calendar usage limits exceeded.";
                if (ex.RequestError.Code == 403 && ex.RequestError.Message == error403Signature)
                    MoveAttendeesToDescriptionAndRetry(e);
                else
                    throw;
            }
        }

	    private void MoveAttendeesToDescriptionAndRetry(Event e)
	    {
            e.Description += "Attendees:\r\n";
            foreach (var attendee in e.Attendees)
                e.Description += attendee.DisplayName + "[" + attendee.Email + "]\r\n";

            e.Attendees.Clear();
	        e.Attendees = null;

            var result = service.Events.Insert(e, OGSSettings.Instance.UseGoogleCalendar.CalendarID).Fetch();
	    }

	    //returns the Google Time Format String of a given .Net DateTime value
		//Google Time Format = "2012-08-20T00:00:00+02:00"
		public string GoogleTimeFrom(DateTime dt)
		{
            var timezone = TimeZoneInfo.Local.GetUtcOffset(dt).ToString();
            if (timezone[0] != '-') timezone = '+' + timezone;
            timezone = timezone.Substring(0,6);
            
            var result = dt.GetDateTimeFormats('s')[0] + timezone;
            return result;
		}		
	}
}

﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Microsoft.Identity.Client.Core;
using Microsoft.Identity.Client.Extensibility;
using Microsoft.Identity.Client.Http;
using Microsoft.Identity.Client.Internal.Broker;
using Microsoft.Identity.Client.Utils;
#if SUPPORTS_SYSTEM_TEXT_JSON
using System.Text.Json;
using Microsoft.Identity.Client.Platforms.net6;
using JObject = System.Text.Json.Nodes.JsonObject;
using JsonProperty = System.Text.Json.Serialization.JsonPropertyNameAttribute;
#else
using Microsoft.Identity.Json;
#endif

namespace Microsoft.Identity.Client.OAuth2
{
    internal class TokenResponseClaim : OAuth2ResponseBaseClaim
    {
        public const string Code = "code";
        public const string TokenType = "token_type";
        public const string AccessToken = "access_token";
        public const string RefreshToken = "refresh_token";
        public const string IdToken = "id_token";
        public const string Scope = "scope";
        public const string ClientInfo = "client_info";
        public const string ExpiresIn = "expires_in";
        public const string CloudInstanceHost = "cloud_instance_host_name";
        public const string CreatedOn = "created_on";
        public const string ExtendedExpiresIn = "ext_expires_in";
        public const string Authority = "authority";
        public const string FamilyId = "foci";
        public const string RefreshIn = "refresh_in";
        public const string SpaCode = "spa_code";
        public const string ErrorSubcode = "error_subcode";
        public const string ErrorSubcodeCancel = "cancel";

        public const string TenantId = "tenant_id";
        public const string Upn = "username";
        public const string LocalAccountId = "local_account_id";
    }

    [JsonObject]
    [Preserve(AllMembers = true)]
    internal class MsalTokenResponse : OAuth2ResponseBase
    {
        public MsalTokenResponse()
        {

        }

        private const string iOSBrokerErrorMetadata = "error_metadata";
        private const string iOSBrokerHomeAccountId = "home_account_id";
        [JsonProperty(TokenResponseClaim.TokenType)]
        public string TokenType { get; set; }

        [JsonProperty(TokenResponseClaim.AccessToken)]
        public string AccessToken { get; set; }

        [JsonProperty(TokenResponseClaim.RefreshToken)]
        public string RefreshToken { get; set; }

        [JsonProperty(TokenResponseClaim.Scope)]
        public string Scope { get; set; }

        [JsonProperty(TokenResponseClaim.ClientInfo)]
        public string ClientInfo { get; set; }

        [JsonProperty(TokenResponseClaim.IdToken)]
        public string IdToken { get; set; }

        [JsonProperty(TokenResponseClaim.ExpiresIn)]
        public long ExpiresIn { get; set; }

        [JsonProperty(TokenResponseClaim.ExtendedExpiresIn)]
        public long ExtendedExpiresIn { get; set; }

        [JsonProperty(TokenResponseClaim.RefreshIn)]
        public long? RefreshIn { get; set; }

        /// <summary>
        /// Optional field, FOCI support.
        /// </summary>
        [JsonProperty(TokenResponseClaim.FamilyId)]
        public string FamilyId { get; set; }

        [JsonProperty(TokenResponseClaim.SpaCode)]
        public string SpaAuthCode { get; set; }

        [JsonProperty(TokenResponseClaim.Authority)]
        public string AuthorityUrl { get; set; }

        public string TenantId { get; set; }

        public string Upn { get; set; }

        public string AccountUserId { get; set; }

        public string WamAccountId { get; set; }

        public TokenSource TokenSource { get; set; }

        public HttpResponse HttpResponse { get; set; }

        internal static MsalTokenResponse CreateFromiOSBrokerResponse(Dictionary<string, string> responseDictionary)
        {
            if (responseDictionary.TryGetValue(BrokerResponseConst.BrokerErrorCode, out string errorCode))
            {
                string metadataOriginal = responseDictionary.ContainsKey(MsalTokenResponse.iOSBrokerErrorMetadata) ? responseDictionary[MsalTokenResponse.iOSBrokerErrorMetadata] : null;
                Dictionary<string, string> metadataDictionary = null;

                if (metadataOriginal != null)
                {
                    string brokerMetadataJson = Uri.UnescapeDataString(metadataOriginal);
#if SUPPORTS_SYSTEM_TEXT_JSON
                    metadataDictionary = new Dictionary<string, string>();
                    foreach (var item in JsonDocument.Parse(brokerMetadataJson).RootElement.EnumerateObject())
                    {
                        metadataDictionary.Add(item.Name, item.Value.GetString());
                    }
#else
                    metadataDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(brokerMetadataJson);
#endif
                }

                string homeAcctId = null;
                metadataDictionary?.TryGetValue(MsalTokenResponse.iOSBrokerHomeAccountId, out homeAcctId);
                return new MsalTokenResponse
                {
                    Error = errorCode,
                    ErrorDescription = responseDictionary.ContainsKey(BrokerResponseConst.BrokerErrorDescription) ? CoreHelpers.UrlDecode(responseDictionary[BrokerResponseConst.BrokerErrorDescription]) : string.Empty,
                    SubError = responseDictionary.ContainsKey(OAuth2ResponseBaseClaim.SubError) ? responseDictionary[OAuth2ResponseBaseClaim.SubError] : string.Empty,
                    AccountUserId = homeAcctId != null ? AccountId.ParseFromString(homeAcctId).ObjectId : null,
                    TenantId = homeAcctId != null ? AccountId.ParseFromString(homeAcctId).TenantId : null,
                    Upn = (metadataDictionary?.ContainsKey(TokenResponseClaim.Upn) ?? false) ? metadataDictionary[TokenResponseClaim.Upn] : null,
                    CorrelationId = responseDictionary.ContainsKey(BrokerResponseConst.CorrelationId) ? responseDictionary[BrokerResponseConst.CorrelationId] : null,
                };
            }

            var response = new MsalTokenResponse
            {
                AccessToken = responseDictionary[BrokerResponseConst.AccessToken],
                RefreshToken = responseDictionary.ContainsKey(BrokerResponseConst.RefreshToken)
                    ? responseDictionary[BrokerResponseConst.RefreshToken]
                    : null,
                IdToken = responseDictionary[BrokerResponseConst.IdToken],
                TokenType = BrokerResponseConst.Bearer,
                CorrelationId = responseDictionary[BrokerResponseConst.CorrelationId],
                Scope = responseDictionary[BrokerResponseConst.Scope],
                ExpiresIn = responseDictionary.TryGetValue(BrokerResponseConst.ExpiresOn, out string expiresOn) ?
                                DateTimeHelpers.GetDurationFromNowInSeconds(expiresOn) :
                                0,
                ClientInfo = responseDictionary.ContainsKey(BrokerResponseConst.ClientInfo)
                                ? responseDictionary[BrokerResponseConst.ClientInfo]
                                : null,
                TokenSource = TokenSource.Broker
            };

            if (responseDictionary.ContainsKey(TokenResponseClaim.RefreshIn))
            {
                response.RefreshIn = long.Parse(
                    responseDictionary[TokenResponseClaim.RefreshIn],
                    CultureInfo.InvariantCulture);
            }

            return response;
        }

        internal static MsalTokenResponse CreateFromAppProviderResponse(AppTokenProviderResult tokenProviderResponse)
        {
            ValidateTokenProviderResult(tokenProviderResponse);

            var response = new MsalTokenResponse
            {
                AccessToken = tokenProviderResponse.AccessToken,
                RefreshToken = null,
                IdToken = null,
                TokenType = BrokerResponseConst.Bearer,
                ExpiresIn = tokenProviderResponse.ExpiresInSeconds,
                ClientInfo = null,
                TokenSource = TokenSource.IdentityProvider,
                TenantId = null //Leaving as null so MSAL can use the original request Tid. This is ok for confidential client scenarios
            };

            response.RefreshIn = tokenProviderResponse.RefreshInSeconds;

            return response;
        }

        private static void ValidateTokenProviderResult(AppTokenProviderResult TokenProviderResult)
        {
            if (string.IsNullOrEmpty(TokenProviderResult.AccessToken))
            {
                HandleInvalidExternalValueError(nameof(TokenProviderResult.AccessToken));
            }

            if (TokenProviderResult.ExpiresInSeconds == 0 || TokenProviderResult.ExpiresInSeconds < 0)
            {
                HandleInvalidExternalValueError(nameof(TokenProviderResult.ExpiresInSeconds));
            }
        }

        private static void HandleInvalidExternalValueError(string nameOfValue)
        {
            throw new MsalClientException(MsalError.InvalidTokenProviderResponseValue, MsalErrorMessage.InvalidTokenProviderResponseValue(nameOfValue));
        }

        /// <remarks>
        /// This method does not belong here - it is more tied to the Android code. However, that code is
        /// not unit testable, and this one is. 
        /// The values of the JSON response are based on 
        /// https://github.com/AzureAD/microsoft-authentication-library-common-for-android/blob/dev/common/src/main/java/com/microsoft/identity/common/internal/broker/BrokerResult.java
        /// </remarks>
        internal static MsalTokenResponse CreateFromAndroidBrokerResponse(string jsonResponse, string correlationId)
        {
            var authResult = JsonHelper.ParseIntoJsonObject(jsonResponse);
            var errorCode = authResult[BrokerResponseConst.BrokerErrorCode]?.ToString();

            if (!string.IsNullOrEmpty(errorCode))
            {
                return new MsalTokenResponse
                {
                    Error = errorCode,
                    ErrorDescription = authResult[BrokerResponseConst.BrokerErrorMessage]?.ToString(),
                    AuthorityUrl = authResult[BrokerResponseConst.Authority]?.ToString(),
                    TenantId = authResult[BrokerResponseConst.TenantId]?.ToString(),
                    Upn = authResult[BrokerResponseConst.UserName]?.ToString(),
                    AccountUserId = authResult[BrokerResponseConst.LocalAccountId]?.ToString(),
                };
            }

            MsalTokenResponse msalTokenResponse = new MsalTokenResponse()
            {
                AccessToken = authResult[BrokerResponseConst.AccessToken].ToString(),
                IdToken = authResult[BrokerResponseConst.IdToken].ToString(),
                CorrelationId = correlationId, // Android response does not expose Correlation ID
                Scope = authResult[BrokerResponseConst.AndroidScopes].ToString(), // sadly for iOS this is "scope" and for Android "scopes"
                ExpiresIn = DateTimeHelpers.GetDurationFromNowInSeconds(authResult[BrokerResponseConst.ExpiresOn].ToString()),
                ExtendedExpiresIn = DateTimeHelpers.GetDurationFromNowInSeconds(authResult[BrokerResponseConst.ExtendedExpiresOn].ToString()),
                ClientInfo = authResult[BrokerResponseConst.ClientInfo].ToString(),
                TokenType = authResult[BrokerResponseConst.TokenType]?.ToString() ?? "Bearer",
                TokenSource = TokenSource.Broker,
                AuthorityUrl = authResult[BrokerResponseConst.Authority]?.ToString(),
                TenantId = authResult[BrokerResponseConst.TenantId]?.ToString(),
                Upn = authResult[BrokerResponseConst.UserName]?.ToString(),
                AccountUserId = authResult[BrokerResponseConst.LocalAccountId]?.ToString(),
            };

            return msalTokenResponse;
        }

        public void Log(ILoggerAdapter logger, LogLevel logLevel)
        {
            if (logger.IsLoggingEnabled(logLevel))
            {
                StringBuilder withPii = new StringBuilder();
                StringBuilder withoutPii = new StringBuilder();

                withPii.AppendLine($"{Environment.NewLine}[MsalTokenResponse]");
                withPii.AppendLine($"Error: {Error}");
                withPii.AppendLine($"ErrorDescription: {ErrorDescription}");
                withPii.AppendLine($"Scopes: {Scope} ");
                withPii.AppendLine($"ExpiresIn: {ExpiresIn}");
                withPii.AppendLine($"RefreshIn: {RefreshIn}");
                withPii.AppendLine($"AccessToken returned: {!string.IsNullOrEmpty(AccessToken)}");
                withPii.AppendLine($"AccessToken Type: {TokenType}");
                withPii.AppendLine($"RefreshToken returned: {!string.IsNullOrEmpty(RefreshToken)}");
                withPii.AppendLine($"IdToken returned: {!string.IsNullOrEmpty(IdToken)}");
                withPii.AppendLine($"ClientInfo: {ClientInfo}");
                withPii.AppendLine($"FamilyId: {FamilyId}");
                withPii.AppendLine($"WamAccountId exists: {!string.IsNullOrEmpty(WamAccountId)}");

                withoutPii.AppendLine($"{Environment.NewLine}[MsalTokenResponse]");
                withoutPii.AppendLine($"Error: {Error}");
                withoutPii.AppendLine($"ErrorDescription: {ErrorDescription}");
                withoutPii.AppendLine($"Scopes: {Scope} ");
                withoutPii.AppendLine($"ExpiresIn: {ExpiresIn}");
                withoutPii.AppendLine($"RefreshIn: {RefreshIn}");
                withoutPii.AppendLine($"AccessToken returned: {!string.IsNullOrEmpty(AccessToken)}");
                withoutPii.AppendLine($"AccessToken Type: {TokenType}");
                withoutPii.AppendLine($"RefreshToken returned: {!string.IsNullOrEmpty(RefreshToken)}");
                withoutPii.AppendLine($"IdToken returned: {!string.IsNullOrEmpty(IdToken)}");
                withoutPii.AppendLine($"ClientInfo returned: {!string.IsNullOrEmpty(ClientInfo)}");
                withoutPii.AppendLine($"FamilyId: {FamilyId}");
                withoutPii.AppendLine($"WamAccountId exists: {!string.IsNullOrEmpty(WamAccountId)}");

                logger.Log(logLevel, withPii.ToString(), withoutPii.ToString());
            }
        }
    }
}

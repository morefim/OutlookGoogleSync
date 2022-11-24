﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Identity.Client.Cache.Keys;
using Microsoft.Identity.Client.OAuth2;
using Microsoft.Identity.Client.Utils;
#if SUPPORTS_SYSTEM_TEXT_JSON
using JObject = System.Text.Json.Nodes.JsonObject;
#else
using Microsoft.Identity.Json.Linq;
#endif

namespace Microsoft.Identity.Client.Cache.Items
{
    internal class MsalRefreshTokenCacheItem : MsalCredentialCacheItemBase
    {
        internal MsalRefreshTokenCacheItem()
        {
            CredentialType = StorageJsonValues.CredentialTypeRefreshToken;
        }

        internal MsalRefreshTokenCacheItem(
            string preferredCacheEnv,
            string clientId,
            MsalTokenResponse response,
            string homeAccountId)
            : this(
                  preferredCacheEnv,
                  clientId,
                  response.RefreshToken,
                  response.ClientInfo,
                  response.FamilyId,
                  homeAccountId)
        {
        }

        internal MsalRefreshTokenCacheItem(
            string preferredCacheEnv,
            string clientId,
            string secret,
            string rawClientInfo,
            string familyId,
            string homeAccountId)
            : this()
        {
            ClientId = clientId;
            Environment = preferredCacheEnv;
            Secret = secret;
            RawClientInfo = rawClientInfo;
            FamilyId = familyId;
            HomeAccountId = homeAccountId;
        }

        /// <summary>
        /// Optional. A value here means the token in an FRT.
        /// </summary>
        public string FamilyId { get; set; }

        /// <summary>
        /// Used to find the token in the cache.
        /// Can be a token assertion hash (normal OBO flow) or a user provided key (long-running OBO flow).
        /// </summary>
        internal string OboCacheKey { get; set; }

        /// <summary>
        /// Family Refresh Tokens, can be used for all clients part of the family
        /// </summary>
        public bool IsFRT => !string.IsNullOrEmpty(FamilyId);

        internal MsalRefreshTokenCacheKey GetKey()
        {
            return new MsalRefreshTokenCacheKey(Environment, ClientId, HomeAccountId, FamilyId);
        }

        internal static MsalRefreshTokenCacheItem FromJsonString(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }

            return FromJObject(JsonHelper.ParseIntoJsonObject(json));
        }

        internal static MsalRefreshTokenCacheItem FromJObject(JObject j)
        {
            var item = new MsalRefreshTokenCacheItem();
            item.FamilyId = JsonHelper.ExtractExistingOrEmptyString(j, StorageJsonKeys.FamilyId);
            item.OboCacheKey = JsonHelper.ExtractExistingOrEmptyString(j, StorageJsonKeys.UserAssertionHash);

            item.PopulateFieldsFromJObject(j);

            return item;
        }

        internal override JObject ToJObject()
        {
            var json = base.ToJObject();
            SetItemIfValueNotNull(json, StorageJsonKeys.FamilyId, FamilyId);
            SetItemIfValueNotNull(json, StorageJsonKeys.UserAssertionHash, OboCacheKey);
            return json;
        }

        internal string ToJsonString()
        {
            return ToJObject().ToString();
        }
    }
}

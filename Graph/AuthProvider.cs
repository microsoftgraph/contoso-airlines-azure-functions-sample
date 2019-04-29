// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
using CreateFlightTeam.DocumentDB;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace CreateFlightTeam.Graph
{
    public static class AuthProvider
    {
        private static readonly string appId = Environment.GetEnvironmentVariable("AppId");
        private static readonly string tid = Environment.GetEnvironmentVariable("TenantId");
        private static readonly string authority = $"https://login.microsoftonline.com/{tid}";
        private static readonly ClientCredential clientCreds = new ClientCredential(
            Environment.GetEnvironmentVariable("AppSecret"));
        private static readonly string redirectUri = Environment.GetEnvironmentVariable("RedirectUri");
        private static readonly string[] scopes = { "https://graph.microsoft.com/.default" };
        private static readonly string notifAppId = Environment.GetEnvironmentVariable("NotificationAppId");
        private static readonly ClientCredential notifClientCreds = new ClientCredential(
            Environment.GetEnvironmentVariable("NotificationAppSecret"));
        private static readonly string[] notifScopes = { "Notifications.ReadWrite.CreatedByApp" };

        private static ILogger logger;

        public static ILogger AzureLogger
        {
            get { return logger; }
            set { logger = value; }
        }

        public static async Task GetTokenOnBehalfOfAsync(string authHeader, ILogger log)
        {
            logger = log;
            if (string.IsNullOrEmpty(authHeader))
            {
                throw new MsalException("missing_auth", "Authorization header is not present on request.");
            }

            // Parse the auth header
            var parsedHeader = AuthenticationHeaderValue.Parse(authHeader);

            if (parsedHeader.Scheme.ToLower() != "bearer")
            {
                throw new MsalException("invalid_scheme", "Authorization header is missing the 'bearer' scheme.");
            }

            var confidentialClient = new ConfidentialClientApplication(notifAppId,
                authority, redirectUri, notifClientCreds, BlobTokenCache.GetMsalCacheInstance(), null);

            //Logger.LogCallback = AuthLog;
            //Logger.Level = Microsoft.Identity.Client.LogLevel.Verbose;
            //Logger.PiiLoggingEnabled = true;
            var userAssertion = new UserAssertion(parsedHeader.Parameter);

            try
            {
                var result = await confidentialClient.AcquireTokenOnBehalfOfAsync(notifScopes, userAssertion);
            }
            catch (Exception ex)
            {
                logger.LogError($"Error getting OBO token: {ex.Message}");
                throw ex;
            }
        }

        public static async Task<string> GetUserToken(string userId)
        {
            var confidentialClient = new ConfidentialClientApplication(notifAppId,
                authority, redirectUri, notifClientCreds, BlobTokenCache.GetMsalCacheInstance(), null);

            //Logger.LogCallback = AuthLog;
            //Logger.Level = Microsoft.Identity.Client.LogLevel.Verbose;
            //Logger.PiiLoggingEnabled = true;

            var account = await confidentialClient.GetAccountAsync($"{userId}.{tid}");

            if (account == null)
            {
                return string.Empty;
            }

            try
            {
                var result = await confidentialClient.AcquireTokenSilentAsync(notifScopes, account);
                return result.AccessToken;
            }
            catch (MsalException)
            {
                return string.Empty;
            }
        }

        private static void AuthLog(Microsoft.Identity.Client.LogLevel level, string message, bool containsPII)
        {
            logger.LogInformation($"MSAL: {message}");
        }
    }
}

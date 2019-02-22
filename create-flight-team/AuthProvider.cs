// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
//using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Identity.Client;
using System;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace create_flight_team
{
    public static class AuthProvider
    {
        private static readonly string appId = Environment.GetEnvironmentVariable("AppId");
        private static readonly string tid = Environment.GetEnvironmentVariable("TenantId");
        private static readonly string authority = $"https://login.microsoftonline.com/{tid}";
        private static readonly ClientCredential clientCreds = new ClientCredential(
            Environment.GetEnvironmentVariable("AppSecret"));
        private static readonly string[] scopes = { "https://graph.microsoft.com/.default" };

        public static async Task<string> GetTokenOnBehalfOfAsync(string authHeader)
        {
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

            // Create an assertion based on the provided token
            var userAssertion = new UserAssertion(parsedHeader.Parameter);

            var confidentialClient = new ConfidentialClientApplication(appId, authority, clientCreds, null, null);

            // Exchange the provided token for a Graph token
            var result = await confidentialClient.AcquireTokenOnBehalfOfAsync(scopes, userAssertion, authority);

            return result.AccessToken;
        }
    }
}

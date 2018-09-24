// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace create_flight_team
{
    public static class AuthProvider
    {
        private static readonly string tid = Environment.GetEnvironmentVariable("TenantId");
        private static readonly string authority = $"https://login.microsoftonline.com/{tid}";
        private static readonly ClientCredential clientCreds = new ClientCredential(
            Environment.GetEnvironmentVariable("AppId"),
            Environment.GetEnvironmentVariable("AppSecret"));

        public static async Task<string> GetTokenOnBehalfOfAsync(string authHeader)
        {
            if (string.IsNullOrEmpty(authHeader))
            {
                throw new AdalException("missing_auth", "Authorization header is not present on request.");
            }

            // Parse the auth header
            var parsedHeader = AuthenticationHeaderValue.Parse(authHeader);
            if (parsedHeader.Scheme.ToLower() != "bearer")
            {
                throw new AdalException("invalid_scheme", "Authorization header is missing the 'bearer' scheme.");
            }

            // Create an assertion based on the provided token
            var userAssertion = new UserAssertion(parsedHeader.Parameter, "urn:ietf:params:oauth:grant-type:jwt-bearer");
            var authContext = new AuthenticationContext(authority);

            // Exchange the provided token for a Graph token
            var result = await authContext.AcquireTokenAsync("https://graph.microsoft.com",
                clientCreds, userAssertion);

            return result.AccessToken;
        }
    }
}

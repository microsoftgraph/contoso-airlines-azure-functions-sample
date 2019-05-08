using Microsoft.Identity.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;
using System;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CreateFlightTeam.Graph
{
    public static class BlobTokenCache
    {
        private static readonly string connectionString = Environment.GetEnvironmentVariable("AzureWebJobsStorage");
        private static string tokenBlobETag = string.Empty;

        private static TokenCache cache = new TokenCache();

        public static TokenCache GetMsalCacheInstance()
        {
            cache.SetBeforeAccess(BeforeAccessNotification);
            cache.SetAfterAccess(AfterAccessNotification);
            Load().Wait();
            return cache;
        }

        private static async Task<CloudBlockBlob> GetTokenStorageBlob()
        {
            try
            {
                var storageAccount = CloudStorageAccount.Parse(connectionString);
                var client = storageAccount.CreateCloudBlobClient();
                var container = client.GetContainerReference("tokencache");

                await container.CreateIfNotExistsAsync();

                return container.GetBlockBlobReference("MsalUserTokenCache");
            }
            catch
            {
                return null;
            }
        }
        private static async Task Load()
        {
            var tokenBlob = await GetTokenStorageBlob();

            if (await tokenBlob.ExistsAsync())
            {
                await tokenBlob.FetchAttributesAsync();

                // Check if we need to reload from the blob
                if (tokenBlob.Properties.ETag.CompareTo(tokenBlobETag) != 0)
                {
                    var blobCacheBytes = new byte[tokenBlob.Properties.Length];
                    await tokenBlob.DownloadToByteArrayAsync(blobCacheBytes, 0);

                    cache.Deserialize(blobCacheBytes);
                    tokenBlobETag = tokenBlob.Properties.ETag;
                }
            }
        }

        private static async Task Persist()
        {
            // Reflect changes in the persistent store
            var cacheBytes = cache.Serialize();
            var tokenBlob = await GetTokenStorageBlob();

            await tokenBlob.UploadFromByteArrayAsync(cacheBytes, 0, cacheBytes.Length);
        }

        // Triggered right before MSAL needs to access the cache.
        // Reload the cache from the persistent store in case it changed since the last access.
        private static void BeforeAccessNotification(TokenCacheNotificationArgs args)
        {
            Load().Wait();
        }

        // Triggered right after MSAL accessed the cache.
        private static void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            // if the access operation resulted in a cache update
            if (args.HasStateChanged)
            {
                Persist().Wait();
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using CreateFlightTeam.Models;
using Microsoft.Azure.Documents;
using Microsoft.Azure.Documents.Client;
using Microsoft.Azure.Documents.Linq;

namespace CreateFlightTeam.DocumentDB
{
    public static class DatabaseHelper
    {
        private static readonly string databaseUri = Environment.GetEnvironmentVariable("DatabaseUri");
        private static readonly string databaseKey = Environment.GetEnvironmentVariable("DatabaseKey");
        private static readonly string databaseName = "FlightTeamProvisioning";
        private static readonly string flightCollection = "FlightTeams";
        private static readonly string subscriptionCollection = "Subscriptions";

        private static DocumentClient client = null;

        #region Initialization

        public static void Initialize()
        {
            if (client == null)
            {
                client = new DocumentClient(new Uri(databaseUri), databaseKey);
            }
            CreateDatabaseIfNotExistsAsync().Wait();
            CreateCollectionsIfNotExistsAsync().Wait();
        }

        private static async Task CreateDatabaseIfNotExistsAsync()
        {
            try
            {
                await client.ReadDatabaseAsync(UriFactory.CreateDatabaseUri(databaseName));
            }
            catch (DocumentClientException e)
            {
                if (e.StatusCode == HttpStatusCode.NotFound)
                {
                    await client.CreateDatabaseAsync(new Database { Id = databaseName });
                }
                else
                {
                    throw;
                }
            }
        }

        private static async Task CreateCollectionsIfNotExistsAsync()
        {
            try
            {
                await client.ReadDocumentCollectionAsync(UriFactory.CreateDocumentCollectionUri(databaseName, flightCollection));
            }
            catch (DocumentClientException e)
            {
                if (e.StatusCode == HttpStatusCode.NotFound)
                {
                    await client.CreateDocumentCollectionAsync(
                        UriFactory.CreateDatabaseUri(databaseName),
                        new DocumentCollection { Id = flightCollection },
                        new RequestOptions { OfferThroughput = 1000 });
                }
                else
                {
                    throw;
                }
            }

            try
            {
                await client.ReadDocumentCollectionAsync(UriFactory.CreateDocumentCollectionUri(databaseName, subscriptionCollection));
            }
            catch (DocumentClientException e)
            {
                if (e.StatusCode == HttpStatusCode.NotFound)
                {
                    await client.CreateDocumentCollectionAsync(
                        UriFactory.CreateDatabaseUri(databaseName),
                        new DocumentCollection { Id = subscriptionCollection },
                        new RequestOptions { OfferThroughput = 1000 });
                }
                else
                {
                    throw;
                }
            }
        }

        #endregion

        #region FlightTeam operations

        public static async Task<IEnumerable<FlightTeam>> GetFlightTeamsAsync()
        {
            return await GetItemsAsync<FlightTeam>(flightCollection);
        }

        public static async Task<IEnumerable<FlightTeam>> GetFlightTeamsAsync(Expression<Func<FlightTeam, bool>> predicate)
        {
            return await GetItemsAsync(predicate, flightCollection);
        }

        public static async Task<FlightTeam> GetFlightTeamAsync(string id)
        {
            return await GetItemAsync<FlightTeam>(id, flightCollection);
        }

        public static async Task<FlightTeam> CreateFlightTeamAsync(FlightTeam flightTeam)
        {
            return await CreateItemAsync(flightTeam, flightCollection);
        }

        public static async Task<FlightTeam> UpdateFlightTeamAsync(string id, FlightTeam flightTeam)
        {
            return await UpdateItemAsync(id, flightTeam, flightCollection);
        }

        public static async Task DeleteFlightTeamAsync(string id)
        {
            await DeleteItemAsync(id, flightCollection);
        }

        #endregion

        #region ListSubscription operations

        public static async Task<IEnumerable<ListSubscription>> GetListSubscriptionsAsync()
        {
            return await GetItemsAsync<ListSubscription>(subscriptionCollection);
        }

        public static async Task<IEnumerable<ListSubscription>> GetListSubscriptionsAsync(Expression<Func<ListSubscription, bool>> predicate)
        {
            return await GetItemsAsync(predicate, subscriptionCollection);
        }

        public static async Task<ListSubscription> GetListSubscriptionAsync(string id)
        {
            return await GetItemAsync<ListSubscription>(id, subscriptionCollection);
        }

        public static async Task<ListSubscription> CreateListSubscriptionAsync(ListSubscription subscription)
        {
            // Check if there is an existing record and don't
            // let the create happen if there is.
            var existingSubscriptions = await GetListSubscriptionsAsync(s => s.Resource.CompareTo(subscription.Resource) == 0);
            if (existingSubscriptions.Count() > 0)
            {
                throw new InvalidOperationException("A subscription record already exists.");
            }
            return await CreateItemAsync(subscription, subscriptionCollection);
        }

        public static async Task<ListSubscription> UpdateListSubscriptionAsync(string id, ListSubscription subscription)
        {
            return await UpdateItemAsync(id, subscription, subscriptionCollection);
        }

        public static async Task DeleteListSubscriptionAsync(string id)
        {
            await DeleteItemAsync(id, subscriptionCollection);
        }

        #endregion

        #region Generic operations

        private static async Task<IEnumerable<T>> GetItemsAsync<T>(string collection)
        {
            IDocumentQuery<T> query = client.CreateDocumentQuery<T>(
                UriFactory.CreateDocumentCollectionUri(databaseName, collection))
                .AsDocumentQuery();

            var results = new List<T>();
            while (query.HasMoreResults)
            {
                results.AddRange(await query.ExecuteNextAsync<T>());
            }

            return results;
        }

        private static async Task<IEnumerable<T>> GetItemsAsync<T>(Expression<Func<T, bool>> predicate, string collection)
        {
            IDocumentQuery<T> query = client.CreateDocumentQuery<T>(
                UriFactory.CreateDocumentCollectionUri(databaseName, collection))
                .Where(predicate)
                .AsDocumentQuery();

            var results = new List<T>();
            while (query.HasMoreResults)
            {
                results.AddRange(await query.ExecuteNextAsync<T>());
            }

            return results;
        }

        private static async Task<T> GetItemAsync<T>(string id, string collection)
        {
            try
            {
                Document document = await client.ReadDocumentAsync(UriFactory.CreateDocumentUri(databaseName, collection, id));
                return (T)(dynamic)document;
            }
            catch (DocumentClientException e)
            {
                if (HttpStatusCode.NotFound == e.StatusCode)
                {
                    return default(T);
                }
                else
                {
                    throw;
                }
            }
        }

        private static async Task<T> CreateItemAsync<T>(T item, string collection)
        {
            Document document = await client.CreateDocumentAsync(
                UriFactory.CreateDocumentCollectionUri(databaseName, collection),
                item);

            return (T)(dynamic)document;
        }

        private static async Task<T> UpdateItemAsync<T>(string id, T item, string collection)
        {
            Document document = await client.ReplaceDocumentAsync(
                UriFactory.CreateDocumentUri(databaseName, collection, id), item);

            return (T)(dynamic)document;
        }

        private static async Task DeleteItemAsync(string id, string collection)
        {
            try
            {
                await client.DeleteDocumentAsync(UriFactory.CreateDocumentUri(databaseName, collection, id));
            }
            catch (DocumentClientException e)
            {
                if (e.StatusCode != HttpStatusCode.NotFound)
                {
                    throw;
                }
            }
        }
        #endregion
    }
}

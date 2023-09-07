using Azure.Core;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Net.Http.Headers;
using System.Net.Sockets;

namespace appsvc_fnc_dev_gcxonboardsync_dotnet001
{
    internal class Auth
    {


        public GraphServiceClient graphAuth(ILogger log)
        {
            IConfiguration config = new ConfigurationBuilder()

           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            log.LogInformation("C# HTTP trigger function processed a request.");
            var scopes = new string[] { "https://graph.microsoft.com/.default" };
            var keyVaultUrl = config["keyVaultUrl"];
            var clientId = config["clientId"];
            var tenantid = config["tenantid"];
            var secretName = config["secretName"];

            //SecretClientOptions options = new SecretClientOptions()
            //{
            //    Retry =
            //{
            //    Delay= TimeSpan.FromSeconds(2),
            //    MaxDelay = TimeSpan.FromSeconds(16),
            //    MaxRetries = 5,
            //    Mode = RetryMode.Exponential
            // }
            //};
            //var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential(), options);

            //KeyVaultSecret secret = client.GetSecret(secretName);

            //string secretValue = secret.Value;

            string secretValue = "ngJ8Q~v9YUTU7ycgJhLVtRNRss0-hJD5YkrKccAi";

            // using Azure.Identity;
            var tokenOptions = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(tenantid, clientId, secretValue, tokenOptions);

            try
            {
                var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
                return graphClient;
            }
            catch (Exception e)
            {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");
                return null;
            }
        }



    }
}

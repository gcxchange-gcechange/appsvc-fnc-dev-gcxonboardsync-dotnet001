using System;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using System.Threading.Tasks;
using Microsoft.Kiota.Abstractions;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;
using Microsoft.WindowsAzure.Storage;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware.Options;


//todo
// - fine tune api permissions (e.g. delegated access)
// - fine tune welcome group selection - discuss with Steph
// - figure out issue with lt (less than) in query with createdDateTime - done but not elegant
// - check that user is not already a member of group before adding


namespace appsvc_fnc_dev_gcxonboardsync_dotnet001
{
    public class Function1
    {
        [FunctionName("Function1")]
        //public void Run([TimerTrigger("0 */5 * * * *")]TimerInfo myTimer, ILogger log)
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string assignedGroupId = config["assignedGroupId"];
            string AzureWebJobsStorage = config["AzureWebJobsStorageSync"];
            string containerName = config["containerName"];
            string fileNameSuffix = config["fileNameSuffix"];
            string listId = config["departmentSyncListId"];
            string siteId = config["siteId"];
            string welcomeGroupIds = config["welcomeGroupIds"];

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            List<ListItem> departmentList = await GetSyncedDepartmentList(graphAPIAuth, siteId, listId, log);

            if (departmentList != null)
            {
                log.LogInformation($"departmentList.Count = {departmentList.Count}");
                foreach (ListItem item in departmentList)
                {
                    string Abbreviation = item.Fields.AdditionalData["Abbreviation"].ToString();
                    DateTime? LastSyncDate = item.Fields.AdditionalData.Keys.Contains("LastSyncDate") ? (DateTime)item.Fields.AdditionalData["LastSyncDate"] : null;
                    string RGCode = item.Fields.AdditionalData["RGCode"].ToString();
                    string itemId = item.Id;
                    var groupId = GetSecurityGroupId(Abbreviation, AzureWebJobsStorage, containerName, fileNameSuffix, log);

                    if (LastSyncDate != null)
                    {
                        // Fetch all user base on last run – Can we use delegated access? Having access to a service account to be owner of the group? Only read access
                        var userIds = GetUsersToOnboard(graphAPIAuth, groupId, (DateTime)LastSyncDate, log);
                        string welcomeGroupId = GetWelcomeGroupId(graphAPIAuth, welcomeGroupIds, log).Result;
                        await AssignUserstoGroups(graphAPIAuth, userIds.Result, assignedGroupId, welcomeGroupId, log);
                        await UpdateLastSyncDate(graphAPIAuth, siteId, listId, itemId, log);
                    }
                    else
                    {
                        // ERROR – (Need a way to get this info… Maybe in the sp list?)
                    }
                }
            }
            else
            {
                log.LogInformation("null synced department list");
            }

            return new OkResult();
        }

        private static async Task<List<ListItem>> GetSyncedDepartmentList(GraphServiceClient graphAPIAuth, string siteId, string listId, ILogger log)
        {
            log.LogInformation("GetSyncedDepartmentList received a request.");

            List<ListItem> itemList = new List<ListItem>();

            try
            {
                var items = await graphAPIAuth.Sites[siteId].Lists[listId].Items.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Expand = new string[] { "fields($select=Abbreviation,RGCode,LastSyncDate)" };
                });

                itemList.AddRange(items.Value);

                while (items.OdataNextLink != null)
                {
                    var nextPageRequestInformation = new RequestInformation
                    {
                        HttpMethod = Method.GET,
                        UrlTemplate = items.OdataNextLink
                    };

                    items = await graphAPIAuth.RequestAdapter.SendAsync(nextPageRequestInformation, (parseNode) => new ListItemCollectionResponse());
                    itemList.AddRange(items.Value);
                }
            }
            catch (ODataError odataError)
            {
                log.LogError(odataError.Error.Code);
                log.LogError(odataError.Error.Message);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("GetSyncedDepartmentList processed a request.");

            return itemList;
        }


        public class SyncConfig
        {
            public string Enabled { get; set; }
            public string TenantCacheRefreshEnabled { get; set; }
            public string DeptAlias { get; set; }
            public string B2BGroupSyncAlias { get; set; }
            public string B2BGroupSyncSource { get; set; }
            public string DataContainerName { get; set; }
            public string NotifyUsersOnAccessGranted { get; set; }
            public string NotifyUsersOnAccessRemoved { get; set; }
            public string AccessGrantedNotificationTemplate { get; set; }
            public string AccessGrantedNotificationSubject { get; set; }
            public string AccessRemovedNotificationTemplate { get; set; }
            public string AccessRemovedNotificationSubject { get; set; }
            public string[] EmailNotificationListForUsersThatCannotBeInvited { get; set; }
            //public GroupAliasToResourceTenantGroup GroupAliasToResourceTenantGroupObjectIdMapping { get; set; }  // -> CIPHER_Group -> ResourceTenantGroupObjectId
            public object GroupAliasToResourceTenantGroupObjectIdMapping { get; set; }  // -> CIPHER_Group -> ResourceTenantGroupObjectId

            // 	"GroupAliasToResourceTenantGroupObjectIdMapping": {
            //         "CIPHER_Group": { 
            //             "ResourceTenantGroupObjectId": "3928b327-581b-480d-9a69-46a6e6f0bab9"
            //         }
            //     }
        }

        // Refactor this shit!!

        public class GroupAliasToResourceTenantGroup {
            public ResourceTenantGroup TenantGroup;
        }

        public class ResourceTenantGroup {
            public string ResourceTenantGroupObjectId;
        }

        private static string GetSecurityGroupId(string departmentAbbreviation, string AzureWebJobsStorage, string containerName, string fileNameSuffix, ILogger log)
        {
            log.LogInformation("GetSecurityGroupId received a request.");

            string fileName = $"{departmentAbbreviation}{fileNameSuffix}";

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(AzureWebJobsStorage);
            CloudBlobClient serviceClient = storageAccount.CreateCloudBlobClient();
            CloudBlobContainer container = serviceClient.GetContainerReference($"{containerName}");
            CloudBlockBlob blob = container.GetBlockBlobReference($"{fileName}");

            string contents = blob.DownloadTextAsync().Result;
            var result = JsonConvert.DeserializeObject<SyncConfig>(contents);

            // get groupId - this is messy, consider a more elegant implementation
            string mapping = result.GroupAliasToResourceTenantGroupObjectIdMapping.ToString();
            const string searchFor = "\"ResourceTenantGroupObjectId\": \"";
            int startIndex = mapping.IndexOf(searchFor) + searchFor.Length;
            string groupId = mapping.Substring(startIndex, 36); // e.g. length of 3928b327-581b-480d-9a69-46a6e6f0bab9

            log.LogInformation("GetSecurityGroupId processed a request.");

            return groupId;
        }

        private static async Task<List<string>> GetUsersToOnboard(GraphServiceClient graphAPIAuth, string securityGroupId, DateTime LastSyncDate, ILogger log)
        {
            log.LogInformation("GetUsersToOnboard received a request.");


            log.LogInformation($"securityGroupId: {securityGroupId}");

            List<string> userIds = new();
            var format = "yyyy-MM-ddTHH:mm:ssK";

            try
            {
               
                var members = await graphAPIAuth.Groups[securityGroupId].Members.GraphUser.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Count = true;
                    requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName", "createdDateTime" };

                    // Error: Request_UnsupportedQuery Operator: 'Less' is not supported
                    // cannot use less than operator so need to do another check in the foreach loop
                    //requestConfiguration.QueryParameters.Filter = $"createdDateTime lt {LastSyncDate.ToString(format)}";
                    requestConfiguration.QueryParameters.Filter = $"createdDateTime le {LastSyncDate.ToString(format)}";

                    requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                });

                foreach (User member in members.Value)
                {
                    if (member.CreatedDateTime < LastSyncDate)
                    {
                        userIds.Add(member.Id);
                    }
                }
            }
            catch (ODataError odataError)
            {
                log.LogError(odataError.Error.Code);
                log.LogError(odataError.Error.Message);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("GetUsersToOnboard processed a request.");

            return userIds;
        }

        private static async Task<string> GetWelcomeGroupId(GraphServiceClient graphAPIAuth, string welcomeGroupIds, ILogger log)
        {
            log.LogInformation("GetWelcomeGroupId received a request.");

            string welcomeGroupId = null;

            try
            {

                var retryHandlerOption = new RetryHandlerOption
                {
                    MaxRetry = 7,
                    ShouldRetry = (delay, attempt, message) => true
                };




                foreach (string groupId in welcomeGroupIds.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var MemberCount = await graphAPIAuth.Groups[groupId].Members.Count.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                    });

                    //  Get member count of first welcome group – Already have service account available for this
                    // 	if lest than ## add user
                    // 	if else- fetch next group until available member group is found

                    if (MemberCount < 25000)
                    {
                        welcomeGroupId = groupId;
                        break;
                    }
                }
            }
            catch (ODataError odataError)
            {
                log.LogError(odataError.Error.Code);
                log.LogError(odataError.Error.Message);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("GetWelcomeGroupId processed a request.");

            return welcomeGroupId;
        }

        private static async Task<bool> AssignUserstoGroups(GraphServiceClient graphAPIAuth, List<string> userIds, string assignedGroupId, string welcomeGroupId, ILogger log)
        {
            log.LogInformation("AssignUserstoGroups received a request.");
            
            int userCount = userIds.Count;
            List<string> userList = new();

            foreach (string userId in userIds)
            {
                userList.Add($"https://graph.microsoft.com/v1.0/directoryObjects/{userId}");
            }

            try {

                var requestBody = new Group
                {
                    AdditionalData = new Dictionary<string, object>
                    {
                        {
                            "members@odata.bind" , userList
                        }
                    }
                };

                foreach(var user in userList)
                {
                    log.LogInformation($"user: {user}");

                }

                // error when user already a member
                // Request_BadRequest
                // One or more added object references already exist for the following modified properties: 'members'.

                log.LogInformation($"assignedGroupId: {assignedGroupId}");
                log.LogInformation($"welcomeGroupId: {welcomeGroupId}");

                await graphAPIAuth.Groups[assignedGroupId].PatchAsync(requestBody);
                await graphAPIAuth.Groups[welcomeGroupId].PatchAsync(requestBody);
            }
            catch (ODataError odataError)
            {
                log.LogError(odataError.Error.Code);
                log.LogError(odataError.Error.Message);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("AssignUserstoGroups processed a request.");

            return true;
        }

        private static async Task<bool> UpdateLastSyncDate(GraphServiceClient graphAPIAuth, string siteId, string listId, string itemId, ILogger log)
        {
            log.LogInformation("UpdateLastSyncDate received a request.");

            try
            {
                var fieldValueSet = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"LastSyncDate", DateTime.Now}
                    }
                };

                await graphAPIAuth.Sites[siteId].Lists[listId].Items[itemId].Fields.PatchAsync(fieldValueSet);
            }
            catch (ODataError odataError)
            {
                log.LogError(odataError.Error.Code);
                log.LogError(odataError.Error.Message);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("UpdateLastSyncDate processed a request.");

            return true;
        }
    }
}
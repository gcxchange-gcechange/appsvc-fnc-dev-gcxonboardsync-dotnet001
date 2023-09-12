using System;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using System.Threading.Tasks;
using Microsoft.Kiota.Abstractions;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;
using Microsoft.WindowsAzure.Storage;
using Microsoft.Graph;
using static appsvc_fnc_dev_gcxonboardsync_dotnet001.Auth;

//todo
// - figure out issue with lt (less than) in query with createdDateTime - done but not elegant

// - check that user is not already a member of group before adding
//      - One or more added object references already exist for the following modified properties: 'members'.
//      - ok to ignore error for now?
//      - email notification on failure?

namespace appsvc_fnc_dev_gcxonboardsync_dotnet001
{
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

        // Refactor this!!

        //public class GroupAliasToResourceTenantGroup {
        //    public ResourceTenantGroup TenantGroup;
        //}

        //public class ResourceTenantGroup {
        //    public string ResourceTenantGroupObjectId;
        //}
    }

    public class OnboardUsers
    {
        [FunctionName("OnboardUsers")]
        public static async Task RunAsync([TimerTrigger("0 */5 * * * *")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"OnboardUsers received a request.");

            const int WelcomeGroupMemberLimit = 24900;

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string assignedGroupId = config["assignedGroupId"];
            string AzureWebJobsStorage = config["AzureWebJobsStorageSync"];
            string containerName = config["containerName"];
            string fileNameSuffix = config["fileNameSuffix"];
            string listId = config["departmentSyncListId"];
            string siteId = config["siteId"];
            string welcomeGroupIds = config["welcomeGroupIds"];

            var onboardUserClient = new GraphServiceClient(new ROPCConfidentialTokenCredential(config["onboardUserName"], config["onboardUserSecret"], log));
            var welcomeUserClient = new GraphServiceClient(new ROPCConfidentialTokenCredential(config["welcomeUserName"], config["welcomeUserSecret"], log));

            List<ListItem> departmentList = await GetSyncedDepartmentList(onboardUserClient, siteId, listId, log);

            if (departmentList != null)
            {
                log.LogInformation($"departmentList.Count = {departmentList.Count}");
                foreach (ListItem item in departmentList)
                {
                    string Abbreviation = item.Fields.AdditionalData["Abbreviation"].ToString();
                    DateTime? LastSyncDate = item.Fields.AdditionalData.Keys.Contains("LastSyncDate") ? (DateTime)item.Fields.AdditionalData["LastSyncDate"] : null;
                    //string RGCode = item.Fields.AdditionalData["RGCode"].ToString();
                    string itemId = item.Id;
                    var groupId = GetSecurityGroupId(Abbreviation, AzureWebJobsStorage, containerName, fileNameSuffix, log);

                    if (LastSyncDate != null)
                    {
                        var userIds = GetUsersToOnboard(onboardUserClient, groupId, (DateTime)LastSyncDate, log);
                        string welcomeGroupId = GetWelcomeGroupId(welcomeUserClient, welcomeGroupIds, userIds.Result.Count, WelcomeGroupMemberLimit, log).Result;
                        await AssignUserstoGroups(welcomeUserClient, userIds.Result, assignedGroupId, welcomeGroupId, log);
                        await UpdateLastSyncDate(onboardUserClient, siteId, listId, itemId, log);
                    }
                    else
                    {
                        // ERROR – (Need a way to get this info… Maybe in the sp list?)
                        // What if we make it a mandatory field to prevent this error?
                    }
                }
            }
            else
            {
                log.LogInformation("null synced department list");
            }
            
            log.LogInformation($"OnboardUsers processed a request.");
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

        private static async Task<string> GetWelcomeGroupId(GraphServiceClient graphAPIAuth, string welcomeGroupIds, int UserCount, int WelcomeGroupMemberLimit, ILogger log)
        {
            log.LogInformation("GetWelcomeGroupId received a request.");

            string welcomeGroupId = null;

            try
            {
                foreach (string groupId in welcomeGroupIds.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var MemberCount = await graphAPIAuth.Groups[groupId].Members.Count.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                    });

                    if ((MemberCount + UserCount) <= WelcomeGroupMemberLimit)
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

            List<string> userList = new();

            foreach (string userId in userIds)
            {
                userList.Add($"https://graph.microsoft.com/v1.0/directoryObjects/{userId}");
            }

            log.LogInformation($"Add to assigned group id = {assignedGroupId}");
            try
            {
                await graphAPIAuth.Groups[assignedGroupId].PatchAsync(new Group { AdditionalData = new Dictionary<string, object> { { "members@odata.bind", userList } } });
            }
            catch (ODataError odataError)
            {
                log.LogError($"odataError.Error.Code: {odataError.Error.Code}");
                log.LogError($"odataError.Error.Message: {odataError.Error.Message}");
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation($"Add to welcome group id = {welcomeGroupId}");
            try
            {
                await graphAPIAuth.Groups[welcomeGroupId].PatchAsync(new Group { AdditionalData = new Dictionary<string, object> { { "members@odata.bind", userList } } });
            }
            catch (ODataError odataError)
            {
                log.LogError($"odataError.Error.Code: {odataError.Error.Code}");
                log.LogError($"odataError.Error.Message: {odataError.Error.Message}");
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

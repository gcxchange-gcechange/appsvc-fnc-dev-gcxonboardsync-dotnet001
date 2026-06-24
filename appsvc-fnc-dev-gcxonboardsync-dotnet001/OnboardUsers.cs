using System;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using System.Threading.Tasks;
using Microsoft.Kiota.Abstractions;
using Microsoft.Graph.Models.ODataErrors;
using Newtonsoft.Json;
using Microsoft.Graph;
using static appsvc_fnc_dev_gcxonboardsync_dotnet001.Auth;
using System.Linq;

namespace appsvc_fnc_dev_gcxonboardsync_dotnet001
{
    public class OnboardUsers
    {
        private class WelcomeGroup
        {
            public string Id;
            public int MemberCount;
        }

        private static List<WelcomeGroup> WelcomeGroups;

        // Timer: at minute 30 past every 2nd hour
        [FunctionName("OnboardUsers")]
        public static async Task RunAsync([TimerTrigger("0 30 */2 * * *")] TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"OnboardUsers received a request.");

            const int AssignUserLimit = 20;
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

            string abbreviation;
            string groupId;
            string itemId;
            DateTime? LastSyncDate;
            string RGCode;

            Auth auth = new Auth();
            var graphClient = auth.graphAuth(log);

            try
            {
                List<ListItem> departmentList = await GetSyncedDepartmentList(onboardUserClient, siteId, listId, log);

                foreach (var department in departmentList)
                {
                    abbreviation = department.Fields.AdditionalData["Abbreviation"].ToString();
                    RGCode = department.Fields.AdditionalData["RGCode"].ToString();

                    var groups = await graphClient.Groups.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Count = true;
                        requestConfiguration.QueryParameters.Filter = "NOT(groupTypes/any(c:c eq 'Unified'))";
                        requestConfiguration.QueryParameters.Search = $"\"displayName:{RGCode}_{abbreviation}_B2B_Sync\"";
                        requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName" };
                        requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                    });

                    if (groups.Value.Count > 0)
                    {
                        // there should only be one group so get the first value
                        var group = groups.Value[0];

                        groupId = group.Id;
                        itemId = department.Id;
                        LastSyncDate = department.Fields.AdditionalData.Keys.Contains("LastSyncDate") ? (DateTime)department.Fields.AdditionalData["LastSyncDate"] : null;

                        log.LogInformation($"group.DisplayName = {group.DisplayName}");
                        log.LogInformation($"RGCode = {RGCode}");
                        log.LogInformation($"abbreviation = {abbreviation}");
                        log.LogInformation($"LastSyncDate = {LastSyncDate}");
                        log.LogInformation($"itemId = {itemId}");

                        WelcomeGroups = await GetWelcomeGroups(graphClient, welcomeGroupIds, log);

                        if (LastSyncDate != null)
                        {
                            var userIds = GetUsersToOnboard(onboardUserClient, groupId, (DateTime)LastSyncDate, log);

                            log.LogInformation($"userIds.Result.Count: {userIds.Result.Count}");

                            if (userIds.Result.Count > 0)
                            {
                                string response = "";
                                int userIdCount = 0;
                                List<string> userIdList = new();
                                string welcomeGroupId;

                                foreach (string userId in userIds.Result)
                                {
                                    userIdList.Add(userId);
                                    userIdCount += 1;

                                    if (userIdCount == AssignUserLimit)
                                    {
                                        welcomeGroupId = GetWelcomeGroupId(userIdCount, WelcomeGroupMemberLimit, log);
                                        response = await AssignUserstoGroups(welcomeUserClient, userIdList, assignedGroupId, welcomeGroupId, log);

                                        userIdCount = 0;
                                        userIdList.Clear();

                                        if (response != "")
                                        {
                                            break;  // an error occurred, break from the loop
                                        }
                                    }
                                }

                                // catch any stragglers and process if no errors, i.e. response equals empty string
                                if ((userIdCount > 0) && (response == ""))
                                {
                                    welcomeGroupId = GetWelcomeGroupId(userIdCount, WelcomeGroupMemberLimit, log);
                                    response = await AssignUserstoGroups(welcomeUserClient, userIdList, assignedGroupId, welcomeGroupId, log);
                                }

                                if (response == "")
                                {
                                    await UpdateLastSyncDate(onboardUserClient, siteId, listId, itemId, log);
                                }
                                else
                                {
                                    string details = $"{response.Replace(Environment.NewLine, "<br / >")}<br />Department: {abbreviation}<br />Group Id: {groupId}<br />User Ids: {JsonConvert.SerializeObject(userIds.Result)}";
                                    sendEmail(config["emailUserName"], config["emailUserSecret"], config["recipientAddress"], "Assign users failed", details, log);
                                }
                            }
                        }
                        else
                        {
                            log.LogInformation($"Null sync date - Department: {abbreviation} - Group Id: {groupId}");
                            string details = $"Department: {abbreviation}<br />Group Id: {groupId}";
                            sendEmail(config["emailUserName"], config["emailUserSecret"], config["recipientAddress"], "Null sync date", details, log);
                        }
                    }
                    else
                    {
                        sendEmail(config["emailUserName"], config["emailUserSecret"], config["recipientAddress"], "Can't find synced department", abbreviation, log);
                        continue;  // continue with next item in the foreach loop
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
                    requestConfiguration.QueryParameters.Top = 999;

                    // Error: Request_UnsupportedQuery Operator: 'Less' is not supported
                    // cannot use less than operator (lt) so need to do another check in the foreach loop
                    // https://learn.microsoft.com/en-us/graph/aad-advanced-queries?tabs=http#application-properties
                    requestConfiguration.QueryParameters.Filter = $"createdDateTime ge {LastSyncDate.ToString(format)}";

                    requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                });

                foreach (User member in members.Value)
                {
                    if (member.CreatedDateTime > LastSyncDate)
                    {
                        userIds.Add(member.Id);
                    }
                }

                while (members.OdataNextLink != null)
                {
                    var nextPageRequestInformation = new RequestInformation
                    {
                        HttpMethod = Method.GET,
                        UrlTemplate = members.OdataNextLink
                    };
                    members = await graphAPIAuth.RequestAdapter.SendAsync(nextPageRequestInformation, (parseNode) => new UserCollectionResponse());

                    foreach (User member in members.Value)
                    {
                        if (member.CreatedDateTime > LastSyncDate)
                        {
                            userIds.Add(member.Id);
                        }
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

        private static string GetWelcomeGroupId(int UserCount, int WelcomeGroupMemberLimit, ILogger log)
        {
            log.LogInformation("GetWelcomeGroupId received a request.");

            string welcomeGroupId = null;
            int? idx = null;
            int combinedCount = 0;

            try
            {
                for (int i = 0; i < WelcomeGroups.Count; i++)
                {
                    combinedCount = WelcomeGroups[i].MemberCount + UserCount;

                    if (combinedCount <= WelcomeGroupMemberLimit)
                    {
                        idx = i;
                        break;
                    }
                }

                if (idx != null)
                {
                    welcomeGroupId = WelcomeGroups[(int)idx].Id;
                    WelcomeGroups[(int)idx].MemberCount = combinedCount;
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

        private static async Task<List<WelcomeGroup>> GetWelcomeGroups(GraphServiceClient graphAPIAuth, string welcomeGroupIds, ILogger log)
        {
            log.LogInformation("GetWelcomeGroups received a request.");

            List<WelcomeGroup> welcomeGroups = new List<WelcomeGroup>();

            try
            {
                foreach (string groupId in welcomeGroupIds.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var MemberCount = await graphAPIAuth.Groups[groupId].Members.Count.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                    });

                    WelcomeGroup group = new WelcomeGroup
                    {
                        Id = groupId,
                        MemberCount = (int)MemberCount
                    };

                    welcomeGroups.Add(group);
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

            log.LogInformation("GetWelcomeGroups processed a request.");

            return welcomeGroups;
        }

        private static async Task<string> AssignUserstoGroups(GraphServiceClient graphAPIAuth, List<string> userIds, string assignedGroupId, string welcomeGroupId, ILogger log)
        {
            log.LogInformation("AssignUserstoGroups received a request.");

            List<string> userList = new();
            string errorMessage = "";

            foreach (string userId in userIds)
            {
                userList.Add($"https://graph.microsoft.com/v1.0/directoryObjects/{userId}");

                log.LogInformation($"userId: {userId}");
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

                if (odataError.Error.Code == "Request_BadRequest")
                {
                    log.LogWarning($"Retry assignment to assigned group after removing existing members");
                    errorMessage = await HandleDuplicates(graphAPIAuth, userIds.ToList(), assignedGroupId, log);
                }
                else
                {
                    errorMessage = $"{odataError.Error.Message} - assigned group id: {assignedGroupId}";
                }
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                errorMessage = $"{e.Message} - assigned group id: {assignedGroupId}";
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

                if (odataError.Error.Code == "Request_BadRequest")
                {
                    log.LogWarning($"Retry assignment to welcome group after removing existing members");
                    errorMessage = await HandleDuplicates(graphAPIAuth, userIds.ToList(), welcomeGroupId, log);
                }
                else
                {
                    errorMessage = errorMessage + Environment.NewLine + $"{odataError.Error.Message} - welcome group id: {welcomeGroupId}";
                }
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                errorMessage = errorMessage + Environment.NewLine + $"{e.Message} - welcome group id: {welcomeGroupId}";
            }

            log.LogInformation("AssignUserstoGroups processed a request.");

            return errorMessage;
        }

        private static async Task<string> HandleDuplicates(GraphServiceClient graphAPIAuth, List<string> userIds, string groupId, ILogger log)
        {
            string errorMessage = "";
            string idsToFilter = string.Format("'{0}'", string.Join("','", userIds));

            var members = await graphAPIAuth.Groups[groupId].Members.GraphUser.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Count = true;
                requestConfiguration.QueryParameters.Select = new string[] { "id" };
                requestConfiguration.QueryParameters.Filter = $"id in ({idsToFilter})";
                requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
            });

            foreach (User member in members.Value)
            {
                if (userIds.Contains(member.Id))
                {
                    userIds.Remove(member.Id);
                }
            }

            if (userIds.Count > 0)
            {
                log.LogInformation($"Retrying assignment of {userIds.Count} user(s) to groupId: {groupId}");
                List<string> userList = new();
                foreach (string userId in userIds)
                {
                    userList.Add($"https://graph.microsoft.com/v1.0/directoryObjects/{userId}");
                }

                try
                {
                    await graphAPIAuth.Groups[groupId].PatchAsync(new Group { AdditionalData = new Dictionary<string, object> { { "members@odata.bind", userList } } });
                }
                catch (ODataError odataError)
                {
                    log.LogError($"odataError.Error.Code: {odataError.Error.Code}");
                    log.LogError($"odataError.Error.Message: {odataError.Error.Message}");
                    errorMessage = $"{odataError.Error.Message} - retry for group id: {groupId}";
                }
                catch (Exception e)
                {
                    log.LogError($"Message: {e.Message}");
                    if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                    log.LogError($"StackTrace: {e.StackTrace}");
                    errorMessage = $"{e.Message} - retry for group id: {groupId}";
                }
            }
            else
            {
                log.LogInformation($"No retry because all users already assigned to groupId: {groupId}");
            }

            return errorMessage;
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

        public static async void sendEmail(string emailUserName, string emailUserSecret, string recipientAddress, string failureReason, string details, ILogger log)
        {
            var graphAPIAuth = new GraphServiceClient(new ROPCConfidentialTokenCredential(emailUserName, emailUserSecret, log));

            try
            {
                var requestBody = new Microsoft.Graph.Me.SendMail.SendMailPostRequestBody
                {
                    Message = new Message
                    {
                        Subject = $"GCX - Onboarding Error: {failureReason}",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = @$"<p>{details}</p>"
                        },
                        ToRecipients = new List<Recipient>
                        {
                            new Recipient { EmailAddress = new EmailAddress { Address =  $"{recipientAddress}" } }
                        }
                    }
                };

                await graphAPIAuth.Me.SendMail.PostAsync(requestBody);
            }
            catch (ODataError odataError)
            {
                log.LogError(odataError.Error.Code);
                log.LogError(odataError.Error.Message);
            }
            catch (Exception e)
            {
                log.LogInformation($"Error: {e.Message}");
            }
        }
    }
}
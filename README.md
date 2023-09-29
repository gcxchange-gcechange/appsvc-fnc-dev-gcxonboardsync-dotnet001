# GCX Onboarding Synchronization

## Summary

Syncs users between IDF (Identity Federation) and GCXchange 
- get a list of synced deparments
- loop through the list and perform the following actions for each department:
  - get the security group id from a json file in the storage account
  - use the group id to get the list of users to onboard since the last sync date
  - get the welcome group id from a pool of welcome groups based on capacity
  - assign the users to the assigned and welcome groups

## Prerequisites

The following user accounts (as reflected in the app settings) are required:

| Account         | Membership requirements                               |
| --------------- | ----------------------------------------------------- |
| emailUserName   | n/a                                                   |
| onboardUserName | Site with synced department list, assigned user group |
| welcomeUserName | welcome user group(s)                                 |

Note that user account design can be modified to suit your environment

## Version 

![dotnet 6](https://img.shields.io/badge/net6.0-blue.svg)

## API permission

MSGraph

| API / Permissions name    | Type      | Admin consent | Justification                       |
| ------------------------- | --------- | ------------- | ----------------------------------- |
| GroupMember.ReadWrite.All | Delegated | Yes           | Read and assign members to groups   |
| Mail.Send                 | Delegated | Yes           | Send failure notifications by email | 
| Sites.ReadWrite.All       | Delegated | Yes           | Read and update SharePoint list     |
| User.Read.All             | Delegated | Yes           | Read createdDateTime property       |

Sharepoint

n/a

## App setting

| Name                    | Description                                                                   |
| ----------------------- | ----------------------------------------------------------------------------- |
| assignedGroupId 		    | Object Id for the assigned users group                                        |
| AzureWebJobsStorage     | Connection string for the storage acoount                                     |
| AzureWebJobsStorageSync | Connection string for the storage account with sync data                      |
| clientId                | The application (client) ID of the app registration                           |
| containerName           | Than name of the container that hosts the sync files                          |
| departmentSyncListId    | Id of the SharePoint list for synced departments                              |
| emailUserName           | Email address used to send failure notifications                              |
| emailUserSecret         | Secret name for emailUserSecret                                               |
| fileNameSuffix          | The common suffix for filenames containing sync data                          |
| keyVaultUrl             | Address for the key vault                                                     |
| onboardUserName         | User principal name for the service account that performs onboarding tasks    |
| onboardUserSecret       | Secret name for onboardUserName                                               |
| recipientAddress        | Email address that received failure notifications                             |
| secretName              | Secret name used to authorize the function app                                |
| siteId                  | Id of the SharePoint site that hosts the list of synced departments           |
| tenantId                | Id of the Azure tenant that hosts the function app                            |
| welcomeGroupIds         | Comma separated list of Ids for the welcome group(s)                          |
| welcomeUserName         | User principal name for the service account that performs welcome group tasks |
| welcomeUserSecret       | Secret name for welcomeUserName                                               |

## Version history

Version|Date|Comments
-------|----|--------
1.0|TBD|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

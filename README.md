# Overview

This script is provided as a way to inventory SharePoint Sites, Webs, Lists, and Items and their permissions

# Script Reference

| Script | Description | Permissions Required | Dependencies | 
| --- | --- | --- | --- |
| Get-SPODetails.ps1 | Iterates through Site Collections, Webs, Lists, and Items to gather information at each level | [Details](https://github.com/jameswh3/SharePoint-Inventory-CSOM/tree/main?tab=readme-ov-file#permissions-required) | PnP PowerShell


# PowerShell Requirements

*   [Windows PowerShell 7.0 or higher](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4)
*   [PnP.PowerShell module 2.99.74 or higher](https://pnp.github.io/powershell/articles/installation.html)
*   [Entra ID Application Registered to use with PnP PowerShell]((https://pnp.github.io/powershell/articles/registerapplication))

# Script Details

## Get-SPODetails.ps1

Iterates through Site Collections, Webs, Lists, and Items to gather inventory information at each level

### Permissions Required

| API | Type | Permission | Justification |
| --- | --- | --- | --- |
| SharePoint | Application | Sites.FullControl.All | Required to retrieve tenant site properties. |
| Microsoft Graph | Application | Groups.ReadWrite.All | Required to retrieve M365 Group properties and associated endpoints. |
| Microsoft Graph | Application | User.ReadWriteAll | Required to retrieve User Details from Graph |

### Configuration

Before executing, update the lines below with your environment's parameters

``` PowerShell
#Runs the full script with default params
#Update the parameters with <> to reflect your environment
Get-SPODetails -ReportOutputPath "c:\temp\spinventory" `
    -ClientId "<Your Entra App Client Id>" `
    -CertificatePath "<Path to your PFX file>" `
    -Tenant "<Your tenant name>.onmicrosoft.com" `
    -SPOAdminUrl "https://<Your tenant name>-admin.sharepoint.com" `
    -GetWebDetails `
    -GetWebPermissions `
    -GetListDetails `
    -GetListPermissions `
    -IncludeSystemLists:$false `
    -GetItemPermissions `
    -GetItemDetails `
    -ClearPriorLogs
```
### Outputs

#### ItemDetails.csv

| Field Name | Description |
| --- | --- |
| Area | Area where script is running (e.g. Item) |
| ItemId | Item Id |
| SiteId | GUID of Site |
| WebId | GUID of Web |
| ListId | GUID of List |
| ItemName | Name of Item |
| FileRef | File Ref URL of Item |
| SensitivityLabel | Sensitivity Label |
| ComplianceTag | Compliance Tag |

#### ItemPermissions.csv

| Field Name | Description |
| --- | --- |
| Area | Area where script is running (e.g. Item) |
| MemberName | Name of Principal |
| PrincipalType | Principal Type |
| LoginName | Login Name of Principal |
| FileSystemObjectType | File Type |
| ItemUniqueId | GUID of Item |
| ItemId | Id of Item |
| HasUniqueRoleAssignments | Item Has Unique Role Assignments |
| ListId | GUID of List |
| WebId | GUID of Web |
| SiteId | GUID of Site |
| PermissionLevels | Permission Levels for the Principal at this Scope |

#### ListDetails.csv

| Field Name | Description |
| --- | --- |
| Area | Area where script is running (e.g. Item) |
| ListId | GUID of List |
| Title | List Title |
| IsSystemList | Is System List |
| WebId | GUID of Web |
| SiteId | GUID of Site |
| RootFolder | Root folder of list |
| HasUniqueRoleAssignments | List Has Unique Role Assignments |
| ItemCount | Number of Items in List |
| LastItemDeletedDate | Last Item Deleted Date |
| LastItemModifiedDate | Last Item Modified Date |
| LastItemUserModifiedDate | Last Modified by a User Date |

#### ListPermissions.csv

| Field Name | Description |
| --- | --- |
| Area | Area where script is running (e.g. Item) |
| MemberName | Name of Principal |
| PrincipalType | Principal Type |
| ListId | GUID of List |
| WebId | GUID of Web |
| SiteId | GUID of Site |
| PermissionLevels | Permission Levels for the Principal at this Scope |

#### SiteDetails.csv

| Field Name | Description |
| --- | --- |
| Area | Area where script is running (e.g. Item) |
| SiteId | GUID of Site |
| Url | Url of Site |
| GroupId | ID of Group |
| Storage | Storage in MB |
| RootWeb | Root Web of Site |
| SiteOwner | Site Owner |
| SharingCapability | External Sharing Capability Setting |
| GroupVisibility | Group Visibility |
| HasTeam | Has Team |
| ResourceProvisioningOptions | Other Resources Associated with Site/Group |
| SiteSensitivityLabel | Site Sensitivity Label |
| GroupOwners | Group Owners (if Group Enabled) |
| IsRestrictedAccessControlPolicyEnforcedOnSite | Restricted Access Control Policy Enforced |
| IsRestrictContentOrgWideSearchPolicyEnforcedOnSite | Restricted Content Org Wide Search Policy Enforced |
| DisableCompanyWideSharingLinks | Disable Company Sharing Links |

#### WebDetails.csv

| Field Name | Description |
| --- | --- |
| Area | Area where script is running (e.g. Item)esc |
| Url | Url of Web |
| WebId | GUID of Web |
| WebTitle | Title of Web |
| SiteId | GUID of Site |
| ParentWebId | GUID of Parent Web |
| WebTemplate | Web Tempate |
| LastItemUserModifiedDate | Last Item User Modified Date |
| LastItemModifiedDate | Last Item Modified Date |
| HasUniqueRoleAssignments | Has Unique Role Assignments |
| NoCrawl | No Crawl Setting |

#### WebGroupDetail.csv

| Field Name | Description |
| --- | --- |
| Area | Area where script is running (e.g. Item) |
| WebId | GUID of Web |
| SiteId | GUID of Site |
| GroupName | Name of Group |
| GroupId | ID of Group |
| UserId | ID of User |
| LoginName | Login Name of User |
| Email | Email of User |
| IsSiteAdmin | Is Site Admin |

#### WebPermissions.csv

| Field Name | Description |
| --- | --- |
| Area | Area where script is running (e.g. Item) |
| MemberName | Principal Name |
| PrincipalType | Principal Type |
| PrincipalId | Id of Principal |
| Description | Principal Description |
| Url | Url of Web |
| WebId | GUID of Web |
| SiteId | GUID of Site |
| PermissionLevels | Permission Levels for the Principal at this Scope |
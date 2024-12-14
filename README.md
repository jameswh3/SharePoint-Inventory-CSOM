# Overview

This script is provided as a way to inventory SharePoint Sites, Webs, Lists, and Items and their permissions

# Script Reference

<table><tbody><tr><td>Script</td><td>Description</td><td>Permissions Required</td><td>Dependencies</td></tr><tr><td>Get-SPODetails.ps1</td><td>Iterates through Site Collections, Webs, Lists, and Items to gather inventory information at each level</td><td><a href="https://github.com/jameswh3/SharePoint-Inventory-CSOM/tree/main?tab=readme-ov-file#Permissions Required">Details</a></td><td>PnP PowerShell</td></tr></tbody></table>

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

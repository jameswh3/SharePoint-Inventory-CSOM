#Requires -Version 7.0
#Requires -Modules @{ModuleName="PnP.PowerShell"; ModuleVersion="2.99.57"}
Function Get-SPOWebPermissionDetails {
    param(
        [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.Web]$Web,
        $ReportOutputPath,
        $SiteId
    )
    BEGIN {
        Get-PnPProperty -ClientObject $web -Property RoleAssignments,HasUniqueRoleAssignments,ParentWeb
        $PermissionData = @()
        $Area="Web"
        $ReportOutput=($ReportOutputPath + "\"+$Area+"Permissions.csv")
    } #begin
    PROCESS {
        Foreach ($RoleAssignment in ($Web).RoleAssignments) {
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Area($Area)           
            $Permissions | Add-Member NoteProperty MemberName($RoleAssignment.Member.Title)
            $Permissions | Add-Member NoteProperty PrincipalType($RoleAssignment.Member.PrincipalType)
            $Permissions | Add-Member NoteProperty PrincipalId($RoleAssignment.PrincipalId)
            $Permissions | Add-Member NoteProperty Description($RoleAssignment.Member.Description)
            $Permissions | Add-Member NoteProperty Url($Web.Url)
            $Permissions | Add-Member NoteProperty WebId($Web.Id)
            $Permissions | Add-Member NoteProperty SiteId($SiteId)
            $Permissions | Add-Member NoteProperty PermissionLevels(($RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join ",")
            $PermissionData += $Permissions
            If ($RoleAssignment.Member.PrincipalType -eq "SharePointGroup") {
                Get-SPOGroupMembers -Principal $RoleAssignment.Member `
                    -ReportOutputPath $ReportOutputPath `
                    -WebId $Web.Id `
                    -SiteId $SiteId
            } #if sharepoint group
        } #foreach roleassignment
    } #Process
    END {
        $PermissionData | Export-Csv -Path $ReportOutput -NoTypeInformation -Append
    } #End
}
Function Get-SPOGroupMembers {
    param (
        [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.Principal]$Principal,
        $ReportOutputPath,
        $WebId,
        $SiteId
    )
    BEGIN {
        $Users=Get-PnPGroupMember -Group $Principal
        $UseGroupData = @()
        $Area="WebGroup"
        $ReportOutput=($ReportOutputPath + "\"+$Area+"Detail.csv")
    }
    PROCESS {     
        foreach ($user in $Users) {
            $userGroupDatum = New-Object PSObject
            $userGroupDatum | Add-Member NoteProperty Area($Area)
            $userGroupDatum | Add-Member NoteProperty WebId($WebId)
            $userGroupDatum | Add-Member NoteProperty SiteId($SiteId)
            $userGroupDatum | Add-Member NoteProperty GroupName($Principal.Title)
            $userGroupDatum | Add-Member NoteProperty GroupId($Principal.Id)
            $userGroupDatum | Add-Member NoteProperty UserId($user.Id)
            $userGroupDatum | Add-Member NoteProperty LoginName($user.LoginName -replace "i:0#\.f\|membership\|", "")
            $userGroupDatum | Add-Member NoteProperty Email($user.Email)
            $userGroupDatum | Add-Member NoteProperty IsSiteAdmin($user.IsSiteAdmin)
            $UseGroupData += $userGroupDatum
        }

    }
    END {
        $UseGroupData | Export-Csv -Path $ReportOutput -NoTypeInformation -Append
    }
}
Function Get-SPOListPermissionDetails {
    param(
        [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.List]$List,
        $ReportOutputPath,
        $WebId,
        $SiteId,
        [Switch]$GetItemPermissions
    )
    BEGIN {
        Write-Host "------Processing List Permissions - $($List.Title)"
        Get-PnPProperty -ClientObject $List -Property RoleAssignments,HasUniqueRoleAssignments,DefaultViewUrl
        $PermissionData = @()
        $Area="List"
        $ReportOutput=($ReportOutputPath + "\"+$Area+"Permissions.csv")
    }
    PROCESS {
        Foreach ($RoleAssignment in ($List).RoleAssignments) {
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Area($Area)
            $Permissions | Add-Member NoteProperty MemberName($RoleAssignment.Member.Title)
            $Permissions | Add-Member NoteProperty PrincipalType($RoleAssignment.Member.PrincipalType)
            $Permissions | Add-Member NoteProperty ListId($List.Id)
            $Permissions | Add-Member NoteProperty WebId($WebId)
            $Permissions | Add-Member NoteProperty SiteId($SiteId)
            $Permissions | Add-Member NoteProperty PermissionLevels(($RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join ",")
            $PermissionData += $Permissions
        }
    }
    END {
        $PermissionData | Export-Csv -Path $ReportOutput -NoTypeInformation -Append
    }
}
Function Get-SPOItemPermissionDetails {
    param(
        [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ListItem]$Item,
        $ReportOutputPath,
        $WebId,
        $SiteId,
        $ListId
    )
    BEGIN {
        Get-PnPProperty -ClientObject $Item -Property RoleAssignments,DisplayName
        $PermissionData = @()
        $Area="Item"
        $ReportOutput=($ReportOutputPath + "\"+$Area+"Permissions.csv")
        Write-Host "--------Processing Item Permissions: $($Item.DisplayName)"
    }
    PROCESS {
        Foreach ($RoleAssignment in ($Item).RoleAssignments) {
            Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Area($Area)
            $Permissions | Add-Member NoteProperty MemberName($RoleAssignment.Member.Title)
            $Permissions | Add-Member NoteProperty PrincipalType($RoleAssignment.Member.PrincipalType)
            $Permissions | Add-Member NoteProperty DisplayName($Item.DisplayName)
            $Permissions | Add-Member NoteProperty FileSystemObjectType($Item.FileSystemObjectType)
            $Permissions | Add-Member NoteProperty UniqueId($Item.FieldValues.UniqueId)
            $Permissions | Add-Member NoteProperty ItemId($Item.Id)
            $Permissions | Add-Member NoteProperty HasUniqueRoleAssignments($Item.HasUniqueRoleAssignments)
            $Permissions | Add-Member NoteProperty ListId($ListId)
            $Permissions | Add-Member NoteProperty WebId($WebId)
            $Permissions | Add-Member NoteProperty SiteId($SiteId)
            $Permissions | Add-Member NoteProperty PermissionLevels(($RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name) -join ",")
            $PermissionData += $Permissions
        }
    }
    END {
        $PermissionData | Export-Csv -Path $ReportOutput -NoTypeInformation -Append
    }
}
Function Get-SPOWebDetails {
    param (
        [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.Web]$Web,
        $ReportOutputPath,
        [Switch]$GetListPermissions,
        [Switch]$GetItemPermissions,
        [switch]$GetListDetails,
        [switch]$IncludeSystemLists,
        [switch]$GetWebPermissions,
        [switch]$GetItemDetails,
        $SiteId
    )
    BEGIN {
        $WebData = @()
        $Area="Web"
        $ReportOutput=($ReportOutputPath + "\"+$Area+"Details.csv")
        Get-PnPProperty -ClientObject $Web -Property ParentWeb,WebTemplate,LastItemModifiedDate,LastItemUserModifiedDate,Title,HasUniqueRoleAssignments,NoCrawl
    } #begin
    PROCESS {
        Write-Host "--Processing Web: $($Web.Url)"
        $WebDatum = New-Object PSObject
        $WebDatum | Add-Member NoteProperty Area($Area)
        $WebDatum | Add-Member NoteProperty Url($Web.Url)
        $WebDatum | Add-Member NoteProperty WebId($Web.Id)
        $WebDatum | Add-Member NoteProperty WebTitle($Web.Title)
        $WebDatum | Add-Member NoteProperty SiteId($SiteId)
        $WebDatum | Add-Member NoteProperty ParentWebId($Web.ParentWeb.Id)
        $WebDatum | Add-Member NoteProperty WebTemplate($Web.WebTemplate)
        $WebDatum | Add-Member NoteProperty LastItemUserModifiedDate($Web.LastItemModifiedDate)
        $WebDatum | Add-Member NoteProperty LastItemModifiedDate($Web.LastItemUserModifiedDate)
        $WebDatum | Add-Member NoteProperty HasUniqueRoleAssignments($Web.HasUniqueRoleAssignments)
        $WebDatum | Add-Member NoteProperty NoCrawl($Web.NoCrawl)
        $WebData += $WebDatum
        If ($GetWebPermissions) {
            Get-SPOWebPermissionDetails -Web $web `
                -ReportOutputPath $ReportOutputPath `
                -SiteId $SiteId
        }
        If ($GetListDetails -or 
            $GetListPermissions -or 
            $GetItemPermissions) {
            $Lists = Get-PnPList -Includes HasUniqueRoleAssignments,DefaultViewUrl,IsSystemList
            foreach ($List in $Lists) {
                if ($IncludeSystemLists) {
                    Get-SPOListDetails -List $List `
                        -ReportOutputPath $ReportOutputPath `
                        -GetItemPermissions:$GetItemPermissions `
                        -GetListPermissions:$GetListPermissions `
                        -WebId $web.Id `
                        -SiteId $SiteId `
                        -GetItemDetails:$GetItemDetails
                } #if include systemlists
                elseif (-not($list.IsSystemList)) {
                    Get-SPOListDetails -List $List `
                    -ReportOutputPath $ReportOutputPath `
                    -GetItemPermissions:$GetItemPermissions `
                    -GetListPermissions:$GetListPermissions `
                    -WebId $web.Id `
                    -SiteId $SiteId `
                    -GetItemDetails:$GetItemDetails
                }
            } #foreach list
        } #if get list or item permissions

    } #process

    END {
        $WebData | Export-Csv -Path $ReportOutput -NoTypeInformation -Append
    } #end

} #Get-SPOWebDetails
Function Get-SPOListDetails {
    param (
        [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.List]$List,
        $ReportOutputPath,
        $WebId,
        $SiteId,
        [Switch]$GetListPermissions,
        [Switch]$GetItemPermissions,
        [Switch]$GetItemDetails
    )
    BEGIN {
        $ListData = @()
        $Area="List"
        $ReportOutput=($ReportOutputPath + "\"+$Area+"Details.csv")
        Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments,DefaultViewUrl,IsSystemList
    } #begin
    PROCESS {
        Write-Host "----Processing List: $($List.Title)"
        $ListDatum = New-Object PSObject
        $ListDatum | Add-Member NoteProperty Area($Area)
        $ListDatum | Add-Member NoteProperty ListId($List.Id)
        $ListDatum | Add-Member NoteProperty Title($List.Title)
        $ListDatum | Add-Member NoteProperty IsSystemList($List.IsSystemList)
        $ListDatum | Add-Member NoteProperty WebId($WebId)
        $ListDatum | Add-Member NoteProperty SiteId($SiteId)
        $ListDatum | Add-Member NoteProperty RootFolder($List.RootFolder.ServerRelativeUrl)
        $ListDatum | Add-Member NoteProperty HasUniqueRoleAssignments($List.HasUniqueRoleAssignments)
        $ListDatum | Add-Member NoteProperty ItemCount($List.ItemCount)
        $ListDatum | Add-Member NoteProperty LastItemDeletedDate($List.LastItemDeletedDate)
        $ListDatum | Add-Member NoteProperty LastItemModifiedDate($List.LastItemModifiedDate)
        $ListDatum | Add-Member NoteProperty LastItemUserModifiedDate($List.LastItemUserModifiedDate)
        $ListData += $ListDatum
        
        If ($List.HasUniqueRoleAssignments -and $GetListPermissions) {
                Get-SPOListPermissionDetails -List $List `
                    -ReportOutputPath $ReportOutputPath `
                    -GetItemPermissions:$GetItemPermissions `
                    -WebId $WebId `
                    -SiteId $SiteId
        } #if hasuniqueroleassignemnts and getlistpermissions
        If ($GetItemPermissions) {
            $items=Get-PnPListItem -list $List -Includes HasUniqueRoleAssignments,DisplayName
            foreach ($item in $items) {
                if ($GetItemDetails -or $GetItemPermissions) {
                    Get-SPOItemDetails -item $item `
                        -ReportOutputPath $ReportOutputPath `
                        -WebId $WebId `
                        -SiteId $SiteId `
                        -ListId $List.Id
                }
            } #foreach item in listitems
        } #if getitempermissions or getitemdetails
    } #process
    END {
        $ListData | Export-Csv -Path $ReportOutput -NoTypeInformation -Append
    } #end
} #get-spolistdetails
Function Get-SPOItemDetails {
    param(
        [Microsoft.SharePoint.Client.ListItem]$item,
        $WebId,
        $SiteId,
        $ListId,
        $ReportOutputPath
    )
    BEGIN {
        $ItemData = @()
        $Area="Item"
        $ReportOutput=($ReportOutputPath + "\"+$Area+"Details.csv")
        Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments | out-null #suppressing output b/c this returns
    }
    PROCESS {
        Write-Host "------Processing Item: $($Item.DisplayName)"
        $ItemDatum = New-Object PSObject
        $ItemDatum | Add-Member NoteProperty Area($Area)
        $ItemDatum | Add-Member NoteProperty ItemId($Item.Id)
        $ItemDatum | Add-Member NoteProperty SiteId($SiteId)
        $ItemDatum | Add-Member NoteProperty WebId($WebId)
        $ItemDatum | Add-Member NoteProperty ListId($ListId)
        $ItemDatum | Add-Member NoteProperty ItemName($Item.DisplayName)
        $ItemDatum | Add-Member NoteProperty FileRef($item.FieldValues["FileRef"])
        $ItemDatum | Add-Member NoteProperty SensitivityLabel($Item.FieldValues["_DisplayName"])
        $ItemDatum | Add-Member NoteProperty ComplianceTag($Item.FieldValues["_ComplianceTag"])
        $ItemData += $ItemDatum
    }
    END {
        $ItemData | Export-Csv -Path $ReportOutput -NoTypeInformation -Append
        if ($item.HasUniqueRoleAssignments) {
            Get-SPOItemPermissionDetails -Item $item `
                -ReportOutputPath $ReportOutputPath `
                -WebId $WebId `
                -SiteId $SiteId `
                -ListId $List.Id
        } #if itemhasuniquepermissions
    }
}
Function Get-SPOSiteDetails {
    param (
        $SiteUrl,
        $SharingCapability,
        $ClientId,
        $Tenant,
        $CertificatePath,
        $ReportOutputPath,
        [Switch]$GetWebPermissions,
        [Switch]$GetListPermissions,
        [Switch]$IncludeSystemLists,
        [Switch]$GetItemPermissions,
        [switch]$GetItemDetails,
        [switch]$GetListDetails,
        [switch]$GetWebDetails
    )
    BEGIN {
        Connect-PnPOnline -Url $SiteUrl `
            -ClientId $ClientId `
            -Tenant $Tenant `
            -CertificatePath $CertificatePath 
        $SiteData = @()
        $Area="Site"
        $ReportOutput=($ReportOutputPath + "\"+$Area+"Details.csv")
        $Webs = Get-PnPSubWeb -Recurse
        $Webs += Get-PnPWeb
    } #begin
    PROCESS {
        $SPOSite=Get-PnPSite -Includes Id,Owner,SecondaryContact,Usage,DisableCompanyWideSharingLinks,RootWeb.Url,GroupId,SensitivityLabel,IsRestrictContentOrgWideSearchPolicyEnforcedOnSite,IsRestrictedAccessControlPolicyEnforcedOnSite
        $SiteId=$SPOSite.Id
        if ($SPOSite.GroupId -ne "00000000-0000-0000-0000-000000000000" -and $SPOSite.GroupId -ne "") {
            $group=Get-PnPMicrosoft365Group -Identity $SPOSite.GroupId -IncludeOwners
        }
        $SiteDatum = New-Object PSObject
        $SiteDatum | Add-Member NoteProperty Area($Area)
        $SiteDatum | Add-Member NoteProperty SiteId($SiteId)
        $SiteDatum | Add-Member NoteProperty Url($SiteUrl)
        $SiteDatum | Add-Member NoteProperty GroupId($SPOSite.GroupId)
        $SiteDatum | Add-Member NoteProperty Storage($SPOSite.Usage.Storage)
        $SiteDatum | Add-Member NoteProperty RootWeb($SPOSite.RootWeb.Url)
        $SiteDatum | Add-Member NoteProperty SiteOwner($SPOSite.Owner.LoginName -replace "i:0#\.f\|membership\|", "")
        $SiteDatum | Add-Member NoteProperty SharingCapability($SharingCapability)
        $SiteDatum | Add-Member NoteProperty GroupVisibility($group.Visibility)
        $SiteDatum | Add-Member NoteProperty HasTeam($group.HasTeam)
        $SiteDatum | Add-Member NoteProperty ResourceProvisioningOptions($group.ResourceProvisioningOptions -join ",")
        $SiteDatum | Add-Member NoteProperty SiteSensitivityLabel($SPOSite.SensitivityLabel.Name)
        $SiteDatum | Add-Member NoteProperty GroupOwners(($group.Owners.UserPrincipalName -join ",") -replace "i:0#\.f\|membership\|", "")
        $SiteDatum | Add-Member NoteProperty IsRestrictedAccessControlPolicyEnforcedOnSite($IsRestrictedAccessControlPolicyEnforcedOnSite)
        $SiteDatum | Add-Member NoteProperty IsRestrictContentOrgWideSearchPolicyEnforcedOnSite($IsRestrictContentOrgWideSearchPolicyEnforcedOnSite)
        $SiteDatum | Add-Member NoteProperty DisableCompanyWideSharingLinks($DisableCompanyWideSharingLinks)
        $SiteData += $SiteDatum
        if ($GetWebPermissions -or 
            $GetListPermissions -or 
            $GetItemPermissions -or
            $GetWebDetails -or
            $GetWebPermissions) {
            foreach ($Web in $Webs) {
                $Web = Get-PnPWeb
                Get-SPOWebDetails -Web $Web `
                    -ReportOutputPath $ReportOutputPath `
                    -GetListDetails:$GetListDetails `
                    -GetListPermissions:$GetListPermissions `
                    -IncludeSystemLists:$IncludeSystemLists `
                    -GetItemPermissions:$GetItemPermissions `
                    -GetWebPermissions:$GetWebPermissions `
                    -GetItemDetails:$GetItemDetails `
                    -SiteId $SiteId
            } #foreach web
        } #if web permissions, list permissions, itempermissions, webdetails, or webpermissions
    }
    END {
        $SiteData | Export-Csv -Path $ReportOutput -NoTypeInformation -Append
        Disconnect-PnPOnline
    } #end
}
Function get-SPODetails {
    Param(
        $ReportOutputPath,
        $ClientId, #App Only Registration
        $CertificatePath, #App Only Registration
        $Tenant,
        $SPOAdminUrl,
        [Switch]$GetListPermissions,
        [Switch]$GetItemPermissions,
        [switch]$GetListDetails,
        [switch]$IncludeSystemLists,
        [switch]$GetItemDetails,
        [switch]$GetWebPermissions,
        [switch]$ClearPriorLogs
    )
    BEGIN {
        if ($ClearPriorLogs) {
            Remove-Item "$ReportOutputPath\*.txt"
            Remove-Item "$ReportOutputPath\*.csv"
            Clear-Host
        }
        Connect-PnPOnline -Url $SPOAdminUrl `
            -ClientId $ClientId `
            -Tenant $Tenant `
            -CertificatePath $CertificatePath
        $SPOSites=Get-PnPTenantSite
    } #begin
PROCESS {
    foreach ($SPOSite in $SPOSites) {
            Write-Host "Processing Site: $($SPOSite.Url)"
            Get-SPOSiteDetails -SiteUrl "$($SPOSite.Url)" `
                -SharingCapability $SPOSite.SharingCapability `
                -ClientId $ClientId `
                -Tenant $Tenant `
                -CertificatePath $CertificatePath `
                -ReportOutputPath $ReportOutputPath `
                -GetWebDetails:$GetWebDetails `
                -GetWebPermissions:$GetWebPermissions `
                -GetListDetails:$GetListDetails `
                -IncludeSystemLists:$IncludeSystemLists `
                -GetListPermissions:$GetListPermissions `
                -GetItemPermissions:$GetItemPermissions `
                -GetItemDetails:$GetItemDetails
        } #foreach sposite
    } #process
}

#Runs the full scritp with default params#Runs the full script with default params
#Update the parameters with <> to reflect your environment
Get-SPODetails -ReportOutputPath "c:\temp\spinventory" `
    -ClientId "<Your Entra App Client Id>" `
    -CertificatePath "<Path to your pfx file>" `
    -Tenant "<Your Tenant Name>.onmicrosoft.com" `
    -SPOAdminUrl "https://<Your Tenant Name>-admin.sharepoint.com" `
    -GetWebDetails `
    -GetWebPermissions `
    -GetListDetails `
    -GetListPermissions `
    -IncludeSystemLists:$false `
    -GetItemPermissions `
    -GetItemDetails `
    -ClearPriorLogs

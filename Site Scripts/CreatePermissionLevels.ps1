[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string] $Site,

    [ValidateSet("CalTF", "FutEE")]
    [string] $Tenant = "CalTF"
)

. ..\SharePoint.ps1

function New-PermissionLevels {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $Site,
    
        [ValidateSet("CalTF", "FutEE")]
        [string] $Tenant = "CalTF"
    )

    Connect-SharePoint -SiteName $Site -Tenant $Tenant
    
    Write-Host "Creating role definitions..."
    
    $BasicMembershipExclusions = @(
        "EditListItems"
        "DeleteListItems"
        "DeleteVersions"
        "BrowseDirectories"
        "EditMyUserInfo"
        "ManagePersonalViews"
        "AddDelPrivateWebParts"
        "UpdatePersonalWebParts"
    )
    $BasicMembershipInfo = @{
        "RoleName"  = "Basic Membership"
        "Clone"     = "Contribute"
        "Exclude"   = $BasicMembershipExclusions
    }
    try {
        New-PermissionLevel @BasicMembershipInfo
        Write-Host "Successfully created Basic Membership"
    } catch {
        Write-Error "Unable to create Basic Membership: $($_.Exception.Message)"
        exit
    }
    
    $BasicListPermissionsExclusions = @(
        "DeleteListItems"
        "ManagePersonalViews"
        "AddDelPrivateWebParts"
        "UpdatePersonalWebParts"
    )
    $BasicListPermissionsInfo = @{
        "RoleName"  = "Basic List Permissions"
        "Clone"     = "Edit"
        "Exclude"   = $BasicListPermissionsExclusions
    }
    try {
        New-PermissionLevel @BasicListPermissionsInfo
        Write-Host "Successfully created Basic List Permissions"
    } catch {
        Write-Error "Unable to create Basic List Permissions: $($_.Exception.Message)"
        exit
    }
    
    $AdvancedMembershipExclusions = @(
        "ManageLists"
        "EditListItems"
        "DeleteVersions"
        "BrowseDirectories"
        "CreateSSCSite"
        "EditMyUserInfo"
        "ManagePersonalViews"
        "AddDelPrivateWebParts"
        "UpdatePersonalWebParts"
    )
    $AdvancedMembershipInfo = @{
        "RoleName"  = "Advanced Membership"
        "Clone"     = "Edit"
        "Exclude"   = $AdvancedMembershipExclusions
    }
    try {
        New-PermissionLevel @AdvancedMembershipInfo
        Write-Host "Successfully created Advanced Membership"
    } catch {
        Write-Error "Unable to create Advanced Membership: $($_.Exception.Message)"
        exit
    }
}

New-PermissionLevels -Site $Site -Tenant $Tenant

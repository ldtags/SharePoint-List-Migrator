. .\SharePoint.ps1


function Set-GroupPermissions {
    <#
    .SYNOPSIS
        Sets the default permission structure of a SharePoint site.

    .DESCRIPTION
        Set-Permissions sets the FutEE default permission structure
        for any site on the FutEE or CalTF domains.

        This structure refers to the following structure:
            Group Owners        - Site Owners
            Advanced Members    - FutEE Employees
            Basic Members       - External Users

        All existing document libraries and lists will also be modified
        to grant the custom groups accurate list permissions.

        SharePoint Admin level access is required.

    .PARAMETER SiteName
        The name of the site whose permissions will be set.

    .PARAMETER Tenant
        The tenant that the SharePoint site is hosted on.

    .NOTES
        Author: Liam D Tangney
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $SiteName,

        [ValidateNotNullOrEmpty()]
        [ValidateSet("FutEE", "CalTF")]
        [string] $Tenant = "CalTF"
    )

    $DomainMap = @{
        "FutEE" = "futeeenergy"
        "CalTF" = "californiatechnicalforum"
    }

    try {
        $Domain = $DomainMap[$Tenant]
    } catch {
        Write-Error "Invalid tenant: $($Tenant)"
        return
    }

    $AdminUrl = "https://$Domain-admin.sharepoint.com"
    try {
        Connect-SharePoint -Url $AdminUrl
    } catch {
        Write-Error "Unable to connect as an admin: $($_.Exception.Message)"
        return
    }

    $SiteUrl = "https://$Domain.sharepoint.com/sites/$SiteName"
    try {
        Connect-SharePoint -Url $SiteUrl
    } catch {
        Write-Error "Unable to connect to $($SiteUrl): $($_.Exception.Message)"
        return
    }

    Write-Host "Creating role definitions..."

    Set-PermissionLevel `
        -RoleName "Basic Membership" `
        -Clone "Contribute" `
        -Exclude EditListItems, DeleteListItems, DeleteVersions, BrowseDirectories, EditMyUserInfo, ManagePersonalViews, AddDelPrivateWebParts, UpdatePersonalWebParts

    Set-PermissionLevel `
        -RoleName "Basic List Permissions" `
        -Clone "Edit" `
        -Exclude DeleteListItems, ManagePersonalViews, AddDelPrivateWebParts, UpdatePersonalWebParts

    Set-PermissionLevel `
        -RoleName "Advanced Membership" `
        -Clone "Edit" `
        -Exclude ManageLists, EditListItems, DeleteVersions, BrowseDirectories, CreateSSCSite, EditMyUserInfo, ManagePersonalViews, AddDelPrivateWebParts, UpdatePersonalWebParts

    $OwnerGroup = Get-PnPGroup -AssociatedOwnerGroup
    if ($null -ne $OwnerGroup) {
        $Owner = $OwnerGroup.Title
    } else {
        $Owner = $null
    }

    Set-Group `
        -GroupName "Basic Members - $SiteName" `
        -Owner $Owner `
        -PermissionsLevel "Basic Membership"

    Set-Group `
        -GroupName "Advanced Members - $SiteName" `
        -Owner $Owner `
        -PermissionsLevel "Advanced Membership"
}


function Set-ListPermissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string[]] $Lists,

        [Parameter(Mandatory)]
        [string] $SiteName,

        [ValidateNotNullOrEmpty()]
        [ValidateSet("FutEE", "CalTF")]
        [string] $Tenant = "CalTF"
    )

    if ("" -ne $SiteName) {
        $DomainMap = @{
            "FutEE" = "futeeenergy"
            "CalTF" = "californiatechnicalforum"
        }
    
        try {
            $Domain = $DomainMap[$Tenant]
        } catch {
            Write-Error "Invalid tenant: $($Tenant)"
            return
        }

        Connect-SharePoint -Url "https://$($Domain).sharepoint.com/sites/$($SiteName)"
    } else {
        $Context = Get-PnPContext
        if ($null -eq $Context) {
            throw [ContextException]::new(
                "No PnP context found.",
                "No existing PnP context found and no site url was provided."
            )
        }
    }

    $CurrentUser = (Get-PnPProperty -ClientObject (Get-PnPWeb) -Property CurrentUser)
    $Lists | ForEach-Object {
        # Reset list permissions to default
        Set-PnPList -Identity $_ -ResetRoleInheritance

        $BasePermissions = @{}
        $List = $_
        Get-PnPGroup | ForEach-Object {
            $Role = Get-PnPListPermissions -Identity $List -PrincipalId $_.Id
            $BasePermissions[$_.Title] = $Role.Name
        }

        # Disable list permission inheritence to add custom list permissions
        Set-PnPList -Identity $_ -BreakRoleInheritance

        # Add base permissions (removed during inheritence break)
        foreach ($Group in $BasePermissions.Keys) {
            Set-PnPListPermission `
                -Identity $_ `
                -Group $Group `
                -AddRole $BasePermissions[$Group][0]
        }

        # Set custom list permissions for basic members
        Set-PnPListPermission `
            -Identity $_ `
            -Group "Basic Members - $SiteName" `
            -AddRole "Basic List Permissions"

        # Set custom list permissions for advanced members
        Set-PnPListPermission `
            -Identity $_ `
            -Group "Advanced Members - $SiteName" `
            -AddRole "Edit"

        # Remove leftover permissions from inheritence break
        Set-PnPListPermission -Identity $_ -User $CurrentUser.Email -RemoveRole "Full Control"
    }
}


# Set-GroupPermissions -SiteName "CMUA"
Set-ListPermissions -SiteName "CMUA" -Lists "Documents"

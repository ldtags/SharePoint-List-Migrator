. .\CustomExceptions.ps1

function Connect-SharePoint {
    [CmdletBinding()]
    param (
        [string] $Url,

        [string] $SiteName,

        [ValidateSet("CalTF", "FutEE", "")]
        [string] $Tenant
    )

    if ("" -ne $Url) {
        $SanitizedUrl = $Url
        if (-not $Url.StartsWith("https://")) {
            $SanitizedUrl = "https://$($SanitizedUrl)"
        }

        $UrlRegex = "^https://[a-zA-Z]+[a-zA-Z0-9]*((?:-admin.sharepoint.com(/)?)|(?:.sharepoint.com/sites/[a-zA-Z]+[a-zA-Z0-9]*(/)?))$"
        if (-not ($SanitizedUrl -match $UrlRegex)) {
            throw [ContextException]::new(
                "Invalid site URL: $($Url)",
                "The site url must be either a SharePoint site url or a SharePoint tenant admin url."
            )
        }
    } elseif (("" -ne $SiteName) -and ("" -ne $Tenant)) {
        switch ($Tenant) {
            "FutEE" {
                $Domain = "futeeenergy"
                break
            }
            "CalTF" {
                $Domain = "californiatechnicalforum"
            }
            default {
                throw [Exception]::new("Unrecognized tenant: $($Tenant)")
            }
        }
        $SanitizedUrl = "https://$($Domain).sharepoint.com/sites/$($SiteName)"
    } else {
        throw [Exception]::new(
            "Usage: Connect-SharePoint [-Url <string> | [-SiteName <string> -Tenant <string>]]"
        )
    }

    try {
        Connect-PnPOnline -Url $SanitizedUrl -ClientId "83ea45f3-5f26-423a-b32b-bc6f64e26b7d" -Interactive
    } catch {
        Write-Error "Unable to connect to $($SanitizedUrl): $($_.Exception.Message)"
        throw
    }
}

function Set-Group {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $GroupName,

        [Parameter(Mandatory)]
        [string] $PermissionsLevel,

        [string] $SiteUrl,

        [string] $Owner,

        [string] $Description
    )

    if ("" -ne $SiteUrl) {
        Connect-SharePoint -Url $SiteUrl
    } else {
        $Context = Get-PnPContext
        if ($null -eq $Context) {
            throw [ContextException]::new(
                "No PnP context found.",
                "No existing PnP context found and no site url was provided."
            )
        }
    }

    $Group = $null
    try {
        $Group = Get-PnPGroup -Identity $GroupName
    } catch {
        New-PnPGroup -Title $GroupName
    }

    if ($null -ne $Group) {
        Get-PnPGroupPermissions -Identity $GroupName | ForEach-Object {
            Set-PnPGroup -Identity $GroupName -RemoveRole $_.Name
        }
    }

    Set-PnPGroup `
        -Identity $GroupName `
        -Owner $Owner `
        -Description $Description `
        -AddRole $PermissionsLevel
}

function New-PermissionLevel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $RoleName,

        [string] $SiteUrl,

        [string] $Clone,

        [Microsoft.SharePoint.Client.PermissionKind[]] $Exclude
    )

    if ("" -ne $SiteUrl) {
        Connect-SharePoint -Url $SiteUrl
    } else {
        $Ctx = Get-PnPContext
        if ($null -eq $Ctx) {
            throw [ContextException]::new(
                "No PnP context found.",
                "No existing PnP context found and no site url was provided."
            )
        }
    }

    try {
        Get-PnPRoleDefinition -Identity $RoleName
        Write-Host "$RoleName already exists."
    } catch {
        try {
            Add-PnPRoleDefinition `
                -RoleName $RoleName `
                -Clone $Clone `
                -Exclude $Exclude
        } catch {
            Write-Error "Unable to create $($RoleName): $($_.Exception.Message)"
            throw
        }
    } 
}

function Get-EntryExistence {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [hashtable] $Table,

        [Parameter(Mandatory)]
        [string] $Key,

        [Parameter(Mandatory)]
        $Value
    )

    if (-not ($Table.Keys.Contains($Key))) {
        return $false
    }

    if ($Table[$Key] -ne $Value) {
        return $false
    }

    return $true
}

function Set-ListPermissions {
    <#
    .SYNOPSIS
        Sets permissions for specified lists on the site.

    .DESCRIPTION
        Set-ListPermissions sets custom permissions for lists on a SharePoint
        site.

        All base permissions are included by default and should be removed
        with a different method.

        This method breaks list permission inheritence.

    .PARAMETER Lists
        A collection of list names that will have their permissions updated.

    .PARAMETER SiteName
        The name of the site whose permissions will be set.

    .PARAMETER Tenant
        The tenant that the SharePoint site is hosted on.

    .PARAMETER CustomPermissions
        A hashtable mapping group names to the role they will be granted for
        each list.

    .NOTES
        Author: Liam D Tangney
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string[]] $Lists,

        [Parameter(Mandatory)]
        [hashtable] $CustomPermissions,

        [string] $SiteUrl
    )

    if ("" -ne $SiteUrl) {
        Connect-SharePoint -Url $SiteUrl
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
    foreach ($ListName in $Lists) {
        try {
            Get-PnPList -Identity $ListName -ThrowExceptionIfListNotFound
        } catch {
            Write-Host "No list named $($ListName) exists."
            continue
        }

        # Reset list permissions to default
        Set-PnPList -Identity $ListName -ResetRoleInheritance

        # Store the default list permissions in a map
        $DefaultPermissions = @{}
        Get-PnPGroup | ForEach-Object {
            $Role = Get-PnPListPermissions -Identity $ListName -PrincipalId $_.Id
            $DefaultPermissions[$_.Title] = $Role.Name
        }

        # Disable list permission inheritence to add custom list permissions
        Set-PnPList -Identity $ListName -BreakRoleInheritance

        # Add base permissions (removed during inheritence break)
        foreach ($Group in $DefaultPermissions.Keys) {
            Set-PnPListPermission `
                -Identity $ListName `
                -Group $Group `
                -AddRole $DefaultPermissions[$Group][0]
        }

        # Add custom permissions
        foreach ($Group in $CustomPermissions.Keys) {
            Set-PnPListPermission `
                -Identity $ListName `
                -Group $Group `
                -AddRole $CustomPermissions[$Group][0]
        }

        # Remove leftover permissions from inheritence break
        $DefPermCheckInfo = @{
            "Table" = $DefaultPermissions
            "Key"   = $CurrentUser.Email
            "Value" = "Full Control"
        }
        $CustomPermCheckInfo = @{
            "Table" = $CustomPermissions
            "Key"   = $CurrentUser.Email
            "Value" = "Full Control"
        }
        if (-not ((Get-EntryExistence @DefPermCheckInfo) -or (Get-EntryExistence @CustomPermCheckInfo))) {
            $PermissionsInfo = @{
                "Identity"      = $ListName
                "User"          = $CurrentUser.Email
                "RemoveRole"    = "Full Control"
            }
            Set-PnPListPermission @PermissionsInfo
        }
    }
}

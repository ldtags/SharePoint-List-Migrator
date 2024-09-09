. .\CustomExceptions.ps1


function Connect-SharePoint {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $Url
    )

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


function Set-PermissionLevel {
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
        Write-Host "$RoleName successfully created."
    } 
}

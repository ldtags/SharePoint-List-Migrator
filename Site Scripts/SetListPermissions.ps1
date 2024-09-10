[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string[]] $Lists,

    [string] $Site,

    [ValidateSet("CalTF", "FutEE")]
    [string] $Tenant = "CalTF"
)

. ..\SharePoint.ps1

try {
    $Context = Get-PnPContext
} catch {
    $Context = $null
}

if ("" -eq $Site) {
    if ($null -eq $Context) {
        Write-Error "No established PnP connection exists"
        exit
    } elseif ($Context.Url -match "^https://[a-zA-Z0-9]+-admin.sharepoint.com/$") {
        Write-Error "Established PnP connection is not a site connection"
        exit
    }
} else {
    Connect-SharePoint -SiteName $Site -Tenant $Tenant
}

$Context = Get-PnPContext
$SiteName = $Context.Url.Split("/")[-1]
$CustomPermissions = @{
    "Basic Members - $SiteName" = "Basic List Permissions"
    "Advanced Members - $SiteName" = "Edit"
}
Set-ListPermissions -Lists @Lists -CustomPermissions $CustomPermissions

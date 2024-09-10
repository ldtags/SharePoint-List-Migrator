[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string] $Site,

    [ValidateSet("CalTF", "FutEE")]
    [string] $Tenant = "CalTF"
)

. ..\SharePoint.ps1

function New-Groups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $Site,
    
        [ValidateSet("CalTF", "FutEE")]
        [string] $Tenant = "CalTF"
    )

    Connect-SharePoint -SiteName $Site -Tenant $Tenant

    $OwnerGroup = Get-PnPGroup -AssociatedOwnerGroup
    if ($null -ne $OwnerGroup) {
        $Owner = $OwnerGroup.Title
    } else {
        $CurrentUser = (Get-PnPProperty -ClientObject (Get-PnPWeb) -Property CurrentUser)
        $Owner = $CurrentUser.Email
    }
    
    $BasicMembersInfo = @{
        "GroupName"         = "Basic Members - $Site"
        "Owner"             = $Owner
        "PermissionsLevel"  = "Basic Membership"
    }
    try {
        Set-Group @BasicMembersInfo
    } catch {
        Write-Error "Unable to create Basic Members group: $($_.Exception.Message)"
        exit
    }
    
    $AdvancedMembersInfo = @{
        "GroupName"         = "Advanced Members - $Site"
        "Owner"             = $Owner
        "PermissionsLevel"  = "Advanced Membership"
    }
    try {
        Set-Group @AdvancedMembersInfo
    } catch {
        Write-Error "Unable to create Advanced Members group: $($_.Exception.Message)"
        exit
    }
}

New-Groups -Site $Site -Tenant $Tenant

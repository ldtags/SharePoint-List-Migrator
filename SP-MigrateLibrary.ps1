function Copy-SPOListItemAttachments {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $SourceUrl,

        [Parameter(Mandatory)]
        [string] $SourceListName,

        [Parameter(Mandatory)]
        [string] $SourceItemId,

        [Parameter(Mandatory)]
        [string] $DestinationUrl,

        [Parameter(Mandatory)]
        [string] $DestinationListName,

        [Parameter(Mandatory)]
        [string] $DestinationItemId
    )

    try {
        Connect-PnPOnline -Url $SourceUrl -Interactive
    } catch {
        Write-Error "Could not establish connection to $($SourceUrl): $($_.Exception.Message)"
        throw
    }

    try {
        $SourceList = Get-PnPList -Identity $SourceListName
    } catch {
        Write-Error "Could not get the list $($SourceListName): $($_.Exception.Message)"
        throw
    }

    try {
        $SourceItem = Get-PnPListItem -List $SourceList -Id $SourceItemId
    } catch {
        Write-Error "Could not get the list item $($SourceItemId): $($_.Exception.Message)"
        throw
    }

    try {
        $Attachments = Get-PnPProperty -ClientObject $SourceItem -Property "AttachmentFiles"
    } catch {
        Write-Error "Could not get attachments: $($_.Exception.Message)"
        throw
    }

    try {
        Connect-PnPOnline -Url $DestinationUrl -Interactive
        $DestCtx = Get-PnPConnection
    } catch {
        Write-Error "Could not establish connection to $($DestinationUrl): $($_.Exception.Message)"
        throw
    }

    try {
        $DestinationList = Get-PnPList -Identity $DestinationListName
    } catch {
        Write-Error "Could not get the list $($DestinationListName): $($_.Exception.Message)"
        throw
    }

    try {
        $DestinationItem = Get-PnPListItem -List $DestinationList -Id $DestinationItemId
    } catch {
        Write-Error "Could not get the list item $($DestinationItemId): $($_.Exception.Message)"
        throw
    }

    $Attachments | ForEach-Object {
        try {
            Connect-PnPOnline -Url $SourceUrl -Interactive
        } catch {
            Write-Error "Could not establish connection to $($SourceUrl): $($_.Exception.Message)"
            throw
        }

        # Download the attachment to TEMP
        try {
            Get-PnPFile -Url $_.ServerRelativeUrl -FileName $_.FileName -Path $Env:TEMP -AsFile -Force
        } catch {
            Write-Error "Could not retrieve the attachment: $($_.Exception.Message)"
            throw
        }

        try {
            # Add attachment to destination list item
            $FileStream = New-Object IO.FileStream(($Env:TEMP + "\" + $_.FileName), [System.IO.FileMode]::Open)
            $AttachmentInfo = New-Object -TypeName Microsoft.SharePoint.Client.AttachmentCreationInformation
            $AttachmentInfo.FileName = $_.FileName
            $AttachmentInfo.ContentStream = $FileStream
            $DestinationItem.AttachmentFiles.Add($AttachmentInfo)
            Invoke-PnPQuery -Connection $DestCtx
        } catch {
            Write-Error "Could not copy attachment: $($_.Exception.Message)"
            throw
        } finally {
            Remove-Item -Path $Env:TEMP\$($_.FileName) -Force
        }
    }
}

function Copy-SPOListItems {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $SourceListName,

        [Parameter(Mandatory)]
        [string] $DestinationListName,

        [Parameter(Mandatory)]
        [string] $SourceUrl,

        [Parameter(Mandatory)]
        [string] $DestinationUrl
    )

    try {
        Connect-PnPOnline -Url $SourceUrl -Interactive
    } catch {
        Write-Error "Could not establish connection to $($SourceUrl): $($_.Exception.Message)"
        throw
    }

    try {
        Connect-PnPOnline -Url $DestinationUrl -Interactive
    } catch {
        Write-Error "Could not establish connection to $($DestinationUrl): $($_.Exception.Message)"
        throw
    }

    Write-Progress -Activity "Reading source..." -Status "Getting items from the source list. Please wait..."
    try {
        Connect-PnPOnline -Url $SourceUrl -Interactive
    } catch {
        Write-Error "Could not establish connection to $($SourceUrl): $($_.Exception.Message)"
        throw
    }

    try {
        $SourceList = Get-PnPList -Identity $SourceListName -ThrowExceptionIfListNotFound
    } catch {
        Write-Error "Could not get list $($SourceListName): $($_.Exception.Message)"
        throw
    }

    try {
        $ListItems = Get-PnPListItem -List $SourceList
    } catch {
        Write-Error "Could not get list items: $($_.Exception.Message)"
        throw
    }

    $ListItemsCount = $ListItems.Count
    Write-Host "Total number of items found: $ListItemsCount"

    try {
        # Get fields from the source list
        # Field types skipped: Read Only, Hidden, Content Type, Attachments
        $SourceListFields = Get-PnPField -List $SourceList -Connection $SourceConnection | Where-Object { `
            (-not ($_.ReadOnlyField)) `
            -and (-not ($_.Hidden)) `
            -and ($_.InternalName -ne "ContentType") `
            -and ($_.InternalName -ne "Attachments")
        }
    } catch {
        Write-Error "Could not get list fields: $($_.Exception.Message)"
        throw
    }

    [int] $Counter = 1
    foreach ($ListItem in $ListItems) {
        $ItemValue = @{}
        foreach ($Field in $SourceListFields) {
            if ($null -ne $ListItem[$Field.InternalName]) {
                $FieldType = $Field.TypeAsString
                $FieldName = $Field.InternalName
                $FieldValue = $ListItem[$FieldName]
                switch -Regex ($FieldType) {
                    "User|UserMulti" {
                        $PeoplePickerValues = $FieldValue | ForEach-Object { $_.Email }
                        $ItemValue.Add($FieldName, $PeoplePickerValues)
                        break
                    }
                    "Lookup|LookupMulti" {
                        $LookupIds = $FieldValue | ForEach-Object { $_.LookupID.ToString() }
                        $ItemValue.Add($FieldName, $LookupIds)
                        break
                    }
                    "URL" {
                        $Url = $FieldValue.URL
                        $Description = $FieldValue.Description
                        $ItemValue.Add($FieldName, "$Url, $Description")
                        break
                    }
                    "TaxonomyFieldType|TaxonomyFieldTypeMulti" {
                        $TermGuids = $FieldValue | ForEach-Object { $_.TermGuid.ToString() }
                        $ItemValue.Add($FieldName, $TermGuids)
                        break
                    }
                    default {
                        $ItemValue.Add($FieldName, $FieldValue)
                        break
                    }
                }
            }
        }

        $ItemValue.Add("Created", $ListItem["Created"])
        $ItemValue.Add("Modified", $ListItem["Modified"])
        $ItemValue.Add("Author", $ListItem["Author"].Email)
        $ItemValue.Add("Editor", $ListItem["Editor"].Email)

        Write-Progress `
            -Activity "Copying list items:" `
            -Status "Copying item ID '$($ListItem.Id)' from source list [$($Counter) of $($ListItemsCount)]" `
            -PercentComplete (($Counter / $SourceListItemsCount) * 100)

        try {
            Connect-PnPOnline -Url $DestinationUrl -Interactive
        } catch {
            Write-Error "Could not establish connection to $($DestinationUrl): $($_.Exception.Message)"
            throw
        }

        try {
            $DestinationList = Get-PnPList -Identity $DestinationListName -ThrowExceptionIfListNotFound
        } catch {
            Write-Error "Could not get list $($DestinationListName): $($_.Exception.Message)"
            throw
        }


        try {
            # Copy the item and any attachments
            $NewItem = Add-PnPListItem -List $DestinationList -Values $ItemValue
        } catch {
            Write-Error "Could not copy the list item: $($_.Exception.Message)"
            throw
        }

        try {
            Copy-SPOListItemAttachments `
                -SourceUrl $SourceUrl `
                -SourceListName $SourceListName `
                -SourceItemId $ListItem.Id `
                -DestinationUrl $DestinationUrl `
                -DestinationListName $DestinationListName `
                -DestinationItemId $NewItem.Id

            $Counter++
        } catch {
            Write-Error "Could not copy attachments: $($_.Exception.Message)"
            throw
        }
    }
}

function Copy-SPOLibraryItems {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $SourceListName,

        [Parameter(Mandatory)]
        [string] $DestinationListName,

        [Parameter(Mandatory)]
        [string] $SourceUrl,

        [Parameter(Mandatory)]
        [string] $DestinationUrl
    )

    try {
        Connect-PnPOnline -Url $DestinationUrl -Interactive
        $DestCtx = Get-PnPContext
    } catch {
        Write-Error "Could not establish connection to $($DestinationUrl): $($_.Exception.Message)"
        throw
    }

    try {
        $DestinationLibrary = Get-PnPList $DestinationListName -Includes RootFolder
    } catch {
        Write-Error "Could not get the list $($DestinationListName): $($_.Exception.Message)"
        throw
    }

    try {
        Connect-PnPOnline -Url $SourceUrl -Interactive
        $SrcCtx = Get-PnPContext
    } catch {
        Write-Error "Could not establish connection to $($SourceUrl): $($_.Exception.Message)"
        throw
    }

    try {
        $SourceLibrary = Get-PnPList $SourceListName -Includes RootFolder
    } catch {
        Write-Error "Could not get the list $($SourceListName): $($_.Exception.Message)"
        throw
    }

    $global:counter = 0
    $ListItems = Get-PnPListItem `
        -List $SourceListName `
        -PageSize 500 `
        -Fields ID `
        -ScriptBlock {
            param (
                $Items
            ) $global:counter += $Items.Count;
            Write-Progress `
                -PercentComplete (($global:counter / $SourceLibrary.ItemCount) * 100) `
                -Activity "Getting items from $($SourceLibrary.Title)" `
                -Status "Getting item $($global:counter) of $($SourceLibrary.ItemCount)"
        }

    $RootFolderItems = $ListItems | Where-Object {
        $_.FieldValues.FileRef.Replace(
            ("/" + $_.FieldValues.FileLeafRef),
            [string]::Empty
        ) -eq $SourceLibrary.RootFolder.ServerRelativeUrl
    }

    Write-Progress `
        -Activity "Completed getting items from $($SourceLibrary.Title)" `
        -Completed

    Write-Host "Copying files from $($SourceUrl) to $($DestinationUrl)"

    # copy all files to the destination site
    $global:counter = 1
    $DestinationLibraryUrl = $DestinationLibrary.RootFolder.ServerRelativeUrl
    $RootFolderItems | ForEach-Object {
        Write-Progress `
            -PercentComplete (($global:counter / $RootFolderItems.Count) * 100) `
            -Activity "Copying files from $($SourceUrl) to $($DestinationUrl)" `
            -Status "Copying file $($global:counter) of $($RootFolderItems.Count)"

        $File = $_
        try {
            Copy-PnPFile `
                -SourceUrl "$($_.FieldValues['FileRef']) $($_.Title)" `
                -TargetUrl $DestinationLibraryUrl `
                -Force `
                -OverwriteIfAlreadyExists
        } catch {
            Write-Error "Could not copy file $($File.Id): $($_.Exception.Message)"
        }

        $global:counter++
    }

    Write-Progress `
        -Activity "Completed copying files from $($SourceLibrary.Title)" `
        -Completed

    try {
        Set-PnPContext -Context $DestCtx
    } catch {
        Write-Error "Could not switch contexts: $($_.Exception.Message)"
        throw
    }

    try {
        $DestinationItems = Get-PnPListItem `
            -List $DestinationListName `
            -PageSize 500
    } catch {
        Write-Error "Could not retrieve items from $($DestinationLibrary.Title): $($_.Exception.Message)"
        throw
    }

    Write-Host "Copying metadata from $($SourceLibrary.Title) to $($DestinationLibrary.Title)"

    # copy metadata for all files
    $global:counter = 1
    Set-PnPContext -Context $SrcCtx
    foreach ($ListItem in $ListItems) {
        Write-Progress `
            -PercentComplete (($global:counter / $ListItems.Count) * 100) `
            -Activity "Copying metadata from $($SourceUrl) to $($DestinationUrl)" `
            -Status "File $($global:counter) of $($ListItems.Count)"

        $Metadata = @{
            "Title" = $ListItem.FieldValues.Title
            "Created" = $ListItem.FieldValues.Created.DateTime
            "Modified" = $ListItem.FieldValues.Modified.DateTime
            "Author" = $ListItem.FieldValues.Author.Email
            "Editor" = $ListItem.FieldValues.Editor.Email
        }

        Set-PnPContext -Context $DestCtx
        $DestLibUrl = $DestinationLibrary.RootFolder.ServerRelativeUrl

        Set-PnPContext -Context $SrcCtx
        $SourceLibUrl = $SourceLibrary.RootFolder.ServerRelativeUrl
        $DestRelativeUrl = $ListItem.FieldValues.FileLeafRef.Replace($SourceLibUrl, $DestLibUrl)

        Set-PnPContext -Context $DestCtx
        $MatchingItem = $DestinationItems | Where-Object {
            $_.FieldValues.FileLeafRef -eq $DestRelativeUrl
        }

        Set-PnPListItem `
            -List $DestinationListName `
            -Identity $MatchingItem.Id `
            -Values $Metadata `
            | Out-Null

        $global:counter++
    }
    Write-Progress `
        -Activity "Completed copying metadata from $($SourceLibrary.Title)" `
        -Completed
}

function Copy-SPOList {
    <#
    .SYNOPSIS
        Moves a document library from one SharePoint site to another.

    .DESCRIPTION
        Move-SPLibrary is a function that accepts a document library hosted
        on a SharePoint site and moves it to another SharePoint site. Either
        SharePoint site can be a subsite of another site.

    .PARAMETER SourceSite
        The name of the site that the document library being migrated
        is currently stored on.

    .PARAMETER ListName
        The name of the SharePoint list being migrated.

    .PARAMETER DestinationSite
        The name of the site that the document library being migrated
        will be migrated to.
        If this site is a subsite, enter the relative path to the subsite
        starting from the top-level site.

    .PARAMETER Tenant
        [Optional] The name of the tenant that the SharePoint site is
        hosted on. This defaults to the California Technical Forum domain.

    .NOTES
        Author: Liam D Tangney
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $SourceSite,

        [Parameter(Mandatory)]
        [string] $ListName,

        [Parameter(Mandatory)]
        [string] $DestinationSite,

        [ValidateNotNullOrEmpty()]
        [ValidateSet("List", "Document Library")]
        [string] $ListType = "List",

        [ValidateNotNullOrEmpty()]
        [string] $Tenant = "californiatechnicalforum"
    )

    $AdminUrl = "https://$Tenant-admin.sharepoint.com"
    try {
        Connect-PnPOnline -Url $AdminUrl -Interactive
    } catch {
        Write-Error "Could not establish an admin-level connection: $($_.Exception.Message)"
        return
    }

    $SourceUrl = "https://$Tenant.sharepoint.com/sites/$SourceSite"
    try {
        Connect-PnPOnline -Url $SourceUrl -Interactive
    } catch {
        Write-Error "Could not connect to $($SourceUrl): $($_.Exception.Message)"
        return
    }

    $DestinationUrl = "https://$Tenant.sharepoint.com/sites/$DestinationSite"
    try {
        Connect-PnPOnline -Url $DestinationUrl -Interactive   
    } catch {
        Write-Error "Could not connect to $($DestinationUrl): $($_.Exception.Message)"
        return
    }

    $ExistingList = Get-PnPList -Identity $ListName
    if ($null -ne $ExistingList) {
        $Prompt = @(
            "A library named $($ListName) already exists on the site"
            "$($DestinationSite). Would you like to delete the existing"
            "library? [y|n]"
        ) -join " "
        $Response = ""
        while (($Response.ToLower() -ne "y") -and ($Response.ToLower() -ne "n")) {
            $Response = Read-Host $Prompt
        }
        switch ($Response.ToLower()) {
            "y" {
                try {
                    Remove-PnPList -Identity $ListName -Force
                } catch {
                    Write-Error "Could not delete $($ListName): $($_.Exception.Message)"
                    return
                }
                break
            }
            "n" {
                return
            }
            default {
                Write-Error "Invalid user input: $($Response)"
                return
            }
        }
    }

    try {
        Connect-PnPOnline -Url $SourceUrl -Interactive
    } catch {
        Write-Error "Could not connect to $($SourceUrl): $($_.Exception.Message)"
        return
    }

    try {
        # Copies the list structure from the source site to the destination site
        Copy-PnPList -Identity $ListName -Title $ListName -DestinationWebUrl $DestinationUrl
    } catch {
        Write-Error "Could not copy $($ListName) from $($SourceSite) to $($DestinationSite): $($_.Exception.Message)"
        return
    }

    try {
        switch ($ListType) {
            "List" {
                Copy-SPOListItems `
                    -SourceUrl $SourceUrl `
                    -SourceListName $ListName `
                    -DestinationUrl $DestinationUrl `
                    -DestinationListName $ListName
            }
            "Document Library" {
                Copy-SPOLibraryItems `
                    -SourceUrl $SourceUrl `
                    -SourceListName $ListName `
                    -DestinationUrl $DestinationUrl `
                    -DestinationListName $ListName
            }
        }
    } catch {
        Write-Error $_.Exception.Message
        return
    }

    Write-Host "Succesfully copied $ListName from $SourceSite to $DestinationSite"
}

Copy-SPOList -SourceSite 'CalTFManagement' -ListName 'Recruiting' -DestinationSite 'CalTFManagement/Leadership' -ListType 'Document Library'

﻿function Add-SPListItem {
    param(
        [string]$ListUrl,
        [string]$Title,
        [DateTime]$Date
    )

    # Encode the title for use in the REST API URL
    $EncodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)

    # Construct the REST API URL for adding the item
    $AddItemUrl = "$ListUrl/_api/lists/getbytitle('List Name')/items"

    # Get the X-RequestDigest value for authentication
    $FormDigestUrl = "$ListUrl/_api/contextinfo"
    $FormDigest = Invoke-RestMethod -Uri $FormDigestUrl -Method Post -ContentType "application/json;odata=verbose" -Headers @{
        "Accept" = "application/json;odata=verbose"
    } -UseDefaultCredentials
    $RequestDigest = $FormDigest.d.GetContextWebInformation.FormDigestValue

    # Construct the JSON payload for the new item
    $NewItemPayload = @{
        "__metadata" = @{
            "type" = "SP.Data.ListNameListItem"
        }
        "Title" = $Title
        "Date" = $Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    } | ConvertTo-Json -Depth 4

    # Send the REST API request to add the item
    $Response = Invoke-RestMethod -Uri $AddItemUrl -Method Post -Body $NewItemPayload -ContentType "application/json;odata=verbose" -Headers @{
        "X-RequestDigest" = $RequestDigest
    } -UseDefaultCredentials

    # Return the new item's ID
    return $Response.d.Id
}

function Remove-AllSPListItems {
    param(
        [string]$ListUrl,
        [string]$ListName
    )

    # Construct the REST API URL for getting all items in the list
    $GetItemsUrl = "$ListUrl/_api/lists/getbytitle('$ListName')/items"

    # Get the X-RequestDigest value from SharePoint
    $ContextInfo = Get-SPContextInfo $ListUrl
    $RequestDigest = $ContextInfo.GetContextWebInformation.FormDigestValue

    # Call the REST API to get all items in the list
    $Items = Invoke-RestMethod -Uri $GetItemsUrl -Method Get -ContentType "application/json;odata=verbose" -Headers @{ "X-RequestDigest" = $RequestDigest } -UseDefaultCredentials

    # Loop through all items and delete them one by one
    foreach ($Item in $Items.d.results) {
        $DeleteUrl = "$ListUrl/_api/lists/getbytitle('$ListName')/items($($Item.Id))"
        Invoke-RestMethod -Uri $DeleteUrl -Method Delete -ContentType "application/json;odata=verbose" -Headers @{ "X-RequestDigest" = $RequestDigest } -UseDefaultCredentials
        Write-Host "Deleted item with ID $($Item.Id)" -ForegroundColor Green
    }
}

function Get-SPListItemDateValue {
    param (
        [string]$SiteUrl,
        [string]$ListName,
        [string]$DateColumnName,
        [string]$Username,
        [string]$Password
    )

    # Set up the credentials to access the SharePoint site
    $secpasswd = ConvertTo-SecureString $Password -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential ($Username, $secpasswd)

    # Set up the REST API endpoint to retrieve all items in the list
    $endpointUrl = "$SiteUrl/_api/web/lists/getbytitle('$ListName')/items"

    # Send a GET request to the endpoint and retrieve the response
    $headers = @{
        "Accept" = "application/json;odata=verbose"
    }
    $response = Invoke-RestMethod -Uri $endpointUrl -Credential $cred -Headers $headers -Method Get

    # Loop through each item in the response and output the value in the date column,
    # along with the item ID and title
    foreach ($item in $response.d.results) {
        $dateValue = $item[$DateColumnName]
        $itemId = $item.Id
        $title = $item.Title
        Write-Output "Item ID: $itemId, Title: $title, Date Value: $dateValue"
    }
}

function Update-SPListItemDate {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true)]
        [string]$ListName,
        [Parameter(Mandatory=$true)]
        [string]$Title,
        [Parameter(Mandatory=$true)]
        [string]$DateColumnName,
        [Parameter(Mandatory=$true)]
        [datetime]$NewDate,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential
    )
    $EndpointUrl = "$SiteUrl/_api/web/lists/getbytitle('$ListName')/items?\$filter=Title eq '$Title'"
    $Header = @{
        'Authorization' = 'Bearer ' + $Credential.GetNetworkCredential().Password
        'Accept' = 'application/json;odata=verbose'
        'Content-Type' = 'application/json;odata=verbose'
    }
    $items = Invoke-RestMethod -Uri $EndpointUrl -Headers $Header -Method Get
    foreach ($item in $items.d.results) {
        $itemId = $item.Id
        $itemUrl = $item.__metadata.uri
        $itemType = $item.__metadata.type
        $itemETag = $item.__metadata.etag
        $body = @{
            '__metadata' = @{
                'type' = $itemType
            }
            $DateColumnName = $NewDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        } | ConvertTo-Json -Depth 1
        Invoke-RestMethod -Uri $itemUrl -Headers $Header -Method Merge -Body $body -ContentType "application/json;odata=verbose" -IfMatch $itemETag
        Write-Output "Item ID: $itemId, Title: $Title, Date Column ($DateColumnName) updated to: $NewDate"
    }
}

function Update-SPListItemDate {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true)]
        [string]$ListName,
        [Parameter(Mandatory=$true)]
        [string]$Title,
        [Parameter(Mandatory=$true)]
        [string]$DateColumnName,
        [Parameter(Mandatory=$true)]
        [datetime]$NewDate,
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential
    )
    $EndpointUrl = "$SiteUrl/_api/web/lists/getbytitle('$ListName')/items?\$filter=Title eq '$Title'"
    $Header = @{
        'Authorization' = 'Bearer ' + $Credential.GetNetworkCredential().Password
        'Accept' = 'application/json;odata=verbose'
        'Content-Type' = 'application/json;odata=verbose'
    }
    $items = Invoke-RestMethod -Uri $EndpointUrl -Headers $Header -Method Get
    foreach ($item in $items.d.results) {
        $itemId = $item.Id
        $itemUrl = $item.__metadata.uri
        $itemType = $item.__metadata.type
        $itemETag = $item.__metadata.etag
        $header = @{
            'Authorization' = 'Bearer ' + $Credential.GetNetworkCredential().Password
            'Accept' = 'application/json;odata=verbose'
            'If-Match' = $itemETag
            'X-HTTP-Method' = 'MERGE'
            'Content-Type' = 'application/json;odata=verbose'
        }
        $body = @{
            '__metadata' = @{
                'type' = $itemType
            }
            $DateColumnName = $NewDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        } | ConvertTo-Json -Depth 1
        Invoke-RestMethod -Uri $itemUrl -Headers $header -Method Post -Body $body -ContentType "application/json;odata=verbose"
        Write-Output "Item ID: $itemId, Title: $Title, Date Column ($DateColumnName) updated to: $NewDate"
    }
}

function New-SPDocumentSet {
    param (
        [Parameter(Mandatory=$true)][string]$SiteUrl,
        [Parameter(Mandatory=$true)][string]$ListName,
        [Parameter(Mandatory=$true)][string]$DocumentSetName,
        [Parameter(Mandatory=$true)][string]$KeyBlah,
        [Parameter(Mandatory=$true)][string]$DealBlah,
        [Parameter(Mandatory=$true)][string]$ContentTypeId,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$UserAgent = "PowerShell"
    )

    # Get the list details
    $listUrl = "$SiteUrl/_api/web/lists/getbytitle('$ListName')"
    $listResponse = Invoke-RestMethod -Uri $listUrl -Headers @{ "Accept" = "application/json; odata=verbose" } -Credential $Credential -Method Get
    $listId = $listResponse.d.Id

    # Get the parent folder for the document set
    $folderUrl = "$SiteUrl/_api/web/lists(guid'$listId')/rootfolder"
    $folderResponse = Invoke-RestMethod -Uri $folderUrl -Headers @{ "Accept" = "application/json; odata=verbose" } -Credential $Credential -Method Get
    $folderServerRelativeUrl = $folderResponse.d.ServerRelativeUrl

    # Get the form digest value
    $formDigestValue = (Get-SPFormDigest -SiteUrl $SiteUrl -Credential $Credential).d.GetContextWebInformation.FormDigestValue

    # Create the document set
    $createDocSetUrl = "$SiteUrl/_api/web/GetFolderByServerRelativeUrl('$folderServerRelativeUrl')/files/AddUsingPath(DecodedUrl='$DocumentSetName',overwrite=true)"
    $headers = @{
        "Accept" = "application/json; odata=verbose";
        "X-RequestDigest" = $formDigestValue;
        "Slug" = "$folderServerRelativeUrl/$DocumentSetName|$ContentTypeId";
        "User-Agent" = $UserAgent
    }

    $createDocSetResponse = Invoke-RestMethod -Uri $createDocSetUrl -Headers $headers -Credential $Credential -Method Post

    # Set the properties for the document set
    $itemUrl = "$SiteUrl/_api/web/lists(guid'$listId')/items(" + $createDocSetResponse.d.Id + ")"
    $itemPayload = @{
        "__metadata" = @{
            "type" = $createDocSetResponse.d['__metadata'].type
        }
        "KeyBlah" = $KeyBlah
        "DealBlah" = $DealBlah
    } | ConvertTo-Json

    $headers = @{
        "Accept" = "application/json; odata=verbose";
        "X-RequestDigest" = $formDigestValue;
        "Content-Type" = "application/json; odata=verbose";
        "IF-MATCH" = "*";
        "X-HTTP-Method" = "MERGE";
        "User-Agent" = $UserAgent
    }

    Invoke-RestMethod -Uri $itemUrl -Headers $headers -Credential $Credential -Method Post -Body $itemPayload
    Write-Host "Document set '$DocumentSetName' created and properties updated successfully." -ForegroundColor Green
}

function Remove-SPDocumentSet {
    param (
        [Parameter(Mandatory=$true)][string]$SiteUrl,
        [Parameter(Mandatory=$true)][string]$ListName,
        [Parameter(Mandatory=$true)][string]$DocumentSetName,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$UserAgent = "PowerShell"
    )

    # Get the list details
    $listUrl = "$SiteUrl/_api/web/lists/getbytitle('$ListName')"
    $listResponse = Invoke-RestMethod -Uri $listUrl -Headers @{ "Accept" = "application/json; odata=verbose" } -Credential $Credential -Method Get
    $listId = $listResponse.d.Id

    # Get the document set details
    $docSetUrl = "$SiteUrl/_api/web/lists(guid'$listId')/items?`$filter=ContentTypeId%20ne%20null%20and%20Title%20eq%20'$DocumentSetName'"
    $docSetResponse = Invoke-RestMethod -Uri $docSetUrl -Headers @{ "Accept" = "application/json; odata=verbose" } -Credential $Credential -Method Get

    if ($docSetResponse.d.results.Count -eq 0) {
        Write-Host "Document set not found." -ForegroundColor Red
        return
    }

    $docSetId = $docSetResponse.d.results[0].Id

    # Get the form digest value
    $formDigestValue = (Get-SPFormDigest -SiteUrl $SiteUrl -Credential $Credential).d.GetContextWebInformation.FormDigestValue

    # Delete the document set
    $deleteDocSetUrl = "$SiteUrl/_api/web/lists(guid'$listId')/items($docSetId)"
    $headers = @{
        "Accept" = "application/json; odata=verbose";
        "X-RequestDigest" = $formDigestValue;
        "IF-MATCH" = "*";
        "X-HTTP-Method" = "DELETE";
        "User-Agent" = $UserAgent
    }

    Invoke-RestMethod -Uri $deleteDocSetUrl -Headers $headers -Credential $Credential -Method Post
    Write-Host "Document set '$DocumentSetName' deleted successfully." -ForegroundColor Green
}

function Get-SPFormDigest {
    param (
        [Parameter(Mandatory=$true)][string]$SiteUrl,
        [string]$AccessToken
    )

    $formDigestUrl = "$SiteUrl/_api/contextinfo"

    if ($null -eq $AccessToken) {
        $AccessToken = Get-SPAccessToken -SiteUrl $SiteUrl
    }

    $headers = @{
        "Accept" = "application/json; odata=verbose";
        "Authorization" = "Bearer $AccessToken"
    }

    $response = Invoke-RestMethod -Uri $formDigestUrl -Headers $headers -Method Post
    return $response
}

function New-SPDocumentSet {
    param (
        [Parameter(Mandatory=$true)][string]$SiteUrl,
        [Parameter(Mandatory=$true)][string]$ListName,
        [Parameter(Mandatory=$true)][string]$DocumentSetName,
        [Parameter(Mandatory=$true)][string]$KeyBlah,
        [Parameter(Mandatory=$true)][string]$DealBlah,
        [Parameter(Mandatory=$true)][string]$ContentTypeId,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$UserAgent = "PowerShell"
    )

    # Get the form digest value
    $formDigestValue = (Get-SPFormDigest -SiteUrl $SiteUrl -Credential $Credential).d.GetContextWebInformation.FormDigestValue

    # Create the document set
    $createDocSetUrl = "$SiteUrl/_vti_bin/listdata.svc/$ListName"
    $headers = @{
        "Accept" = "application/json";
        "X-RequestDigest" = $formDigestValue;
        "Content-Type" = "application/json";
        "User-Agent" = $UserAgent
    }

    $body = @{
        "ContentTypeID" = $ContentTypeId
        "Path" = "/$ListName/$DocumentSetName"
        "KeyBlah" = $KeyBlah
        "DealBlah" = $DealBlah
    } | ConvertTo-Json

    $createDocSetResponse = Invoke-RestMethod -Uri $createDocSetUrl -Headers $headers -Credential $Credential -Method Post -Body $body

    Write-Host "Document set '$DocumentSetName' created with content type '$ContentTypeId' and properties updated successfully." -ForegroundColor Green
}

function New-SPDocumentSet {
    param (
        [Parameter(Mandatory=$true)][string]$SiteUrl,
        [Parameter(Mandatory=$true)][string]$ListName,
        [Parameter(Mandatory=$true)][string]$DocumentSetName,
        [Parameter(Mandatory=$true)][string]$KeyBlah,
        [Parameter(Mandatory=$true)][string]$DealBlah,
        [Parameter(Mandatory=$true)][string]$ContentTypeId,
        [System.Management.Automation.PSCredential]$Credential
    )

    # Load SharePoint CSOM Assemblies
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.DocumentManagement.dll"

    # Client Context
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.UserName, $Credential.Password)

    # Get the list and content type
    $list = $ctx.Web.Lists.GetByTitle($ListName)
    $ctx.Load($list)
    $ctx.Load($list.RootFolder)
    $ctx.ExecuteQuery()

    $contentType = $ctx.Web.ContentTypes.GetById($ContentTypeId)
    $ctx.Load($contentType)
    $ctx.ExecuteQuery()

    # Create document set
    $docSetCreateInfo = New-Object Microsoft.SharePoint.Client.DocumentSet.DocumentSetCreateInfo
    $docSetCreateInfo.Path = $list.RootFolder.ServerRelativeUrl
    $docSetCreateInfo.Title = $DocumentSetName
    $docSetCreateInfo.DocumentSetTemplateId = $contentType.Id

    $newDocSet = [Microsoft.SharePoint.Client.DocumentSet.DocumentSet]::Create($ctx, $list.RootFolder, $docSetCreateInfo)
    $ctx.Load($newDocSet)
    $ctx.ExecuteQuery()

    # Update properties
    $newDocSet.ListItemAllFields["KeyBlah"] = $KeyBlah
    $newDocSet.ListItemAllFields["DealBlah"] = $DealBlah
    $newDocSet.ListItemAllFields.Update()
    $ctx.ExecuteQuery()

    Write-Host "Document set '$DocumentSetName' created with content type '$ContentTypeId' and properties updated successfully." -ForegroundColor Green
}

function New-SPDocumentSet {
    param (
        [Parameter(Mandatory=$true)][string]$SiteUrl,
        [Parameter(Mandatory=$true)][string]$ListName,
        [Parameter(Mandatory=$true)][string]$DocumentSetName,
        [Parameter(Mandatory=$true)][string]$KeyBlah,
        [Parameter(Mandatory=$true)][string]$DealBlah,
        [Parameter(Mandatory=$true)][string]$ContentTypeId,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$UserAgent = "PowerShell"
    )

    # Custom WebRequestCreator for setting the UserAgent
    Add-Type -TypeDefinition @"
        using System;
        using System.Net;
        public class CustomWebRequestCreator : IWebRequestCreate
        {
            public WebRequest Create(Uri uri)
            {
                HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(uri);
                webRequest.UserAgent = "$UserAgent";
                return webRequest;
            }
        }
"@

    # Register custom WebRequestCreator
    [System.Net.WebRequest]::RegisterPrefix("http://", [CustomWebRequestCreator]::new())
    [System.Net.WebRequest]::RegisterPrefix("https://", [CustomWebRequestCreator]::new())

    # Load SharePoint CSOM Assemblies
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

    # Client Context
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $ctx.Credentials = New-Object Microsoft.SharePoint.Client.NetworkCredential($Credential.UserName, $Credential.Password, $Credential.Domain)

    # Get the list and content type
    $list = $ctx.Web.Lists.GetByTitle($ListName)
    $ctx.Load($list)
    $ctx.Load($list.RootFolder)
    $ctx.ExecuteQuery()

    $contentType = $ctx.Web.ContentTypes.GetById($ContentTypeId)
    $ctx.Load($contentType)
    $ctx.ExecuteQuery()

    # Create document set
    $docSetFolder = [Microsoft.SharePoint.Client.Folder]::CreateFolderDirect($ctx, $list.RootFolder, $DocumentSetName)
    $ctx.Load($docSetFolder)
    $ctx.ExecuteQuery()

    # Update properties and content type
    $docSetItem = $docSetFolder.ListItemAllFields
    $docSetItem["ContentTypeId"] = $ContentTypeId
    $docSetItem["KeyBlah"] = $KeyBlah
    $docSetItem["DealBlah"] = $DealBlah
    $docSetItem.Update()
    $ctx.ExecuteQuery()

    Write-Host "Document set '$DocumentSetName' created with content type '$ContentTypeId' and properties updated successfully." -ForegroundColor Green
}

function New-SPDocumentSet {
    param (
        [Parameter(Mandatory=$true)][string]$SiteUrl,
        [Parameter(Mandatory=$true)][string]$ListName,
        [Parameter(Mandatory=$true)][string]$DocumentSetName,
        [Parameter(Mandatory=$true)][string]$KeyBlah,
        [Parameter(Mandatory=$true)][string]$DealBlah,
        [Parameter(Mandatory=$true)][string]$ContentTypeId,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$UserAgent = "PowerShell"
    )

    $FormDigestValue = Get-SPFormDigest -SiteUrl $SiteUrl -Credential $Credential

    $ListUrl = $SiteUrl + "/_api/web/lists/GetByTitle('" + $ListName + "')"
    $ListItemsUrl = $SiteUrl + "/_api/web/lists/GetByTitle('" + $ListName + "')/items"

    $Headers = @{
        "Accept" = "application/json;odata=verbose"
        "Content-Type" = "application/json;odata=verbose"
        "X-RequestDigest" = $FormDigestValue
        "UserAgent" = $UserAgent
    }

    $DocumentSetFolderUrl = $SiteUrl + "/_api/web/folders/AddUsingPath(decodedurl='" + $ListName + "/" + $DocumentSetName + "',overwrite=true,url='')"
    $Response = Invoke-WebRequest -Uri $DocumentSetFolderUrl -Method Post -Headers $Headers -Credential $Credential

    $FolderId = (ConvertFrom-Json $Response.Content).d.Id
    $UpdateUrl = $SiteUrl + "/_api/web/lists/GetByTitle('" + $ListName + "')/items(" + $FolderId + ")"
    $UpdatePayload = @{
        "__metadata" = @{
            "type" = "SP.Data." + $ListName + "ListItem"
        }
        "ContentTypeId" = $ContentTypeId
        "KeyBlah" = $KeyBlah
        "DealBlah" = $DealBlah
    } | ConvertTo-Json

    Invoke-WebRequest -Uri $UpdateUrl -Method Post -Headers $Headers -Body $UpdatePayload -Credential $Credential
}
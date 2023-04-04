function Add-SPListItem {
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
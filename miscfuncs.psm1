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
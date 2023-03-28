function Add-SPListItem {
    param(
        [string]$ListUrl,
        [string]$Title,
        [DateTime]$Date
    )

    # Encode the title for use in the REST API URL
    $EncodedTitle = [System.Web.HttpUtility]::UrlEncode($Title)

    # Construct the REST API URL for adding the item
    $AddItemUrl = "$ListUrl/items"

    # Construct the JSON payload for the new item
    $NewItemPayload = @{
        "__metadata" = @{
            "type" = "SP.Data.ListNameListItem"
        }
        "Title" = $Title
        "Date" = $Date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    } | ConvertTo-Json -Depth 4

    # Send the REST API request to add the item
    $Response = Invoke-RestMethod -Uri $AddItemUrl -Method Post -Body $NewItemPayload -ContentType "application/json;odata=verbose"

    # Return the new item's ID
    return $Response.d.Id
}

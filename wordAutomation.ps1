# URL of a document in SharePoint (ensure it's a valid document URL)
$docUrl = "https://yoursharepointsite/sites/YourSite/Shared Documents/Example.docx"

# Path to the Office application (Word in this case)
$wordPath = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"

# Check if Word exists at the specified path
if (-Not (Test-Path $wordPath)) {
    Write-Host "Error: Word application not found at $wordPath. Check the path and try again." -ForegroundColor Red
    exit 1
}

# Open the SharePoint document in Word
Write-Host "Opening document in Word to establish authentication and persist the FedAuth cookie..."
Start-Process -FilePath $wordPath -ArgumentList $docUrl

# Instructions for the user
Write-Host "The document has been opened in Word. Authenticate if prompted. Once authenticated, the cookie will be persisted."
Write-Host "After authentication, you can close Word and test Explorer View in File Explorer."

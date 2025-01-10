# Load the SharePoint PowerShell snap-in
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Variables
$backupFolder = "C:\SharePointWebConfigBackups"  # Path where backups will be stored
$webAppsToBackup = @(
    "http://webapp1",    # Replace with your web application URLs
    "http://webapp2"
)

# Create the backup folder if it doesn't exist
if (-not (Test-Path -Path $backupFolder)) {
    New-Item -ItemType Directory -Path $backupFolder | Out-Null
}

# Backup process
foreach ($webAppUrl in $webAppsToBackup) {
    try {
        # Get the web application object
        $webApp = Get-SPWebApplication -Identity $webAppUrl

        # Filter out empty or null keys
        $validZones = $webApp.IisSettings.Keys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        # Iterate through valid zones
        foreach ($zone in $validZones) {
            $iisSettings = $webApp.IisSettings[$zone]

            # Get the path to the web.config file for this zone
            $webConfigPath = Join-Path $iisSettings.Path.FullName "web.config"

            if (Test-Path -Path $webConfigPath) {
                # Create a backup file name with the zone included
                $backupFileName = "$($webApp.Url -replace '[^a-zA-Z0-9]', '_')_$zone_web.config.backup"
                $backupFilePath = Join-Path $backupFolder $backupFileName

                # Copy the web.config file to the backup folder
                Copy-Item -Path $webConfigPath -Destination $backupFilePath -Force
                Write-Host "Backup created for $webAppUrl ($zone) at $backupFilePath" -ForegroundColor Green
            } else {
                Write-Warning "web.config not found for $webAppUrl ($zone) at $webConfigPath"
            }
        }
    } catch {
        Write-Error "Failed to back up web.config for $webAppUrl. Error: $_"
    }
}

Write-Host "Backup process completed. Backups are stored in $backupFolder" -ForegroundColor Cyan

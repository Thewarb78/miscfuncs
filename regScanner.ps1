# Specify the search term
$searchTerm = "your_search_string"

# Specify the registry root keys to search
$rootKeys = @("HKLM:\", "HKCU:\", "HKCR:\", "HKU:\", "HKCC:\")

# Loop through each root key and search
foreach ($rootKey in $rootKeys) {
    Write-Host "Searching in $rootKey..."
    try {
        Get-ChildItem -Path $rootKey -Recurse -ErrorAction SilentlyContinue |
        ForEach-Object {
            # Search for the key name
            if ($_ -match $searchTerm) {
                Write-Host "Found in Key: $($_.PSPath)"
            }

            # Search for values within the key
            Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue |
            ForEach-Object {
                $_.PSObject.Properties | ForEach-Object {
                    if ($_.Value -match $searchTerm) {
                        Write-Host "Found in Key: $($_.PSPath) - Value: $($_.Name) = $($_.Value)"
                    }
                }
            }
        }
    } catch {
        Write-Warning "Error accessing $rootKey"
    }
}

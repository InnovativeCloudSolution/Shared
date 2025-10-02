
# Function to load a JSON payload from a file
function Load-JSONPayload {
    param (
        [string]$filePath
    )

    if (Test-Path $filePath) {
        return Get-Content -Path $filePath -Raw
    } else {
        Write-Error "File not found: $filePath"
    }
}

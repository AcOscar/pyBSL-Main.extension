# Load secrets directly from client.env (im gleichen Ordner wie dieses Script)
$dotenvPath = Join-Path -Path $PSScriptRoot -ChildPath "client.env"
if (Test-Path $dotenvPath) {
    # Lese alle Zeilen ein und filtere Kommentare
    $secretLines = Get-Content $dotenvPath | Where-Object { $_ -and $_ -notmatch '^\s*#' }
    # Erstelle Hashtable für Secrets
    $secrets = @{}
    foreach ($line in $secretLines) {
        $parts = $line -split '=', 2
        $key   = $parts[0].Trim()
        $value = $parts[1].Trim()
        $secrets[$key] = $value
    }

    # Weise ClientId und ClientSecret direkt zu
    $ClientId     = $secrets['CLIENT_ID']
    $ClientSecret = $secrets['CLIENT_SECRET']
} else {
    Write-Error "Secrets-Datei '$dotenvPath' nicht gefunden."
}

param (
    [string]$ProjectGuid,
    [string]$ItemUrn
)

$ProjectId = "b." + $ProjectGuid

#Write-Output "Token wird geholt..."

$body = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
    scope         = "data:read"
}

try {
    $tokenResponse = Invoke-RestMethod `
        -Uri "https://developer.api.autodesk.com/authentication/v2/token" `
        -Method POST `
        -Body $body

    $accessToken = $tokenResponse.access_token

    #if (-not $accessToken) {
    #    Write-Output "Kein Zugriffstoken erhalten."
    #    exit 1
    #}

    #Write-Output "Zugriffstoken empfangen."
}
catch {
    #Write-Output "Fehler beim Tokenabruf:"
    #Write-Output $_
    exit 1
}

# Modellversionen abfragen
#Write-Output "Hole Modellversionen..."

$headers = @{ Authorization = "Bearer $accessToken" }
$url = "https://developer.api.autodesk.com/data/v1/projects/$ProjectId/items/$ItemUrn/versions"

try {
    $versions = Invoke-RestMethod -Uri $url -Headers $headers -Method GET
    $latest = $versions.data[0]

    if ($null -eq $latest) {
        #Write-Output "Keine Versionen gefunden."
        exit 1
    }

    $size = $latest.attributes.storageSize
    if ($size) {
        $sizeMB = ($size)
        Write-Output "$size"
    }
    else {
        #Write-Output "Keine Dateigröße verfügbar (in Version)."
        Write-Output "-"
    }
}
catch {
    Write-Output "Fehler beim Datenabruf:"
    Write-Output $_
}

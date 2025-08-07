param (
    [string]$ProjectGuid,
    [string]$ItemUrn
)

$ClientId = "QXGVf1neHIouY41kH8o0TVGAPIPbQebbulB371893AD6AMqE"
$ClientSecret = "eNAFwASrEs6mneWCY2GcVwIgmYhFnPX2RB3wbjvrWsp2ClwSx3gaW7DmsO9jpip4"
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

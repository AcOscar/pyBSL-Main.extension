# Load secrets directly from client.env (same folder as this script)
# the user.env has content like:
# CLIENT_ID=AKNS4zhkILkwpt62m56v0w1agw2kgKc4amKrl5
# CLIENT_SECRET=Gh6KliQsfZbkuWjQoPD5WcNMaf267dAcILkwpt62m56v0wamtdlw3

$dotenvPath = Join-Path -Path $PSScriptRoot -ChildPath "client.env"

if (Test-Path $dotenvPath) {
    # read all rows and filter comments
    $secretLines = Get-Content $dotenvPath | Where-Object { $_ -and $_ -notmatch '^\s*#' }
    # creating hash table for secrets
    $secrets = @{}
    foreach ($line in $secretLines) {
        $parts = $line -split '=', 2
        $key   = $parts[0].Trim()
        $value = $parts[1].Trim()
        $secrets[$key] = $value
    }

    # Directly assign ClientId and ClientSecret
    $ClientId     = $secrets['CLIENT_ID']
    $ClientSecret = $secrets['CLIENT_SECRET']
} else {
    Write-Error "Error: Secrets-file '$dotenvPath' not found."
}

$ProjectGuid = $args[0]  # argument 1
$ItemUrn = $args[1]      # argument 2

$ProjectId = "b." + $ProjectGuid

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

    if (-not $accessToken) {
       Write-Output "Error: No access token received."
       exit 1
    }

}
catch {
    Write-Output "Error: Retrieving token:"
    Write-Output $_
    exit 1
}

$headers = @{ Authorization = "Bearer $accessToken" }
$url = "https://developer.api.autodesk.com/data/v1/projects/$ProjectId/items/$ItemUrn/versions"

try {
    $versions = Invoke-RestMethod -Uri $url -Headers $headers -Method GET
    $latest = $versions.data[0]

    if ($null -eq $latest) {
        exit 1
    }

    $size = $latest.attributes.storageSize
    if ($size) {
        Write-Output "$size"
    }
    else {
        Write-Output "-"
    }
}
catch {
    Write-Output "Error: during data calling:"
    Write-Output $_
}

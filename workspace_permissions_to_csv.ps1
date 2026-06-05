param (
    [switch]$ExcludePersonalWorkspaces,
    [string]$CapacityId
)

$ErrorActionPreference = "Continue"

$outputFile = "./workspace_users.csv"
$errorFile  = "./workspace_errors.csv"

# ------------------------------------------------------------
# Authenticate to Microsoft Entra and get Fabric API token
# ------------------------------------------------------------
try {
    Import-Module Az.Accounts -ErrorAction Stop
}
catch {
    Write-Error "Az.Accounts module is required. Install-Module Az.Accounts"
    throw
}

# Disable token cache persistence to avoid MSAL assembly version conflicts
# (may fail on some environments with SafeConfigManager - that's OK)
Disable-AzContextAutosave -Scope Process -ErrorAction SilentlyContinue 2>$null | Out-Null

try {
    Connect-AzAccount -UseDeviceAuthentication -ErrorAction Stop | Out-Null
}
catch {
    # If Disable-AzContextAutosave caused SafeConfigManager issues, retry without it
    if ($_.Exception.Message -match "SafeConfigManager|safe mode") {
        Write-Host "Retrying authentication without context autosave change..."
        # Clear the broken context state and retry
        try {
            Clear-AzContext -Force -ErrorAction SilentlyContinue | Out-Null
        } catch {}
        Connect-AzAccount -UseDeviceAuthentication -ErrorAction Stop | Out-Null
    }
    else {
        Write-Error "Failed to authenticate with Azure."
        throw
    }
}

try {
    $tokenResult = Get-AzAccessToken -ResourceUrl "https://api.fabric.microsoft.com" -ErrorAction Stop
    # Handle both plain string (older Az.Accounts) and SecureString (newer Az.Accounts)
    if ($tokenResult.Token -is [System.Security.SecureString]) {
        $accessToken = [System.Net.NetworkCredential]::new('', $tokenResult.Token).Password
    } else {
        $accessToken = $tokenResult.Token
    }
}
catch {
    Write-Error "Failed to acquire Fabric API access token."
    throw
}

$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type"  = "application/json"
}

# Token refresh helper - refreshes token if older than 45 minutes
$script:tokenAcquiredTime = Get-Date

function Refresh-FabricToken {
    $elapsed = (Get-Date) - $script:tokenAcquiredTime
    if ($elapsed.TotalMinutes -ge 45) {
        Write-Host "Access token nearing expiry. Refreshing..."
        try {
            $tokenResult = Get-AzAccessToken -ResourceUrl "https://api.fabric.microsoft.com" -ErrorAction Stop
            if ($tokenResult.Token -is [System.Security.SecureString]) {
                $script:accessToken = [System.Net.NetworkCredential]::new('', $tokenResult.Token).Password
            } else {
                $script:accessToken = $tokenResult.Token
            }
            $script:headers = @{
                "Authorization" = "Bearer $script:accessToken"
                "Content-Type"  = "application/json"
            }
            $script:tokenAcquiredTime = Get-Date
            Write-Host "Token refreshed successfully."
        }
        catch {
            Write-Warning "Failed to refresh token: $($_.Exception.Message)"
        }
    }
}

# ------------------------------------------------------------
# Get all workspaces from Fabric Admin API (with pagination)
# ------------------------------------------------------------
$workspaces = [System.Collections.Generic.List[object]]::new()
$continuationToken = $null
$baseUrl = "https://api.fabric.microsoft.com/v1/admin/workspaces"

do {
    try {
        $url = $baseUrl

        if ($continuationToken) {
            $encoded = [System.Web.HttpUtility]::UrlEncode($continuationToken)
            $url = "$baseUrl?continuationToken=$encoded"
        }

        $response = Invoke-RestMethod `
            -Method Get `
            -Uri $url `
            -Headers $headers `
            -ErrorAction Stop

        if ($response.workspaces) {
            foreach ($ws in $response.workspaces) {
                $workspaces.Add($ws)
            }
        }

        $continuationToken = $response.continuationToken
    }
    catch {
        Write-Error "Failed to retrieve workspaces from Fabric Admin API."
        throw
    }
}
while ($continuationToken)

# Optionally exclude personal workspaces
if ($ExcludePersonalWorkspaces) {
    $workspaces = $workspaces | Where-Object { $_.type -ne "Personal" }
}

# Optionally filter by Fabric capacity ID
if ($CapacityId) {
    $workspaces = $workspaces | Where-Object { $_.capacityId -eq $CapacityId }
    Write-Host "Filtered to capacity $CapacityId : $($workspaces.Count) workspaces."
}

if (-not $workspaces -or $workspaces.Count -eq 0) {
    Write-Warning "No workspaces found."
    return
}

Write-Host "Retrieved $($workspaces.Count) workspaces."

# ------------------------------------------------------------
# Collect results
# ------------------------------------------------------------
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$errorResults = [System.Collections.Generic.List[PSCustomObject]]::new()

$totalCount = $workspaces.Count
$currentIndex = 0

foreach ($workspace in $workspaces) {

    $currentIndex++

    # Refresh token if nearing expiry
    Refresh-FabricToken

    $workspaceId   = $workspace.id
    $workspaceName = $workspace.name
    $workspaceType = $workspace.type
    $workspaceState = $workspace.state

    Write-Progress `
        -Activity "Fetching workspace access details" `
        -Status "$currentIndex / $totalCount : $workspaceName" `
        -PercentComplete (($currentIndex / $totalCount) * 100)

    try {
        # Skip personal workspaces if requested
        if ($ExcludePersonalWorkspaces -and $workspaceType -eq "Personal") {
            continue
        }

        $usersUrl = "https://api.fabric.microsoft.com/v1/admin/workspaces/$workspaceId/users"

        # Retry logic for rate limiting (HTTP 429)
        $maxRetries = 5
        $retryCount = 0
        $userResponse = $null

        while ($retryCount -le $maxRetries) {
            try {
                $userResponse = Invoke-RestMethod `
                    -Method Get `
                    -Uri $usersUrl `
                    -Headers $headers `
                    -ErrorAction Stop
                break
            }
            catch {
                $statusCode = $null
                if ($_.Exception.Response) {
                    $statusCode = [int]$_.Exception.Response.StatusCode
                }

                if ($statusCode -eq 401 -and $retryCount -lt $maxRetries) {
                    $retryCount++
                    Write-Host "  Token expired (401). Refreshing and retrying ($retryCount/$maxRetries)..."
                    Refresh-FabricToken
                }
                elseif ($statusCode -eq 429 -and $retryCount -lt $maxRetries) {
                    $retryCount++
                    # Parse Retry-After header or use exponential backoff
                    $retryAfter = $null
                    if ($_.Exception.Response.Headers) {
                        $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                    }
                    if ($retryAfter) {
                        $waitSeconds = [int]$retryAfter
                    } else {
                        $waitSeconds = [math]::Pow(2, $retryCount) * 10
                    }
                    Write-Host "  Rate limited. Waiting $waitSeconds seconds before retry ($retryCount/$maxRetries)..."
                    Start-Sleep -Seconds $waitSeconds
                }
                else {
                    throw
                }
            }
        }

        if ($userResponse.accessDetails) {
            foreach ($entry in $userResponse.accessDetails) {

                $principal = $entry.principal
                $access    = $entry.workspaceAccessDetails

                $userPrincipalName = $null
                $groupType = $null

                if ($principal.userDetails) {
                    $userPrincipalName = $principal.userDetails.userPrincipalName
                }

                if ($principal.groupDetails) {
                    $groupType = $principal.groupDetails.groupType
                }

                $results.Add(
                    [PSCustomObject]@{
                        "Workspace Name"    = $workspaceName
                        "Workspace ID"      = $workspaceId
                        "Workspace Type"    = $workspaceType
                        "Workspace State"   = $workspaceState
                        "Principal Name"    = $principal.displayName
                        "Principal ID"      = $principal.id
                        "Principal Type"    = $principal.type
                        "User PrincipalName"= $userPrincipalName
                        "Group Type"        = $groupType
                        "Access Type"       = $access.type
                        "Workspace Role"    = $access.workspaceRole
                    }
                )
            }
        }
        else {
            # 空応答も見えるように残す
            $results.Add(
                [PSCustomObject]@{
                    "Workspace Name"    = $workspaceName
                    "Workspace ID"      = $workspaceId
                    "Workspace Type"    = $workspaceType
                    "Workspace State"   = $workspaceState
                    "Principal Name"    = ""
                    "Principal ID"      = ""
                    "Principal Type"    = ""
                    "User PrincipalName"= ""
                    "Group Type"        = ""
                    "Access Type"       = ""
                    "Workspace Role"    = ""
                }
            )
        }

        # Admin API はレート制限があるためリクエスト間隔を空ける
        Start-Sleep -Seconds 5
    }
    catch {
        $errorMessage = $_.Exception.Message
        $errorDetail  = $_ | Format-List -Force | Out-String

        Write-Warning "Failed to get users for workspace '$workspaceName' ($workspaceId)"
        Write-Warning $errorMessage

        $errorResults.Add(
            [PSCustomObject]@{
                "Workspace Name" = $workspaceName
                "Workspace ID"   = $workspaceId
                "Workspace Type" = $workspaceType
                "Workspace State"= $workspaceState
                "Error"          = $errorMessage
                "Error Detail"   = $errorDetail
            }
        )

        continue
    }
}

Write-Progress -Activity "Fetching workspace access details" -Completed

# ------------------------------------------------------------
# Export results
# ------------------------------------------------------------
try {
    $results | Export-Csv `
        -Path $outputFile `
        -NoTypeInformation `
        -Encoding UTF8

    Write-Host "Exported $($results.Count) entries to $outputFile"
}
catch {
    Write-Error "Failed to export main CSV."
}

if ($errorResults.Count -gt 0) {
    try {
        $errorResults | Export-Csv `
            -Path $errorFile `
            -NoTypeInformation `
            -Encoding UTF8

        Write-Warning "Some workspaces failed. See $errorFile"
    }
    catch {
        Write-Error "Failed to export error CSV."
    }
}
else {
    Write-Host "Completed with no workspace errors."
}
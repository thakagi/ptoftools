param (
    [switch]$ExcludePersonalWorkspaces
)

# Sign in to Power BI Service
Connect-PowerBIServiceAccount

# Set output file
$outputFile = "./workspace_users.csv"

# Get all workspaces
$workspaces = Get-PowerBIWorkspace -All

# Exclude personal workspaces if specified
if ($ExcludePersonalWorkspaces) {
    $workspaces = $workspaces | Where-Object { $_.Type -ne "PersonalGroup" }
}

if (-not $workspaces) {
    Write-Warning "No workspaces found."
    return
}

# Collect results in memory for better performance
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$totalCount = $workspaces.Count
$currentIndex = 0

# Extract user details for each workspace
foreach ($workspace in $workspaces) {
    $currentIndex++
    $workspaceId = $workspace.Id
    $workspaceName = $workspace.Name

    Write-Progress -Activity "Fetching workspace users" -Status "$currentIndex / $totalCount : $workspaceName" -PercentComplete (($currentIndex / $totalCount) * 100)

    try {
        # Fetch user details using the REST API
        $response = Invoke-PowerBIRestMethod -Url "groups/$workspaceId/users" -Method Get -ErrorAction Stop
        $usersJson = $response | ConvertFrom-Json

        if ($usersJson.value) {
            foreach ($user in $usersJson.value) {
                $results.Add([PSCustomObject]@{
                    "Workspace Name" = $workspaceName
                    "Workspace ID"   = $workspaceId
                    "User Email"     = $user.emailAddress
                    "Role"           = $user.groupUserAccessRight
                })
            }
        }
    }
    catch {
        Write-Warning "Failed to get users for workspace '$workspaceName' ($workspaceId): $_"
    }
}

Write-Progress -Activity "Fetching workspace users" -Completed

# Export all results at once
$results | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

Write-Host "Exported $($results.Count) entries to $outputFile"
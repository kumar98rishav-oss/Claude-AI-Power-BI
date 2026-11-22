# Starts an on-demand refresh of a published semantic model (this CHANGES the service).
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File refresh_now.ps1 -WorkspaceId <guid> -DatasetId <guid>
param(
    [Parameter(Mandatory = $true)][string]$WorkspaceId,
    [Parameter(Mandatory = $true)][string]$DatasetId
)
$ErrorActionPreference = 'Stop'
Import-Module MicrosoftPowerBIMgmt.Profile
try { $null = Get-PowerBIAccessToken } catch { Connect-PowerBIServiceAccount | Out-Null }

Invoke-PowerBIRestMethod -Url "groups/$WorkspaceId/datasets/$DatasetId/refreshes" -Method Post -Body '{}' | Out-Null
Write-Output "Refresh requested at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'). Check the status with the service-check script in about a minute."

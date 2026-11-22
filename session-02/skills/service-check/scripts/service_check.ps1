# Read-only: refresh history, data sources and gateway status of a published semantic model.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File service_check.ps1 -WorkspaceId <guid> -DatasetId <guid> [-Top 5]
# Needs: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser   (one time)
param(
    [Parameter(Mandatory = $true)][string]$WorkspaceId,
    [Parameter(Mandatory = $true)][string]$DatasetId,
    [int]$Top = 5
)
$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile)) {
    Write-Output "MicrosoftPowerBIMgmt is not installed. Run: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser"
    exit 1
}
Import-Module MicrosoftPowerBIMgmt.Profile

# Reuse an existing sign-in if there is one; otherwise open the browser sign-in.
try { $null = Get-PowerBIAccessToken } catch { Connect-PowerBIServiceAccount | Out-Null }

$base = "groups/$WorkspaceId/datasets/$DatasetId"

Write-Output "== Refresh history (last $Top) =="
try {
    $refreshes = (Invoke-PowerBIRestMethod -Url "$base/refreshes?`$top=$Top" -Method Get | ConvertFrom-Json).value
    foreach ($r in $refreshes) {
        $mins = if ($r.endTime) { [math]::Round(((Get-Date $r.endTime) - (Get-Date $r.startTime)).TotalMinutes, 1) } else { 'running' }
        Write-Output ("{0,-10} {1,-10} start {2}  minutes {3}" -f $r.refreshType, $r.status, $r.startTime, $mins)
        if ($r.status -eq 'Failed' -and $r.serviceExceptionJson) { Write-Output "    error: $($r.serviceExceptionJson)" }
    }
} catch { Write-Output "Could not read refresh history (you need Write permission on the model): $($_.Exception.Message)" }

Write-Output ""
Write-Output "== Data sources =="
try {
    $sources = (Invoke-PowerBIRestMethod -Url "$base/datasources" -Method Get | ConvertFrom-Json).value
    foreach ($s in $sources) {
        $gw = if ($s.gatewayId) { "gateway $($s.gatewayId)" } else { 'no gateway' }
        Write-Output ("{0,-12} {1}  ({2})" -f $s.datasourceType, ($s.connectionDetails | ConvertTo-Json -Compress), $gw)
        if ($s.gatewayId -and $s.datasourceId) {
            try {
                Invoke-PowerBIRestMethod -Url "gateways/$($s.gatewayId)/datasources/$($s.datasourceId)/status" -Method Get | Out-Null
                Write-Output "    gateway status: OK"
            } catch { Write-Output "    gateway status: could not check ($($_.Exception.Message))" }
        }
    }
} catch { Write-Output "Could not read data sources: $($_.Exception.Message)" }

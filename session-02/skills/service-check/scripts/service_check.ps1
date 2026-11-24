# Read-only: refresh history, data sources and gateway status of a published semantic model.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File service_check.ps1 -WorkspaceId <guid> -DatasetId <guid> [-Top 5]
# Needs: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser   (one time)
param(
    [Parameter(Mandatory = $true)][string]$WorkspaceId,
    [Parameter(Mandatory = $true)][string]$DatasetId,
    [int]$Top = 5
)
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'   # hides the sign-in library's "WARNING: False MSAL ..." noise

if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile)) {
    Write-Output "MicrosoftPowerBIMgmt is not installed. Run: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser"
    exit 1
}
Import-Module MicrosoftPowerBIMgmt.Profile

# Reuse an existing sign-in if there is one; otherwise open the browser sign-in.
try { $null = Get-PowerBIAccessToken } catch { Connect-PowerBIServiceAccount | Out-Null }

function Get-ApiError {
    # Invoke-PowerBIRestMethod only says "One or more errors occurred."; this pulls out the real reason.
    try {
        $e = Resolve-PowerBIError -Last
        if ($e.Message) { return ($e.Message -replace '\s+', ' ').Trim() }
    } catch { }
    return $_.Exception.Message
}

$base = "groups/$WorkspaceId/datasets/$DatasetId"

Write-Output "== Refresh history (last $Top) =="
try {
    $refreshes = @((Invoke-PowerBIRestMethod -Url "$base/refreshes?`$top=$Top" -Method Get | ConvertFrom-Json).value)
    if ($refreshes.Count -eq 0) {
        Write-Output "No refreshes in the service yet. Publishing from Desktop doesn't count as a refresh."
        Write-Output "Start one (semantic model > Refresh now, or /refresh-now) and run this check again."
    }
    foreach ($r in $refreshes) {
        $mins = if ($r.endTime) { [math]::Round(((Get-Date $r.endTime) - (Get-Date $r.startTime)).TotalMinutes, 1) } else { 'running' }
        Write-Output ("{0,-10} {1,-10} start {2}  minutes {3}" -f $r.refreshType, $r.status, $r.startTime, $mins)
        if ($r.status -eq 'Failed' -and $r.serviceExceptionJson) { Write-Output "    error: $($r.serviceExceptionJson)" }
    }
} catch { Write-Output "Could not read refresh history (you need Write permission on the model): $(Get-ApiError)" }

Write-Output ""
Write-Output "== Data sources =="
try {
    $sources = @((Invoke-PowerBIRestMethod -Url "$base/datasources" -Method Get | ConvertFrom-Json).value)
    $noGateway = @($sources | Where-Object { -not $_.gatewayId })
    foreach ($s in $sources) {
        $detail = if ($s.connectionDetails.path) { Split-Path $s.connectionDetails.path -Leaf } else { $s.connectionDetails | ConvertTo-Json -Compress }
        $gw = if ($s.gatewayId) { 'gateway' } else { 'NO GATEWAY' }
        Write-Output ("{0,-8} {1,-28} {2}" -f $s.datasourceType, $detail, $gw)
    }
    Write-Output ("{0} data sources, {1} without a gateway." -f $sources.Count, $noGateway.Count)

    Write-Output ""
    Write-Output "== Gateway status =="
    $gateways = @($sources | Where-Object { $_.gatewayId } | Group-Object gatewayId)
    if ($gateways.Count -eq 0) { Write-Output "No gateway is bound to this model." }
    foreach ($g in $gateways) {
        $first = $g.Group[0]
        try {
            Invoke-PowerBIRestMethod -Url "gateways/$($g.Name)/datasources/$($first.datasourceId)/status" -Method Get | Out-Null
            Write-Output "Gateway $($g.Name): reachable (checked with $($g.Count) data sources bound)."
        } catch {
            Write-Output "Gateway $($g.Name): status not available through the API: $(Get-ApiError)"
            Write-Output "    This API needs gateway admin rights and may not cover personal-mode gateways."
            Write-Output "    The refresh history above is the reliable check: a refresh that completes means the gateway works."
        }
    }
} catch { Write-Output "Could not read data sources: $(Get-ApiError)" }

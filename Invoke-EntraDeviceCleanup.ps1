<#
Version:     2.0
Author:      Peter Daalmans
Runbook:     Invoke-EntraDeviceCleanup
Description: Report on, disable and delete stale Entra ID device registrations.
             - Disable enabled devices inactive for more than CleanUPDisableDays
             - Delete disabled devices inactive for more than CleanUPDeleteDays
             - Never touches Autopilot-registered, Intune-managed (delete only) or hybrid joined devices
             - Posts a summary Adaptive Card to a Teams Workflows webhook

Runtime:     Azure Automation, PowerShell 7.4 (also works on 5.1)
Modules:     Microsoft.Graph.Authentication
             Microsoft.Graph.Identity.DirectoryManagement
             Microsoft.Graph.DeviceManagement

Graph permissions (application):
             Device.ReadWrite.All
             DeviceManagementManagedDevices.Read.All
Entra role on the identity:
             Cloud Device Administrator (needed to set accountEnabled and to delete devices app-only)

Automation variables:
  CleanUPReadOnly       Boolean  true = report only (default true)
  CleanUPDisableDays    Integer  e.g. 60   (negative values from v1 are accepted)
  CleanUPDeleteDays     Integer  e.g. 90   (must be larger than DisableDays)
  CleanUPMaxDeletes     Integer  safety cap; deletes are skipped when exceeded (default 100)
  CleanUPExcludeHybrid  Boolean  skip hybrid joined (ServerAd) devices (default true)
  CleanUPTeamsURI       String   Teams Workflows webhook URL (optional)
  CleanUPAPPID          String   App registration client id   } leave these empty/absent
  CleanUPTenantID       String   Tenant id                     } to use the Automation
  CleanUPCert           String   Certificate thumbprint        } account's managed identity

Release notes:
  1.0   Original published version.
  1.02  Read-only mode, Graph SDK, Teams message, using Claude AI.
  2.0   Managed identity support, delete via Cloud Device Administrator role,
        fixed Intune match (DeviceId instead of non-existent ObjectId),
        Intune devices fetched once with -All, devices without sign-in date handled,
        already disabled devices no longer re-disabled, hybrid joined exclusion,
        delete safety cap, per-device error handling, JSON built with ConvertTo-Json,
        webhook posted with Invoke-RestMethod (no Graph token sent to the webhook).

The script is provided "AS IS" with no warranties.
#>

$ErrorActionPreference = 'Stop'
$VerbosePreference     = 'SilentlyContinue'   # keeps Graph module import noise out of the job log

#region Helpers ------------------------------------------------------------------------------
function Get-AutoVar {
    param([string]$Name, $Default = $null)
    try {
        $value = Get-AutomationVariable -Name $Name -ErrorAction Stop
        if ($null -ne $value -and "$value" -ne '') { return $value }
    } catch { }
    return $Default
}

function Format-DeviceList {
    param([object[]]$List = @(), [int]$Max = 40)
    $List = @($List | Where-Object { $_ })
    if ($List.Count -eq 0) { return '_None_' }
    $lines = @($List | Select-Object -First $Max | ForEach-Object {
        "- **$($_.Name)** · $($_.OS) · last seen $($_.LastSeen) ($($_.DaysInactive) days)"
    })
    if ($List.Count -gt $Max) { $lines += "- …and $($List.Count - $Max) more (see job output)" }
    return ($lines -join "`n")
}
#endregion

#region Configuration ------------------------------------------------------------------------
$AppId          = Get-AutoVar 'CleanUPAPPID'
$TenantId       = Get-AutoVar 'CleanUPTenantID'
$CertThumbprint = Get-AutoVar 'CleanUPCert'
$TeamsUri       = Get-AutoVar 'CleanUPTeamsURI'

$DisableDays    = [math]::Abs([int](Get-AutoVar 'CleanUPDisableDays' 60))
$DeleteDays     = [math]::Abs([int](Get-AutoVar 'CleanUPDeleteDays'  90))
$MaxDeletes     = [int](Get-AutoVar 'CleanUPMaxDeletes' 100)
$ReadOnly       = [System.Convert]::ToBoolean((Get-AutoVar 'CleanUPReadOnly' $true))
$ExcludeHybrid  = [System.Convert]::ToBoolean((Get-AutoVar 'CleanUPExcludeHybrid' $true))

if ($DeleteDays -le $DisableDays) {
    throw "CleanUPDeleteDays ($DeleteDays) must be larger than CleanUPDisableDays ($DisableDays)."
}

$nowUtc        = (Get-Date).ToUniversalTime()
$disableCutoff = $nowUtc.AddDays(-$DisableDays)
$deleteCutoff  = $nowUtc.AddDays(-$DeleteDays)
$modeText      = if ($ReadOnly) { 'READ ONLY MODE' } else { 'PRODUCTION MODE' }
#endregion

#region Connect ------------------------------------------------------------------------------
if ($AppId -and $TenantId -and $CertThumbprint) {
    Write-Output "Connecting with app registration $AppId (certificate)."
    Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $CertThumbprint -NoWelcome
} else {
    Write-Output "Connecting with the Automation account's managed identity."
    Connect-MgGraph -Identity -NoWelcome
}
#endregion

#region Collect data -------------------------------------------------------------------------
$deviceProps = 'id','deviceId','displayName','accountEnabled','approximateLastSignInDateTime',
               'registrationDateTime','physicalIds','trustType','operatingSystem'
$devices = @(Get-MgDevice -All -Property $deviceProps)
Write-Output "Entra ID devices found: $($devices.Count)"

$intuneIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
Get-MgDeviceManagementManagedDevice -All -Property 'id','azureADDeviceId' | ForEach-Object {
    if ($_.AzureAdDeviceId -and $_.AzureAdDeviceId -ne '00000000-0000-0000-0000-000000000000') {
        [void]$intuneIds.Add($_.AzureAdDeviceId)
    }
}
Write-Output "Intune managed devices found: $($intuneIds.Count)"
#endregion

#region Classify -----------------------------------------------------------------------------
$toDisable      = New-Object System.Collections.Generic.List[object]
$toDelete       = New-Object System.Collections.Generic.List[object]
$skipAutopilot  = New-Object System.Collections.Generic.List[object]
$skipIntune     = New-Object System.Collections.Generic.List[object]
$skipHybrid     = New-Object System.Collections.Generic.List[object]
$failed         = New-Object System.Collections.Generic.List[object]

foreach ($d in $devices) {
    # Devices that never signed in fall back to their registration date
    $lastSeen = $d.ApproximateLastSignInDateTime
    if (-not $lastSeen) { $lastSeen = $d.RegistrationDateTime }
    if (-not $lastSeen) { continue }
    $lastSeen = ([datetime]$lastSeen).ToUniversalTime()

    if ($lastSeen -gt $disableCutoff) { continue }   # active, nothing to do

    $record = [pscustomobject]@{
        Name         = $d.DisplayName
        Id           = $d.Id
        DeviceId     = $d.DeviceId
        OS           = $d.OperatingSystem
        TrustType    = $d.TrustType
        Enabled      = $d.AccountEnabled
        LastSeen     = $lastSeen.ToString('yyyy-MM-dd')
        DaysInactive = [int]($nowUtc - $lastSeen).TotalDays
    }

    if ($ExcludeHybrid -and $d.TrustType -eq 'ServerAd') { $skipHybrid.Add($record); continue }
    if (@($d.PhysicalIds) -match '\[ZTDID\]')            { $skipAutopilot.Add($record); continue }

    if ($d.AccountEnabled) {
        $toDisable.Add($record)
    }
    elseif ($lastSeen -le $deleteCutoff) {
        if ($intuneIds.Contains([string]$d.DeviceId)) { $skipIntune.Add($record) }
        else                                          { $toDelete.Add($record) }
    }
}
#endregion

#region Act ----------------------------------------------------------------------------------
$deleteCapHit = $false
if ($toDelete.Count -gt $MaxDeletes) {
    $deleteCapHit = $true
    Write-Warning "Delete candidates ($($toDelete.Count)) exceed CleanUPMaxDeletes ($MaxDeletes). No devices deleted this run."
}

if (-not $ReadOnly) {
    if (-not $deleteCapHit) {
        foreach ($r in $toDelete) {
            try   { Remove-MgDevice -DeviceId $r.Id -ErrorAction Stop }
            catch { $failed.Add([pscustomobject]@{ Name = $r.Name; Action = 'Delete';  Error = $_.Exception.Message }) }
        }
    }
    foreach ($r in $toDisable) {
        try   { Update-MgDevice -DeviceId $r.Id -BodyParameter @{ accountEnabled = $false } -ErrorAction Stop }
        catch { $failed.Add([pscustomobject]@{ Name = $r.Name; Action = 'Disable'; Error = $_.Exception.Message }) }
    }
}
#endregion

#region Job output ---------------------------------------------------------------------------
Write-Output "`n===== Entra ID Device Cleanup - $modeText ====="
Write-Output "Disable after $DisableDays days | Delete after $DeleteDays days"
Write-Output "`n--- Disable ($($toDisable.Count)) ---";               $toDisable     | Format-Table Name, OS, LastSeen, DaysInactive, DeviceId -AutoSize | Out-String -Width 250 | Write-Output
Write-Output "--- Delete ($($toDelete.Count)) ---";                   $toDelete      | Format-Table Name, OS, LastSeen, DaysInactive, DeviceId -AutoSize | Out-String -Width 250 | Write-Output
Write-Output "--- Skipped: Autopilot ($($skipAutopilot.Count)) ---";  $skipAutopilot | Format-Table Name, OS, LastSeen, Enabled -AutoSize | Out-String -Width 250 | Write-Output
Write-Output "--- Skipped: still in Intune ($($skipIntune.Count)) ---"; $skipIntune  | Format-Table Name, OS, LastSeen -AutoSize | Out-String -Width 250 | Write-Output
Write-Output "--- Skipped: hybrid joined ($($skipHybrid.Count)) ---"; $skipHybrid    | Format-Table Name, OS, LastSeen -AutoSize | Out-String -Width 250 | Write-Output
if ($failed.Count) {
    Write-Output "--- Failed ($($failed.Count)) ---"
    $failed | Format-Table -AutoSize -Wrap | Out-String -Width 250 | Write-Output
    if ($failed.Error -match 'Insufficient privileges|Authorization_RequestDenied|403') {
        Write-Warning "Access denied: assign the Cloud Device Administrator role to the runbook identity."
    }
}
#endregion

#region Teams notification -------------------------------------------------------------------
if ($TeamsUri) {
    try {
        $verbDisable = if ($ReadOnly) { 'Would be disabled' } else { 'Disabled' }
        $verbDelete  = if ($ReadOnly) { 'Would be deleted'  } else { 'Deleted'  }
        if ($deleteCapHit) { $verbDelete = "Delete candidates (NOT deleted, cap of $MaxDeletes exceeded)" }

        # Plain string arrays/strings only, so nothing type-sensitive reaches the card
        $disableText = [string](Format-DeviceList -List $toDisable.ToArray())
        $deleteText  = [string](Format-DeviceList -List $toDelete.ToArray())

        $facts = New-Object System.Collections.ArrayList
        foreach ($pair in @(
            @('Devices scanned',           [string]$devices.Count),
            @('Disable after (days)',      [string]$DisableDays),
            @('Delete after (days)',       [string]$DeleteDays),
            @($verbDisable,                [string]$toDisable.Count),
            @($verbDelete,                 [string]$toDelete.Count),
            @('Skipped (Autopilot)',       [string]$skipAutopilot.Count),
            @('Skipped (still in Intune)', [string]$skipIntune.Count),
            @('Skipped (hybrid joined)',   [string]$skipHybrid.Count),
            @('Failed actions',            [string]$failed.Count)
        )) {
            [void]$facts.Add([ordered]@{ title = [string]$pair[0]; value = [string]$pair[1] })
        }

        $body = New-Object System.Collections.ArrayList
        [void]$body.Add([ordered]@{ type = 'TextBlock'; text = "Entra ID Device Cleanup - $modeText"; style = 'heading'; wrap = $true })
        [void]$body.Add([ordered]@{ type = 'TextBlock'; text = (Get-Date).ToString('dd-MMM-yyyy HH:mm'); isSubtle = $true; spacing = 'None' })
        [void]$body.Add([ordered]@{ type = 'FactSet'; facts = $facts.ToArray() })
        [void]$body.Add([ordered]@{ type = 'TextBlock'; text = $verbDisable; weight = 'Bolder'; separator = $true; wrap = $true })
        [void]$body.Add([ordered]@{ type = 'TextBlock'; text = $disableText; wrap = $true })
        [void]$body.Add([ordered]@{ type = 'TextBlock'; text = $verbDelete; weight = 'Bolder'; separator = $true; wrap = $true })
        [void]$body.Add([ordered]@{ type = 'TextBlock'; text = $deleteText; wrap = $true })

        if ($failed.Count -gt 0) {
            $failText = (@($failed | Select-Object -First 20 | ForEach-Object { "- **$($_.Name)** ($($_.Action)): $($_.Error)" })) -join "`n"
            [void]$body.Add([ordered]@{ type = 'TextBlock'; text = 'Failed'; weight = 'Bolder'; color = 'Attention'; separator = $true })
            [void]$body.Add([ordered]@{ type = 'TextBlock'; text = [string]$failText; wrap = $true })
        }

        $card = [ordered]@{
            type        = 'message'
            attachments = @(
                [ordered]@{
                    contentType = 'application/vnd.microsoft.card.adaptive'
                    contentUrl  = $null
                    content     = [ordered]@{
                        '$schema' = 'http://adaptivecards.io/schemas/adaptive-card.json'
                        type      = 'AdaptiveCard'
                        version   = '1.4'
                        msteams   = [ordered]@{ width = 'Full' }
                        body      = $body.ToArray()
                    }
                }
            )
        }

        $json  = $card | ConvertTo-Json -Depth 20 -Compress
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$json)
        Invoke-RestMethod -Uri $TeamsUri -Method Post -ContentType 'application/json; charset=utf-8' -Body $bytes | Out-Null
        Write-Output "Teams notification sent ($($bytes.Length) bytes)."
    } catch {
        Write-Warning "Teams notification failed: $($_.Exception.Message)"
        Write-Warning ("At: " + $_.InvocationInfo.PositionMessage)
    }
}
#endregion

Disconnect-MgGraph | Out-Null

if ($failed.Count) { throw "$($failed.Count) device action(s) failed. See job output." }

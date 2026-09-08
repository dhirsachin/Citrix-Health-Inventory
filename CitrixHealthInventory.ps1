<#
.SYNOPSIS
    Citrix PVS and XDC Server Information Retrieval Tool.

.DESCRIPTION
    Auto-detects whether the host is a Citrix PVS Master or Citrix Delivery
    Controller (XDC), then presents the matching WPF GUI for retrieving and
    exporting server information.

    Features:
    - Modern styled WPF interface with Citrix-inspired colour scheme
    - Parallel IPv4 ping diagnostics via RunspacePool (PVS mode)
    - CSV export
    - Whole-farm, PVS device-collection, XDC Machine Catalog/Delivery Group, single-server, and pasted server-list retrieval scopes
    - Select All / Deselect All, search/filter, record count, status bar
    - Optional DHCP validation across every discovered PVS Master
    - Structured file logging to a dedicated local application-data folder

.NOTES
    Requires: Citrix PVS PowerShell SDK (PVS Server) or Citrix Broker PowerShell SDK (DDC)
    Platform: Intended for Windows Server 2012/2016/2019/2022; the startup check
              also uses an OS-version fallback when the caption is not recognized.
    Author:   Sachin
    Release:  2.6.0-rc4
    ReleaseDate: 2026-09-04
    ChangeLog: See CHANGELOG.md for release history. Keep this header limited to
               runtime purpose, compatibility, and deployment assumptions.
#>

#Requires -Version 5.1

<#
===============================================================================
MAINTAINER GUIDE AND EXECUTION MAP
===============================================================================

Purpose
-------
This is one script with two role-specific interfaces:

  PVS mode  - Uses the Citrix PVS SDK to retrieve target-device information.
  XDC mode  - Uses the Citrix Broker SDK to retrieve broker-machine information.

The internal name "DDC" is retained in some functions and variables because that
is the traditional Citrix Delivery Controller abbreviation. In the GUI, the same
role is presented to users as XDC.

Startup flow
------------
1. Test-StartupPrereqs checks the OS caption/version, PowerShell 5.1, and WPF.
2. Test-PVSServer checks for a running StreamService. PVS takes precedence when
   both Citrix services happen to be present.
3. If PVS is not detected, Test-DeliveryController checks CitrixBrokerService.
4. Test-RolePrereqs confirms that the detected role's core inventory cmdlet can
   be loaded. Optional per-field cmdlets are handled during retrieval.
5. Show-PVSWindow or Show-DDCWindow builds the WPF interface and owns its events.

Common retrieval flow
---------------------
Both windows use the same overall pipeline:

1. Selecting PVS Device Collection or XDC Machine Catalog/Delivery Group starts
   one dedicated background discovery runspace immediately. It loads the trusted
   Citrix SDK, reads one farm/site inventory, returns identity-only preview data,
   and fills the visible inline dropdown with no default selection.
2. Read and validate the selected scope and output fields on the UI thread.
   Named scopes additionally require a current preview and an explicit dropdown
   choice; a missing preview or opaque key fails closed.
3. Start the retrieval runspace and build one fresh farm/site inventory when
   Retrieve Data is clicked. XDC paging may use several Broker calls to construct
   that one logical inventory; it does not contact each VDA independently.
4. For a named scope, rederive the current choices from the fresh inventory,
   revalidate the opaque identity key, and filter locally. XDC also requires a
   group name plus valid UUID or positive UID. Server List and Single Server
   match requested names locally with Select-ItemsByServerList.
5. For large PVS optional workloads, pause after selection and request approval
   using the exact matched-device count before per-device processing begins.
6. Enrich each matched record only with the optional fields the user selected.
7. Return one completed result envelope to the UI thread; partial rows are never
   published when a run is cancelled or the inventory query fails.
8. Store the complete result in $script:pvsFullData or $script:ddcFullData.
9. Bind that result to the DataGrid; filtering changes the displayed view only.
10. Export and clipboard actions operate on the currently displayed grid rows.

Retrieval-scope rules
---------------------
Whole Farm    Retrieves every record returned by the Citrix SDK. XDC sends this
              as an explicit scope flag; a missing/conflicting scope fails closed.
Device Collection (PVS only) shows a site-aware inline dropdown immediately and
              loads it asynchronously from a farm-inventory preview. Retrieve
              reads one current inventory and filters it locally. CollectionId is preferred;
              older SDK shapes use SiteId or SiteName plus CollectionName. No
              target VDA is contacted by this base inventory query.
Machine Catalog (XDC only) shows an inline catalog dropdown and loads it from the
              Broker inventory preview, then revalidates the key against one
              current inventory before row shaping or optional DNS work. A name
              plus valid UUID/positive UID is required.
Delivery Group (XDC only) follows the same flow for assigned Delivery Groups;
              unassigned machines and unsafe metadata are counted separately;
              neither appears in that picker.
Server List   Accepts newline-, comma-, or semicolon-separated names, preserves
              request order, ignores repeated aliases, and reports names that
              are missing or ambiguous.
Single Server Accepts exactly one name. Its input is stored separately from the
              Server List input so switching modes never copies one into the
              other.

Name matching prefers an exact inventory name, DOMAIN\SERVER value, or FQDN.
Short-name aliases are used only when they resolve to exactly one inventory
record. Ambiguous aliases are reported rather than guessed.

PVS data sources
----------------
Get-PVSDevice supplies the device name and Device Collection. Optional fields:

  vDisk        Get-PVSDiskInfo
  Ping Reachability  Bounded DNS plus parallel ICMP through Get-BulkPingResults
  Reboot Day   Get-PVSDevicePersonality key "Reboot"
  CSR Server   Get-PVSDevicePersonality key "CSAServer"
  XDC Server   Get-PVSDevicePersonality key "XDC_LIST" (normally DR nodes)
  Server IP    Existing SDK property when available, otherwise bounded DNS

Device Collection is intentionally always included in PVS output.

XDC data sources
----------------
Get-BrokerMachineInventory pages through Get-BrokerMachine when the installed
Broker SDK supports -Skip. The selected output columns are shaped from Broker
machine properties; Server IP uses an SDK address when present and bounded DNS
as a fallback.

DHCP validation
---------------
DHCP validation is intentionally opt-in and is not affected by Select All
Information because it can be expensive in large farms.

Resolve-PvsMastersForDhcp discovers the local PVS Master plus every qualified
Master identity returned by Get-PvsServer. DHCP then follows the selected scope:

  Whole Farm    Baseline only: DHCP service, exact scope definitions, and
                default standard server/scope options. No reservations are read.
  Device Collection Targeted: only devices in the explicitly selected collection
                are validated across every Master. The normal path accepts at
                most 1,000 matched targets.
  Server List   Targeted: only matched PVS devices are validated across every
                Master through batched IP point queries. A failed batch is marked
                incomplete and never broadens into a full-scope reservation scan.
                One run accepts at most 1,000 matched targets and 5,000 target-by-Master cells.
  Single Server Targeted: the same path for exactly one matched device, including
                its scope, reservation, and effective options on every Master.
  Deep Audit    Explicit Whole Farm operation that adds every reservation and
                reservation-level option override. It is default-off and warns
                before starting because it can create thousands of remote reads.

Get-DhcpOptionSet reads configured standard options in one call per tier and
represents every requested value as Value, Unset, or ReadFailed. Effective values
follow reservation, scope, then server precedence, but inheritance stops at a
ReadFailed tier because an unknown override may exist. Complete Masters are
compared for options 3, 6, 11, 15, and 67; option 66 is validated against each
individual Master's own IPv4 address. Coverage failures produce Partial,
TimedOut, or Unavailable status and can never produce an authoritative clean
comparison. Targeted and Deep modes also compare reservation names and normalized
client IDs. DHCP policy, vendor-class, and user-class effective values are outside
this default-option validation and are stated as such in the output.

DHCP runs inside the PVS background workflow and uses at most two Master workers.
Calls remain serial within each Master. Master and overall deadlines are soft:
the script requests cooperative stop, but a blocked remote DHCP/RPC call may not
return immediately. The existing second-close Force Exit remains the local escape
path. Only one operator should run DHCP validation per PVS farm.

All DHCP output is routed exclusively to the DHCP Details tab. The normal Output
tab remains dedicated to server/collection data. Completion status and validation
status are reported separately so "Complete" means full coverage, not "healthy."

State, logging, and UI responsiveness
-------------------------------------
$script:pvsFullData and $script:ddcFullData are the unfiltered source datasets.
The scope-state hashtables remember Single Server and Server List text
independently. Separate UI-owned DispatcherTimers consume scope-discovery and
retrieval messages; only the UI dispatcher changes WPF controls. Discovery state
caches identity-only preview data, choices, and independent MC/DG keys. Refresh
invalidates that preview and all saved choices. A preview is accepted for at most
five minutes; Retrieve revalidates its opaque key against one fresh inventory.
Cancel requests are cooperative: the active SDK call may need to return before
cleanup completes.

A machine-wide mutex blocks a second copy on the same Windows server, including
another RDP session. It cannot detect copies on other PVS Masters or DDCs, so one
designated execution host per farm/site remains an operational requirement.

Write-Log appends to one daily file:
  %LOCALAPPDATA%\CitrixDataPull\Logs\PVS_XDC_Tool_yyyyMMdd.log

The log records startup checks, selected scope/fields, result counts, selection
problems, DHCP summaries, export paths, failures, and execution time. It does not
contain the full table unless individual values are included in an error message.
===============================================================================
#>

# -------------------------------------------------------------
# Region: Logging
# -------------------------------------------------------------
# Centralized logging helpers. All modes append to the same date-based file so
# startup, retrieval, export, and error events can be reviewed in one timeline.
$script:ToolRelease = '2.6.0-rc4'
$script:RunCorrelationId = [guid]::NewGuid().ToString('N').Substring(0, 12)
$script:MaximumLogMessageLength = 8192

# Keep operational logs on a local fixed drive. This avoids outbound SMB access
# or shared-file tampering when a caller has redirected the mutable TEMP variable.
function Get-CitrixDataPullLogDirectory {
    $localApplicationData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
    if ([string]::IsNullOrWhiteSpace($localApplicationData)) {
        throw 'Windows did not return a LocalApplicationData directory for the current user.'
    }

    $localRoot = [IO.Path]::GetFullPath($localApplicationData)
    if (-not [IO.Path]::IsPathRooted($localRoot) -or $localRoot.StartsWith('\\')) {
        throw "LocalApplicationData must resolve to a local absolute path; received [$localRoot]."
    }

    $driveRoot = [IO.Path]::GetPathRoot($localRoot)
    $drive = [IO.DriveInfo]::new($driveRoot)
    if ($drive.DriveType -ne [IO.DriveType]::Fixed) {
        throw "LocalApplicationData must be on a fixed local drive; [$driveRoot] is [$($drive.DriveType)]."
    }

    $logDirectory = [IO.Path]::GetFullPath((Join-Path $localRoot 'CitrixDataPull\Logs'))
    [void][IO.Directory]::CreateDirectory($logDirectory)
    $directoryInfo = [IO.DirectoryInfo]::new($logDirectory)
    if (($directoryInfo.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "The execution-log directory cannot be a reparse point: [$logDirectory]."
    }
    return $logDirectory
}

$script:LogDirectory = ''
$script:LogPath = '<logging unavailable>'
try {
    $script:LogDirectory = Get-CitrixDataPullLogDirectory
    $script:LogPath = Join-Path $script:LogDirectory "PVS_XDC_Tool_$(Get-Date -Format 'yyyyMMdd').log"
} catch {
    # Bootstrap failure is fail-closed: do not silently fall back to TEMP or a
    # network path. Console output remains available even before WPF is loaded.
    $bootstrapMessage = "Unable to initialize the protected local execution-log directory: $($_.Exception.Message)"
    try { [Console]::Error.WriteLine($bootstrapMessage) } catch { }
    exit 13
}
$script:LogWriteFailureReported = $false
$script:LogWriteFailureGuiReported = $false
$script:LogWriteFailureMessage = ''

# Removes only common credential-bearing key/value pairs before a message enters
# the persistent execution log. Server names, IP addresses, and Citrix object
# identifiers remain available because they are necessary troubleshooting context.
function ConvertTo-SafeLogMessage {
    param([string]$Message)

    $safeMessage = [string]$Message
    foreach ($pattern in @(
        '(?i)(password|passwd|pwd)\s*[:=]\s*([^|;\r\n\s]+)',
        '(?i)(access[_-]?token|refresh[_-]?token|client[_-]?secret|api[_-]?key)\s*[:=]\s*([^|;\r\n\s]+)'
    )) {
        $safeMessage = [regex]::Replace($safeMessage, $pattern, '$1=<redacted>')
    }
    return $safeMessage
}

<#
.SYNOPSIS
    Writes a timestamped entry to the daily log file.
#>
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [string]$Component = 'Application',
        [string]$Operation = 'General',
        [string]$Role = '',
        [string]$WorkerRunId = '',
        [string]$Retryable = '',
        [string]$SuggestedAction = ''
    )
    $timestamp = [DateTimeOffset]::Now.ToString('yyyy-MM-ddTHH:mm:ss.fffzzz', [Globalization.CultureInfo]::InvariantCulture)
    $singleLineMessage = ConvertTo-SafeLogMessage -Message ([regex]::Replace([string]$Message, '[\r\n]+', ' '))
    $context = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($Component)) { [void]$context.Add("Component=$Component") }
    if (-not [string]::IsNullOrWhiteSpace($Operation)) { [void]$context.Add("Operation=$Operation") }
    if (-not [string]::IsNullOrWhiteSpace($Role)) { [void]$context.Add("Role=$Role") }
    if (-not [string]::IsNullOrWhiteSpace($WorkerRunId)) { [void]$context.Add("WorkerRunId=$WorkerRunId") }
    if (-not [string]::IsNullOrWhiteSpace($Retryable)) { [void]$context.Add("Retryable=$Retryable") }
    if (-not [string]::IsNullOrWhiteSpace($SuggestedAction)) { [void]$context.Add("SuggestedAction=$SuggestedAction") }
    if ($context.Count -gt 0) { $singleLineMessage = (($context -join ' | ') + ' | ' + $singleLineMessage) }
    if ($singleLineMessage.Length -gt $script:MaximumLogMessageLength) {
        $originalLength = $singleLineMessage.Length
        $singleLineMessage = $singleLineMessage.Substring(0, $script:MaximumLogMessageLength) + " ... [truncated; originalChars=$originalLength]"
    }
    $entry = "[$timestamp] [$Level] [Run=$($script:RunCorrelationId)] [PID=$PID] $singleLineMessage"
    try {
        Add-Content -LiteralPath $script:LogPath -Value $entry -Encoding UTF8 -ErrorAction Stop
    } catch {
        $script:LogWriteFailureMessage = $_.Exception.Message
        if (-not $script:LogWriteFailureReported) {
            $script:LogWriteFailureReported = $true
            Write-Warning "The execution log cannot be written at [$($script:LogPath)]: $($_.Exception.Message)"
        }
        Show-LogWriteFailureWarning
    }
}

# Converts an ErrorRecord into one bounded, single-line diagnostic for support.
# User dialogs keep the short exception message; logs retain type, phase, error
# identifier, source position, and worker stack without serializing target data.
function Get-ErrorDiagnosticText {
    param(
        [object]$ErrorRecord,
        [string]$Phase = 'Unspecified phase',
        [ValidateRange(256, 8192)]
        [int]$MaximumLength = 4096
    )

    if ($null -eq $ErrorRecord) { return "Phase=$Phase | ErrorRecord=Unavailable" }
    $exceptionType = if ($null -ne $ErrorRecord.Exception) { $ErrorRecord.Exception.GetType().FullName } else { 'Unavailable' }
    $message = if ($null -ne $ErrorRecord.Exception) { [string]$ErrorRecord.Exception.Message } else { [string]$ErrorRecord }
    $fullyQualifiedErrorId = if ($ErrorRecord.PSObject.Properties.Match('FullyQualifiedErrorId').Count -gt 0) { [string]$ErrorRecord.FullyQualifiedErrorId } else { '' }
    $category = if ($ErrorRecord.PSObject.Properties.Match('CategoryInfo').Count -gt 0) { [string]$ErrorRecord.CategoryInfo } else { '' }
    $position = if ($null -ne $ErrorRecord.InvocationInfo) { [string]$ErrorRecord.InvocationInfo.PositionMessage } else { '' }
    $stack = if ($ErrorRecord.PSObject.Properties.Match('ScriptStackTrace').Count -gt 0) { [string]$ErrorRecord.ScriptStackTrace } else { '' }
    $detail = "Phase=$Phase | ExceptionType=$exceptionType | Message=$message | FullyQualifiedErrorId=$fullyQualifiedErrorId | Category=$category | Position=$position | Stack=$stack"
    $detail = [regex]::Replace($detail, '[\r\n]+', ' ')
    if ($detail.Length -gt $MaximumLength) {
        return $detail.Substring(0, $MaximumLength) + ' ... [diagnostic truncated]'
    }
    return $detail
}

# Uses .NET directly so release traceability does not depend on a profile-defined
# Get-FileHash command. The digest is logged, but organizational code signing is
# still the authenticity control for distributed production copies.
function Get-FileSha256Hex {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $stream = $null
    $algorithm = $null
    try {
        $stream = [IO.File]::Open($LiteralPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $algorithm = [Security.Cryptography.SHA256]::Create()
        $hashBytes = $algorithm.ComputeHash($stream)
        return ([BitConverter]::ToString($hashBytes) -replace '-', '').ToLowerInvariant()
    } finally {
        if ($null -ne $algorithm) { $algorithm.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

# Surfaces an audit-log gap once WPF is available. This helper deliberately does
# not call Write-Log, avoiding recursion when the log location itself is failing.
function Show-LogWriteFailureWarning {
    if (-not $script:LogWriteFailureReported -or $script:LogWriteFailureGuiReported) { return }
    try {
        [System.Windows.MessageBox]::Show(
            "The script cannot write to its daily execution log.`n`nLocation:`n$($script:LogPath)`n`nError:`n$($script:LogWriteFailureMessage)`n`nRetrieval may continue, but this run has an audit-log gap.",
            'Execution Logging Unavailable', 'OK', 'Warning') | Out-Null
        $script:LogWriteFailureGuiReported = $true
    } catch {
        # WPF may not be loaded during the earliest startup writes. The entry
        # point and UI polling callbacks invoke this helper again later.
    }
}

# Opens today's execution log in Notepad from either role-specific window.
function Open-CurrentLog {
    if (-not (Test-Path -LiteralPath $script:LogPath -PathType Leaf)) {
        [System.Windows.MessageBox]::Show(
            "The daily log file could not be found.`n`nExpected location:`n$($script:LogPath)",
            'Log File Not Found', 'OK', 'Warning') | Out-Null
        return
    }

    try {
        Write-Log "Opening daily log file [$($script:LogPath)]"
        $systemDirectory = [Environment]::SystemDirectory
        if ([string]::IsNullOrWhiteSpace($systemDirectory)) {
            throw 'The operating system did not return its protected system directory.'
        }
        $notepadPath = Join-Path $systemDirectory 'notepad.exe'
        if (-not (Test-Path -LiteralPath $notepadPath -PathType Leaf)) {
            throw "The trusted Notepad executable was not found at [$notepadPath]."
        }
        if ($script:LogPath.Contains('"')) {
            throw 'The log path contains an unsupported quotation-mark character.'
        }
        $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
        $startInfo.FileName = $notepadPath
        $startInfo.Arguments = '"' + $script:LogPath + '"'
        $startInfo.UseShellExecute = $false
        $notepadProcess = [System.Diagnostics.Process]::Start($startInfo)
        if ($null -eq $notepadProcess) {
            throw 'Windows did not start the trusted Notepad process.'
        }
        $notepadProcess.Dispose()
    } catch {
        Write-Log "Unable to open daily log file | $(Get-ErrorDiagnosticText -ErrorRecord $_ -Phase 'Open execution log')" -Level ERROR
        [System.Windows.MessageBox]::Show(
            "Unable to open the daily log file.`n`nLocation:`n$($script:LogPath)`n`nError: $($_.Exception.Message)",
            'Unable to Open Log', 'OK', 'Error') | Out-Null
    }
}


# -------------------------------------------------------------
# Region: Server Detection
# -------------------------------------------------------------
# These checks identify which GUI to launch. They deliberately require the
# relevant Citrix service to be running, not merely installed.
$script:RoleDetectionDetails = [System.Collections.Generic.List[string]]::new()
<#
.SYNOPSIS
    Detects if the local machine is a Citrix PVS server.
#>
function Test-PVSServer {
    try {
        $svc = Get-Service -Name 'StreamService' -ErrorAction Stop
        if ($null -ne $svc -and $svc.Status -eq 'Running') {
            Write-Log "StreamService found and running - PVS server detected"
            return $true
        }
        $detail = "PVS StreamService is installed but its status is [$($svc.Status)]; PVS mode requires the service to be Running."
        [void]$script:RoleDetectionDetails.Add($detail)
        Write-Log $detail -Level WARN
        return $false
    } catch {
        if ([string]$_.FullyQualifiedErrorId -like 'NoServiceFoundForGivenName*') {
            $detail = 'PVS StreamService is not installed on this server.'
            [void]$script:RoleDetectionDetails.Add($detail)
            Write-Log $detail
        } else {
            $detail = "Unable to query PVS StreamService: $($_.Exception.Message)"
            [void]$script:RoleDetectionDetails.Add($detail)
            Write-Log $detail -Level ERROR
        }
        return $false
    }
}

<#
.SYNOPSIS
    Detects if the local machine is a Citrix Delivery Controller.
#>
function Test-DeliveryController {
    try {
        $svc = Get-Service -Name 'CitrixBrokerService' -ErrorAction Stop
        if ($null -ne $svc -and $svc.Status -eq 'Running') {
            Write-Log "CitrixBrokerService found and running - DDC detected"
            return $true
        }
        $detail = "XDC CitrixBrokerService is installed but its status is [$($svc.Status)]; XDC mode requires the service to be Running."
        [void]$script:RoleDetectionDetails.Add($detail)
        Write-Log $detail -Level WARN
        return $false
    } catch {
        if ([string]$_.FullyQualifiedErrorId -like 'NoServiceFoundForGivenName*') {
            $detail = 'XDC CitrixBrokerService is not installed on this server.'
            [void]$script:RoleDetectionDetails.Add($detail)
            Write-Log $detail
        } else {
            $detail = "Unable to query XDC CitrixBrokerService: $($_.Exception.Message)"
            [void]$script:RoleDetectionDetails.Add($detail)
            Write-Log $detail -Level ERROR
        }
        return $false
    }
}

# -------------------------------------------------------------
# Region: Parallel Ping Helper (RunspacePool)
# -------------------------------------------------------------
# PVS ping reachability is an ICMP diagnostic, not a Citrix registration or
# electrical power state. DNS and ping run in a bounded pool so failures can be
# classified without accumulating long serial waits.
<#
.SYNOPSIS
    Resolves and pings multiple devices in parallel.
.DESCRIPTION
    Uses bounded DNS resolution followed by ICMP. Returned values are
    Reachable, No ICMP Reply, DNS Resolution Failed, or Ping Error.
#>
function Get-BulkPingResults {
    param(
        [string[]]$ComputerNames,
        [int]$ThrottleLimit = 30,
        [int]$TimeoutMs     = 1000,
        [int]$DnsTimeoutMs  = 1500,
        [int]$DnsCleanupWaitMs = 10000,
        [object]$MessageQueue,
        [System.Threading.CancellationToken]$CancellationToken = [System.Threading.CancellationToken]::None
    )

    $results = @{}
    if (-not $ComputerNames -or $ComputerNames.Count -eq 0) { return $results }

    $pool = $null
    $dnsTaskQueue = $null
    $dnsTaskSlots = $null
    $jobs = [System.Collections.Generic.List[object]]::new()
    try {
        $pool = [RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit)
        $pool.Open()
        # Timed-out .NET DNS tasks cannot be cancelled. A shared semaphore and
        # task queue cap the whole ping phase at the same number of outstanding
        # resolver operations as active ping workers, even after a worker has
        # returned a timeout result.
        $maximumOutstandingDnsTasks = [Math]::Max(1, $ThrottleLimit)
        $dnsTaskQueue = [System.Collections.Concurrent.ConcurrentQueue[System.Threading.Tasks.Task]]::new()
        $dnsTaskSlots = [System.Threading.SemaphoreSlim]::new($maximumOutstandingDnsTasks, $maximumOutstandingDnsTasks)

        $scriptBlock = {
            param($Name, $Timeout, $DnsTimeout, $DnsCleanupWait, $Token, $DnsTaskQueue, $DnsTaskSlots)

            $dnsTask = $null
            $dnsSlotReserved = $false
            $ping = $null
            $phase = 'DNS'
            try {
                if ($Token.IsCancellationRequested) { return 'Cancelled' }

                $targetAddress = $null
                $literalAddress = $null
                if ([System.Net.IPAddress]::TryParse($Name, [ref]$literalAddress) -and
                    $literalAddress.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
                    $targetAddress = $literalAddress
                } else {
                    # Drain completed tasks left by earlier timed-out workers and
                    # return their semaphore leases. Incomplete tasks are rotated
                    # back into the concurrent queue for a later worker to observe.
                    $drainCount = [Math]::Max(0, $DnsTaskQueue.Count)
                    for ($drainIndex = 0; $drainIndex -lt $drainCount; $drainIndex++) {
                        $queuedTask = $null
                        if (-not $DnsTaskQueue.TryDequeue([ref]$queuedTask)) { break }
                        if ($queuedTask.IsCompleted) {
                            try { [void]$queuedTask.Exception } catch { }
                            try { [void]$DnsTaskSlots.Release() } catch { }
                        } else {
                            $DnsTaskQueue.Enqueue($queuedTask)
                        }
                    }

                    if (-not $DnsTaskSlots.Wait(0)) {
                        return 'DNS Resolution Failed'
                    }
                    $dnsSlotReserved = $true
                    try {
                        $dnsTask = [System.Net.Dns]::GetHostAddressesAsync($Name)
                        $DnsTaskQueue.Enqueue($dnsTask)
                        $dnsSlotReserved = $false
                    } catch {
                        if ($dnsSlotReserved) {
                            try { [void]$DnsTaskSlots.Release() } catch { }
                            $dnsSlotReserved = $false
                        }
                        throw
                    }

                    $dnsCompleted = $false
                    try {
                        $dnsCompleted = $dnsTask.Wait($DnsTimeout, $Token)
                    } catch {
                        if ($Token.IsCancellationRequested) { return 'Cancelled' }
                        throw
                    }
                    if (-not $dnsCompleted) {
                        # Keep this pool slot occupied briefly. If the underlying
                        # task still does not return, it remains in the shared,
                        # bounded queue and no unbounded resolver wave can form.
                        $cleanupDeadline = [DateTime]::UtcNow.AddMilliseconds([Math]::Max(0, $DnsCleanupWait))
                        while (-not $dnsTask.IsCompleted -and [DateTime]::UtcNow -lt $cleanupDeadline) {
                            if ($Token.IsCancellationRequested) { return 'Cancelled' }
                            try {
                                [void]$dnsTask.Wait(100, $Token)
                            } catch {
                                if ($Token.IsCancellationRequested) { return 'Cancelled' }
                                break
                            }
                        }
                    }
                    if (-not $dnsTask.IsCompleted) {
                        return 'DNS Resolution Failed'
                    }

                    $targetAddress = @($dnsTask.GetAwaiter().GetResult()) |
                        Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
                        Select-Object -First 1
                }

                if (-not $targetAddress) { return 'DNS Resolution Failed' }
                if ($Token.IsCancellationRequested) { return 'Cancelled' }

                $phase = 'Ping'
                $ping = New-Object System.Net.NetworkInformation.Ping
                $reply = $ping.Send($targetAddress, $Timeout)
                if ($null -eq $reply) { return 'Ping Error' }
                if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                    return 'Reachable'
                }
                return 'No ICMP Reply'
            } catch {
                if ($phase -eq 'DNS') { return 'DNS Resolution Failed' }
                return 'Ping Error'
            } finally {
                if ($null -ne $ping) { $ping.Dispose() }
                if ($dnsSlotReserved) {
                    try { [void]$DnsTaskSlots.Release() } catch { }
                }
            }
        }

        # Bound allocated pipelines as well as active execution. Previously all
        # farm entries were queued up front, so a very large inventory could hold
        # thousands of PowerShell objects even though only 30 were running.
        $pendingNames = [System.Collections.Generic.Queue[string]]::new()
        foreach ($name in $ComputerNames) { $pendingNames.Enqueue([string]$name) }
        $totalNames = $pendingNames.Count
        $completedCount = 0

        while ($pendingNames.Count -gt 0 -or $jobs.Count -gt 0) {
            if ($CancellationToken.IsCancellationRequested) {
                throw [System.OperationCanceledException]::new('Ping reachability check cancelled.')
            }

            while ($pendingNames.Count -gt 0 -and $jobs.Count -lt $ThrottleLimit) {
                $name = $pendingNames.Dequeue()
                $ps = $null
                try {
                    $ps = [PowerShell]::Create().AddScript($scriptBlock).AddArgument($name).AddArgument($TimeoutMs).AddArgument($DnsTimeoutMs).AddArgument($DnsCleanupWaitMs).AddArgument($CancellationToken).AddArgument($dnsTaskQueue).AddArgument($dnsTaskSlots)
                    $ps.RunspacePool = $pool
                    $handle = $ps.BeginInvoke()
                    [void]$jobs.Add([PSCustomObject]@{
                        Pipe = $ps
                        Handle = $handle
                        Name = $name
                    })
                    $ps = $null
                } finally {
                    if ($null -ne $ps) { $ps.Dispose() }
                }
            }

            $completedThisPass = 0
            for ($jobIndex = $jobs.Count - 1; $jobIndex -ge 0; $jobIndex--) {
                $job = $jobs[$jobIndex]
                if ($null -eq $job.Handle -or -not $job.Handle.IsCompleted) { continue }

                try {
                    if ($CancellationToken.IsCancellationRequested) {
                        throw [System.OperationCanceledException]::new('Ping reachability check cancelled.')
                    }
                    $pingState = $job.Pipe.EndInvoke($job.Handle) | Select-Object -First 1
                    if ($pingState -eq 'Cancelled' -or $CancellationToken.IsCancellationRequested) {
                        throw [System.OperationCanceledException]::new('Ping reachability check cancelled.')
                    }
                    $results[$job.Name] = if ($pingState) { [string]$pingState } else { 'Ping Error' }
                } catch {
                    if ($CancellationToken.IsCancellationRequested -or $_.Exception -is [System.OperationCanceledException]) {
                        throw [System.OperationCanceledException]::new('Ping reachability check cancelled.')
                    }
                    $results[$job.Name] = 'Ping Error'
                } finally {
                    if ($null -ne $job.Pipe) {
                        try { $job.Pipe.Dispose() } catch { }
                        $job.Pipe = $null
                    }
                }

                $jobs.RemoveAt($jobIndex)
                $completedCount++
                $completedThisPass++
                if ($null -ne $MessageQueue) {
                    [void]$MessageQueue.Enqueue([PSCustomObject]@{
                        Kind = 'Progress'
                        Message = "Checking ping reachability: $completedCount of $totalNames"
                        Current = $completedCount
                        Total = $totalNames
                        IsIndeterminate = $false
                    })
                }
            }

            if ($completedThisPass -eq 0 -and ($pendingNames.Count -gt 0 -or $jobs.Count -gt 0)) {
                Start-Sleep -Milliseconds 25
            }
        }
    } finally {
        foreach ($job in @($jobs)) {
            if ($null -eq $job -or $null -eq $job.Pipe) { continue }
            try {
                if ($null -ne $job.Handle -and -not $job.Handle.IsCompleted) {
                    $job.Pipe.Stop()
                }
            } catch {
                # Best-effort cleanup during cancellation or runspace failure.
            } finally {
                try { $job.Pipe.Dispose() } catch { }
                $job.Pipe = $null
            }
        }
        if ($null -ne $pool) {
            try { $pool.Close() } catch { }
            try { $pool.Dispose() } catch { }
        }
        if ($null -ne $dnsTaskQueue) {
            $remainingTask = $null
            while ($dnsTaskQueue.TryDequeue([ref]$remainingTask)) {
                if ($null -ne $remainingTask -and $remainingTask.IsCompleted) {
                    try { [void]$remainingTask.Exception } catch { }
                }
                $remainingTask = $null
            }
        }
        if ($null -ne $dnsTaskSlots) { try { $dnsTaskSlots.Dispose() } catch { } }
    }
    return $results
}

# -------------------------------------------------------------
# Region: Shared WPF Styles (injected into every Window)
# -------------------------------------------------------------
# The resource dictionary is injected into both XAML windows. Keeping styles in
# one block ensures the PVS and XDC interfaces use the same visual language.
# Colour palette - Citrix-inspired dark teal / blue-grey
$script:SharedStyles = @'
  <Window.Resources>
    <!-- Colour palette -->
    <SolidColorBrush x:Key="AccentBrush"        Color="#1A5276"/>
    <SolidColorBrush x:Key="AccentLightBrush"    Color="#2980B9"/>
    <SolidColorBrush x:Key="AccentHoverBrush"    Color="#1F6FA5"/>
    <SolidColorBrush x:Key="AccentPressBrush"    Color="#154360"/>
    <SolidColorBrush x:Key="HeaderBgBrush"       Color="#1A5276"/>
    <SolidColorBrush x:Key="HeaderFgBrush"       Color="#FFFFFF"/>
    <SolidColorBrush x:Key="CardBgBrush"         Color="#FAFBFC"/>
    <SolidColorBrush x:Key="CardBorderBrush"     Color="#D5D8DC"/>
    <SolidColorBrush x:Key="BodyBgBrush"         Color="#ECF0F1"/>
    <SolidColorBrush x:Key="StatusBgBrush"       Color="#2C3E50"/>
    <SolidColorBrush x:Key="StatusFgBrush"       Color="#BDC3C7"/>
    <SolidColorBrush x:Key="GridHeaderBrush"     Color="#2C3E50"/>
    <SolidColorBrush x:Key="GridHeaderFgBrush"   Color="#FFFFFF"/>
    <SolidColorBrush x:Key="GridAltRowBrush"     Color="#EBF5FB"/>
    <SolidColorBrush x:Key="GridRowHoverBrush"   Color="#D4E6F1"/>
    <SolidColorBrush x:Key="DangerBrush"         Color="#E74C3C"/>
    <SolidColorBrush x:Key="SuccessBrush"        Color="#27AE60"/>

    <!-- Primary button style -->
    <Style x:Key="PrimaryButton" TargetType="Button">
      <Setter Property="Background"  Value="{StaticResource AccentBrush}"/>
      <Setter Property="Foreground"  Value="White"/>
      <Setter Property="FontSize"    Value="12"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Padding"     Value="16,6"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="border" Background="{TemplateBinding Background}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}"
                    BorderThickness="0">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="border" Property="Background" Value="{StaticResource AccentHoverBrush}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="border" Property="Background" Value="{StaticResource AccentPressBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="border" Property="Background" Value="#95A5A6"/>
                <Setter Property="Foreground" Value="#D5D8DC"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Secondary / outline button -->
    <Style x:Key="SecondaryButton" TargetType="Button">
      <Setter Property="Background"  Value="Transparent"/>
      <Setter Property="Foreground"  Value="{StaticResource AccentBrush}"/>
      <Setter Property="FontSize"    Value="12"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Padding"     Value="16,6"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="border" Background="{TemplateBinding Background}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}"
                    BorderBrush="{StaticResource AccentBrush}" BorderThickness="1.5">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="border" Property="Background" Value="#D4E6F1"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="border" Property="Background" Value="#AED6F1"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Close / danger button -->
    <Style x:Key="CloseButton" TargetType="Button">
      <Setter Property="Background"  Value="#95A5A6"/>
      <Setter Property="Foreground"  Value="White"/>
      <Setter Property="FontSize"    Value="12"/>
      <Setter Property="FontWeight"  Value="SemiBold"/>
      <Setter Property="Padding"     Value="16,6"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="border" Background="{TemplateBinding Background}"
                    CornerRadius="4" Padding="{TemplateBinding Padding}" BorderThickness="0">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="border" Property="Background" Value="#7F8C8D"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="border" Property="Background" Value="#616A6B"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Toggle link button (Select All) -->
    <Style x:Key="LinkButton" TargetType="Button">
      <Setter Property="Background"  Value="Transparent"/>
      <Setter Property="Foreground"  Value="{StaticResource AccentLightBrush}"/>
      <Setter Property="FontSize"    Value="11"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <TextBlock x:Name="txt" Text="{TemplateBinding Content}"
                       Foreground="{TemplateBinding Foreground}"
                       FontSize="{TemplateBinding FontSize}"
                       TextDecorations="Underline"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="txt" Property="Foreground" Value="{StaticResource AccentBrush}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Styled progress bar -->
    <Style x:Key="AccentProgress" TargetType="ProgressBar">
      <Setter Property="Height" Value="6"/>
      <Setter Property="Background" Value="#D5D8DC"/>
      <Setter Property="Foreground" Value="{StaticResource AccentLightBrush}"/>
      <Setter Property="BorderThickness" Value="0"/>
    </Style>

    <!-- Styled TextBox -->
    <Style x:Key="StyledTextBox" TargetType="TextBox">
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="BorderBrush" Value="#BDC3C7"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Style.Triggers>
        <Trigger Property="IsFocused" Value="True">
          <Setter Property="BorderBrush" Value="{StaticResource AccentLightBrush}"/>
          <Setter Property="BorderThickness" Value="1.5"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Styled CheckBox -->
    <Style TargetType="CheckBox">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>

    <!-- DataGrid column header style -->
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background"      Value="{StaticResource GridHeaderBrush}"/>
      <Setter Property="Foreground"      Value="{StaticResource GridHeaderFgBrush}"/>
      <Setter Property="FontWeight"      Value="SemiBold"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Padding"         Value="10,7"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment"   Value="Center"/>
      <Setter Property="BorderBrush"     Value="#34495E"/>
      <Setter Property="BorderThickness" Value="0,0,1,0"/>
    </Style>

    <!-- DataGrid row style -->
    <Style TargetType="DataGridRow">
      <Setter Property="FontSize" Value="12"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="{StaticResource GridRowHoverBrush}"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#AED6F1"/>
          <Setter Property="Foreground" Value="#1A5276"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- DataGrid cell style -->
    <Style TargetType="DataGridCell">
      <Setter Property="Padding" Value="8,5"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment"   Value="Center"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="Transparent"/>
          <Setter Property="Foreground" Value="#1A5276"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- Centered text style for auto-generated DataGrid text columns -->
    <Style x:Key="CenteredDataGridText" TargetType="TextBlock">
      <Setter Property="TextAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>

    <Style x:Key="CenteredDataGridEditingText" TargetType="TextBox">
      <Setter Property="TextAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
  </Window.Resources>
'@

# -------------------------------------------------------------
# Region: Shared UI Helpers
# -------------------------------------------------------------
# Role-neutral helpers for export, clipboard handling, filtering, DNS, server
# selection, summary formatting, DHCP analysis, and DataGrid presentation.
# Neutralizes values that spreadsheet applications could otherwise interpret as
# formulas. Numeric and Boolean values retain their original types; only strings
# with a formula-significant first non-space character are prefixed.
function ConvertTo-SpreadsheetSafeValue {
    param(
        $Value
    )

    if ($null -eq $Value -or $Value -isnot [string]) { return $Value }
    $text = [string]$Value
    if ($text -match '^\s*[=+\-@]') { return "'$text" }
    return $text
}

# Returns the visible columns in the same left-to-right order the operator sees.
# Both CSV and clipboard output use this helper so their headings cannot drift.
function Get-DataGridVisibleColumnDefinitions {
    param($DataGrid)

    $definitions = [System.Collections.Generic.List[object]]::new()
    foreach ($column in @(
        $DataGrid.Columns |
            Where-Object { [string]$_.Visibility -eq 'Visible' } |
            Sort-Object -Property DisplayIndex
    )) {
        $propertyName = [string]$column.SortMemberPath
        if ([string]::IsNullOrWhiteSpace($propertyName) -and
            $column.PSObject.Properties.Match('Binding').Count -gt 0 -and
            $null -ne $column.Binding -and $null -ne $column.Binding.Path) {
            $propertyName = [string]$column.Binding.Path.Path
        }
        if ([string]::IsNullOrWhiteSpace($propertyName)) {
            $propertyName = [string]$column.Header
        }

        [void]$definitions.Add([PSCustomObject]@{
            Header = [string]$column.Header
            PropertyName = $propertyName
        })
    }
    return @($definitions.ToArray())
}

# Exports the displayed (filtered and sorted) DataGrid view. Data is staged in
# the destination directory and committed only after Export-Csv succeeds so a
# failed write cannot truncate a previously valid file.
function Export-DataGridToCSV {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [string]$DefaultFileName = 'Export.csv'
    )

    $visibleRows = @($DataGrid.Items)
    if ($null -eq $DataGrid.ItemsSource -or $visibleRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'No data available to export.', 'Warning', 'OK', 'Warning')
        return
    }

    $columnDefinitions = @(Get-DataGridVisibleColumnDefinitions -DataGrid $DataGrid)
    if ($columnDefinitions.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'No visible columns are available to export.', 'Warning', 'OK', 'Warning')
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    try {
        $dialog.Filter   = 'CSV Files (*.csv)|*.csv'
        $dialog.Title    = 'Save CSV File'
        $dialog.FileName = $DefaultFileName

        if ($dialog.ShowDialog() -eq 'OK') {
            $temporaryPath = $null
            try {
                $destinationPath = [IO.Path]::GetFullPath([string]$dialog.FileName)
                $destinationDirectory = [IO.Path]::GetDirectoryName($destinationPath)
                if ([string]::IsNullOrWhiteSpace($destinationDirectory) -or
                    -not (Test-Path -LiteralPath $destinationDirectory -PathType Container)) {
                    throw "The selected destination directory does not exist: [$destinationDirectory]."
                }
                if (Test-Path -LiteralPath $destinationPath -PathType Container) {
                    throw "The selected CSV destination is a directory: [$destinationPath]."
                }
                if (Test-Path -LiteralPath $destinationPath -PathType Leaf) {
                    $destinationInfo = Get-Item -LiteralPath $destinationPath -Force -ErrorAction Stop
                    if (($destinationInfo.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                        throw 'The selected CSV destination cannot be a symbolic link or reparse point.'
                    }
                }

                $temporaryName = '.{0}.{1}.tmp' -f [IO.Path]::GetFileName($destinationPath), [guid]::NewGuid().ToString('N')
                $temporaryPath = Join-Path $destinationDirectory $temporaryName
                $safeRows = @(
                    foreach ($row in $visibleRows) {
                        $safeRow = [ordered]@{}
                        foreach ($definition in $columnDefinitions) {
                            $cellValue = $null
                            if ($row.PSObject.Properties.Match($definition.PropertyName).Count -gt 0) {
                                $cellValue = $row.($definition.PropertyName)
                            }
                            $safeRow[[string]$definition.Header] = ConvertTo-SpreadsheetSafeValue -Value $cellValue
                        }
                        [PSCustomObject]$safeRow
                    }
                )
                $safeRows | Export-Csv -LiteralPath $temporaryPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                if (Test-Path -LiteralPath $destinationPath -PathType Leaf) {
                    [IO.File]::Replace($temporaryPath, $destinationPath, $null)
                } else {
                    [IO.File]::Move($temporaryPath, $destinationPath)
                }
                $temporaryPath = $null
                [System.Windows.MessageBox]::Show(
                    "Exported $($visibleRows.Count) displayed records to:`n$destinationPath",
                    'Export Complete', 'OK', 'Information')
                Write-Log "Exported displayed data | Rows=$($visibleRows.Count) | Columns=$($columnDefinitions.Count) | Path=[$destinationPath]"
            } catch {
                [System.Windows.MessageBox]::Show(
                    "Export failed: $($_.Exception.Message)", 'Error', 'OK', 'Error')
                Write-Log "Export failed | $(Get-ErrorDiagnosticText -ErrorRecord $_ -Phase 'CSV export')" -Level ERROR
            } finally {
                if (-not [string]::IsNullOrWhiteSpace([string]$temporaryPath) -and
                    (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) {
                    try { [IO.File]::Delete($temporaryPath) } catch {
                        Write-Log "Unable to remove failed CSV staging file [$temporaryPath]: $($_.Exception.Message)" -Level WARN
                    }
                }
            }
        }
    } finally {
        try { $dialog.Dispose() } catch { }
    }
}

# Writes text to the Windows clipboard. Kept separate for focused validation/mocking.
function Set-TextClipboard {
    param(
        [string]$Text
    )
    [System.Windows.Clipboard]::SetText($Text)
}

# Copies server names from selected output rows, or all displayed rows when none are selected.
function Copy-ServerListFromGrid {
    param(
        $DataGrid,
        [string]$PropertyName,
        [string]$FieldDisplayName,
        $StatusText,
        [string]$Context
    )

    $visibleRows = @($DataGrid.Items)
    if ($visibleRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'No output rows are available to copy.',
            'Nothing to Copy', 'OK', 'Information') | Out-Null
        return
    }

    $hasSelection = ($DataGrid.SelectedItems.Count -gt 0)
    $rowsToCopy = if ($hasSelection) {
        @($visibleRows | Where-Object { $DataGrid.SelectedItems.Contains($_) })
    } else {
        $visibleRows
    }

    $serverNames = [System.Collections.Generic.List[string]]::new()
    foreach ($row in @($rowsToCopy)) {
        if ($row.PSObject.Properties.Match($PropertyName).Count -eq 0) { continue }
        $serverName = [string]$row.$PropertyName
        if (-not [string]::IsNullOrWhiteSpace($serverName)) {
            [void]$serverNames.Add($serverName)
        }
    }

    if ($serverNames.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Server names are not available in the current output.`n`nSelect the '$FieldDisplayName' data field and retrieve the data again.",
            'Server Name Field Not Available', 'OK', 'Warning') | Out-Null
        return
    }

    try {
        Set-TextClipboard -Text ($serverNames -join [Environment]::NewLine)
        $sourceText = if ($hasSelection) { 'selected' } else { 'displayed' }
        $nameLabel = if ($serverNames.Count -eq 1) { 'server name' } else { 'server names' }
        $StatusText.Text = "Copied $($serverNames.Count) $nameLabel from $sourceText rows"
        Write-Log "$Context copied $($serverNames.Count) server names to the clipboard from $sourceText output rows"
    } catch {
        Write-Log "$Context clipboard copy failed: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show(
            "Unable to copy the server list to the clipboard.`n`nError: $($_.Exception.Message)",
            'Clipboard Error', 'OK', 'Error') | Out-Null
    }
}

# Converts a displayed grid value into a single clipboard-safe TSV cell.
function ConvertTo-ClipboardCellText {
    param(
        $Value
    )

    if ($null -eq $Value) { return '' }
    $text = ([regex]::Replace([string]$Value, '[\r\n\t]+', ' ')).Trim()
    return [string](ConvertTo-SpreadsheetSafeValue -Value $text)
}

# Copies the complete displayed grid, including headings, in tab-separated format.
function Copy-DataGridTable {
    param(
        $DataGrid,
        $StatusText,
        [string]$Context
    )

    $visibleRows = @($DataGrid.Items)
    if ($visibleRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'No output rows are available to copy.',
            'Nothing to Copy', 'OK', 'Information') | Out-Null
        return
    }

    $columnDefinitions = @(Get-DataGridVisibleColumnDefinitions -DataGrid $DataGrid)
    if ($columnDefinitions.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'No visible output columns are available to copy.',
            'Nothing to Copy', 'OK', 'Information') | Out-Null
        return
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    [void]$lines.Add((@($columnDefinitions | ForEach-Object { ConvertTo-ClipboardCellText -Value $_.Header }) -join "`t"))

    foreach ($row in $visibleRows) {
        $cells = [System.Collections.Generic.List[string]]::new()
        foreach ($definition in $columnDefinitions) {
            $cellValue = $null
            if ($row.PSObject.Properties.Match($definition.PropertyName).Count -gt 0) {
                $cellValue = $row.($definition.PropertyName)
            }
            [void]$cells.Add((ConvertTo-ClipboardCellText -Value $cellValue))
        }
        [void]$lines.Add(($cells -join "`t"))
    }

    try {
        Set-TextClipboard -Text ($lines -join [Environment]::NewLine)
        $StatusText.Text = "Copied table with $($visibleRows.Count) rows and $($columnDefinitions.Count) headings"
        Write-Log "$Context copied displayed table to the clipboard: Rows=$($visibleRows.Count), Columns=$($columnDefinitions.Count), Headings=Included"
    } catch {
        Write-Log "$Context table clipboard copy failed: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show(
            "Unable to copy the output table to the clipboard.`n`nError: $($_.Exception.Message)",
            'Clipboard Error', 'OK', 'Error') | Out-Null
    }
}

# Updates status bar text and optional record count text.
function Update-StatusBar {
    param(
        [System.Windows.Controls.TextBlock]$StatusText,
        [string]$Message
    )
    $StatusText.Dispatcher.Invoke([action]{
        $StatusText.Text = $Message
    })
}

# Forces a WPF dispatcher cycle so UI updates render immediately.
function Invoke-UiRefresh {
    param(
        [System.Windows.Window]$Window
    )
    if ($null -eq $Window) { return }
    $Window.UpdateLayout()
    $Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
}

# Formats a stopwatch duration without wrapping total hours after 24 hours.
function Format-ExecutionTime {
    param(
        [TimeSpan]$Elapsed
    )

    $totalHours = [int][Math]::Floor($Elapsed.TotalHours)
    return ('{0:00}:{1:00}:{2:00}.{3:000}' -f $totalHours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds)
}

<#
.SYNOPSIS
    Applies search filtering to a dataset with optional field-specific syntax.
.DESCRIPTION
    Supports:
    - Plain text: searches all visible fields
    - Field filter: "FieldName:Value" (example: PingReachability:Reachable)
#>
# Applies free-text or Field:Value filtering to grid data.
function Apply-GridFilter {
    param(
        [object[]]$Data,
        [string]$Term
    )

    $allRows = @($Data)
    if ($allRows.Count -eq 0) {
        return [PSCustomObject]@{
            Rows = @()
            Summary = '0 records'
        }
    }

    $trimmed = if ($null -eq $Term) { '' } else { $Term.Trim() }
    if (-not $trimmed) {
        return [PSCustomObject]@{
            Rows = $allRows
            Summary = "$($allRows.Count) records"
        }
    }

    $fieldName = $null
    $fieldValue = $null
    if ($trimmed -match '^\s*([^:]+?)\s*:\s*(.*)\s*$') {
        $fieldName = $matches[1].Trim()
        $fieldValue = $matches[2].Trim()
    }

    if ($fieldName -and $fieldValue) {
        $fieldMap = @{}
        foreach ($prop in $allRows[0].PSObject.Properties.Name) {
            $fieldMap[$prop.ToLowerInvariant()] = $prop
        }
        $resolvedField = $fieldMap[$fieldName.ToLowerInvariant()]
        if ($resolvedField) {
            $filtered = $allRows | Where-Object {
                $value = $_.$resolvedField
                $null -ne $value -and ([string]$value).IndexOf($fieldValue, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            }
            return [PSCustomObject]@{
                Rows = @($filtered)
                Summary = "$(@($filtered).Count) / $($allRows.Count) records (${resolvedField}:`"$fieldValue`")"
            }
        }
    }

    $globalFiltered = $allRows | Where-Object {
        $match = $false
        foreach ($prop in $_.PSObject.Properties) {
            if ($null -ne $prop.Value -and ([string]$prop.Value).IndexOf($trimmed, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $match = $true; break }
        }
        $match
    }

    return [PSCustomObject]@{
        Rows = @($globalFiltered)
        Summary = "$(@($globalFiltered).Count) / $($allRows.Count) records"
    }
}

# Resolves an IPv4 address without allowing a slow DNS lookup to hang retrieval indefinitely.
function Resolve-IPv4Address {
    param(
        [string]$Name,
        [int]$TimeoutMs = 1500,
        [object]$MessageQueue,
        [datetime]$DeadlineUtc = [datetime]::MaxValue,
        [System.Threading.CancellationToken]$CancellationToken = [System.Threading.CancellationToken]::None
    )

    $writeResolverLog = {
        param([string]$Message, [string]$Level)
        if ($null -ne $MessageQueue) {
            [void]$MessageQueue.Enqueue([PSCustomObject]@{
                Kind = 'Log'
                Message = $Message
                Level = $Level
            })
        } else {
            Write-Log -Message $Message -Level $Level
        }
    }

    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    if ($CancellationToken.IsCancellationRequested) {
        throw [System.OperationCanceledException]::new('DNS lookup cancelled.')
    }
    $remainingDeadlineMs = [Math]::Floor(($DeadlineUtc - [datetime]::UtcNow).TotalMilliseconds)
    if ($remainingDeadlineMs -lt 1) { return $null }
    $effectiveTimeoutMs = [int][Math]::Min([double]$TimeoutMs, [double]$remainingDeadlineMs)

    # Task-based DNS does not require the APM End* cleanup that a timed-out
    # BeginGetHostAddresses call previously missed. Keep a bounded list of tasks
    # that outlive the caller timeout so a DNS outage cannot create hundreds of
    # outstanding native resolver requests in one worker runspace.
    if ($null -eq $script:OutstandingDnsLookupTasks) {
        $script:OutstandingDnsLookupTasks = [System.Collections.Generic.List[object]]::new()
    }
    for ($index = $script:OutstandingDnsLookupTasks.Count - 1; $index -ge 0; $index--) {
        $pendingTask = $script:OutstandingDnsLookupTasks[$index]
        if (-not $pendingTask.IsCompleted) { continue }
        try {
            if ($pendingTask.IsFaulted) { $null = $pendingTask.Exception }
            $pendingTask.Dispose()
        } catch { }
        $script:OutstandingDnsLookupTasks.RemoveAt($index)
    }
    $maximumOutstandingDnsTasks = 30
    if ($script:OutstandingDnsLookupTasks.Count -ge $maximumOutstandingDnsTasks) {
        if (-not $script:DnsBacklogWarningIssued) {
            $script:DnsBacklogWarningIssued = $true
            & $writeResolverLog "DNS fallback paused because $maximumOutstandingDnsTasks prior timed-out lookups are still pending. Additional names will be reported unresolved until the backlog clears." 'WARN'
        }
        return $null
    }

    $dnsTask = $null
    try {
        $dnsTask = [System.Net.Dns]::GetHostAddressesAsync($Name)
        if (-not $dnsTask.Wait($effectiveTimeoutMs, $CancellationToken)) {
            & $writeResolverLog "DNS lookup timed out for [$Name] after ${effectiveTimeoutMs}ms" 'WARN'
            [void]$script:OutstandingDnsLookupTasks.Add($dnsTask)
            $dnsTask = $null
            return $null
        }

        if ($CancellationToken.IsCancellationRequested) {
            throw [System.OperationCanceledException]::new('DNS lookup cancelled.')
        }

        $resolved = @($dnsTask.Result) |
            Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
            Select-Object -First 1

        if ($resolved) { return $resolved.IPAddressToString }
    } catch [System.OperationCanceledException] {
        if ($null -ne $dnsTask -and -not $dnsTask.IsCompleted -and
            $script:OutstandingDnsLookupTasks.Count -lt $maximumOutstandingDnsTasks) {
            [void]$script:OutstandingDnsLookupTasks.Add($dnsTask)
            $dnsTask = $null
        }
        throw
    } catch {
        if ($CancellationToken.IsCancellationRequested -or $_.Exception -is [System.OperationCanceledException]) {
            if ($null -ne $dnsTask -and -not $dnsTask.IsCompleted -and
                $script:OutstandingDnsLookupTasks.Count -lt $maximumOutstandingDnsTasks) {
                [void]$script:OutstandingDnsLookupTasks.Add($dnsTask)
                $dnsTask = $null
            }
            throw [System.OperationCanceledException]::new('DNS lookup cancelled.')
        }
        & $writeResolverLog "DNS lookup failed for [$Name]: $($_.Exception.Message)" 'WARN'
    } finally {
        if ($null -ne $dnsTask -and $dnsTask.IsCompleted) {
            try { $dnsTask.Dispose() } catch { }
        }
    }

    return $null
}

<#
.SYNOPSIS
    Retrieves the complete Broker machine inventory with SDK-compatible paging.
.DESCRIPTION
    Uses -Skip in PageSize batches when that parameter exists. Older Broker SDK
    versions without -Skip fall back to a high MaxRecordCount value and record a
    warning because the SDK cannot provide deterministic paging in that mode.
.OUTPUTS
    Citrix Broker machine objects.
#>
function Get-BrokerMachineInventory {
    param(
        [ValidateRange(100, 5000)]
        [int]$PageSize = 1000,
        [ValidateRange(1000, 1000000)]
        [int]$MaximumRecordCount = 100000,
        [object]$MessageQueue,
        [System.Threading.CancellationToken]$CancellationToken = [System.Threading.CancellationToken]::None
    )

    if ($CancellationToken.IsCancellationRequested) {
        throw [System.OperationCanceledException]::new('Broker inventory retrieval cancelled.')
    }

    $cmd = Get-Command Get-BrokerMachine -ErrorAction Stop
    if ($cmd.Parameters.ContainsKey('Skip')) {
        $machines = [System.Collections.Generic.List[object]]::new()
        $seenMachineKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $missingIdentityCount = 0
        $skip = 0
        $supportsStableSort = $cmd.Parameters.ContainsKey('SortBy')
        if (-not $supportsStableSort) {
            $warning = 'Get-BrokerMachine supports paging but not -SortBy; duplicate-page validation is enabled, but concurrent site changes can still make the snapshot inconsistent.'
            if ($null -ne $MessageQueue) {
                [void]$MessageQueue.Enqueue([PSCustomObject]@{ Kind = 'Log'; Message = $warning; Level = 'WARN' })
            } else { Write-Log $warning -Level WARN }
        }

        do {
            if ($CancellationToken.IsCancellationRequested) {
                throw [System.OperationCanceledException]::new('Broker inventory retrieval cancelled.')
            }
            if ($skip -ge $MaximumRecordCount) {
                throw "Broker inventory reached the $MaximumRecordCount-record safety limit before a final partial page proved completeness. Narrow the scope or raise the reviewed limit."
            }
            $queryParameters = @{
                MaxRecordCount = $PageSize
                Skip = $skip
                ErrorAction = 'Stop'
            }
            if ($supportsStableSort) { $queryParameters.SortBy = 'Uid' }
            $batch = @(Get-BrokerMachine @queryParameters)
            if ($machines.Count + $batch.Count -gt $MaximumRecordCount) {
                throw "Broker inventory exceeds the $MaximumRecordCount-record safety limit. Narrow the scope or raise the reviewed limit."
            }
            foreach ($machine in $batch) {
                $machineKey = ''
                foreach ($propertyName in @('Uid','MachineName','DNSName')) {
                    if ($machine.PSObject.Properties.Match($propertyName).Count -gt 0 -and
                        -not [string]::IsNullOrWhiteSpace([string]$machine.$propertyName)) {
                        $machineKey = "${propertyName}:$([string]$machine.$propertyName)"
                        break
                    }
                }
                if ([string]::IsNullOrWhiteSpace($machineKey)) {
                    $missingIdentityCount++
                } elseif (-not $seenMachineKeys.Add($machineKey)) {
                    throw "Broker paging returned duplicate machine identity [$machineKey]. The site changed during paging or the SDK ignored Skip; no potentially incomplete table will be published. Retry the retrieval."
                }
                [void]$machines.Add($machine)
            }
            $skip += $batch.Count
            if ($null -ne $MessageQueue) {
                [void]$MessageQueue.Enqueue([PSCustomObject]@{
                    Kind = 'Progress'
                    Message = "Retrieved $($machines.Count) Broker machines..."
                    Current = $machines.Count
                    Total = 0
                    IsIndeterminate = $true
                })
            }
        } while ($batch.Count -eq $PageSize)

        if ($missingIdentityCount -gt 0) {
            $warning = "$missingIdentityCount Broker record(s) lacked Uid, MachineName, and DNSName, so duplicate-page validation could not cover those records."
            if ($null -ne $MessageQueue) {
                [void]$MessageQueue.Enqueue([PSCustomObject]@{ Kind = 'Log'; Message = $warning; Level = 'WARN' })
            } else { Write-Log $warning -Level WARN }
        }

        return @($machines)
    }

    if ($null -ne $MessageQueue) {
        [void]$MessageQueue.Enqueue([PSCustomObject]@{
            Kind = 'Log'
            Message = 'Get-BrokerMachine paging parameter (-Skip) is unavailable; using high MaxRecordCount fallback.'
            Level = 'WARN'
        })
    } else {
        Write-Log 'Get-BrokerMachine paging parameter (-Skip) is unavailable; using high MaxRecordCount fallback.' -Level WARN
    }
    if ($CancellationToken.IsCancellationRequested) {
        throw [System.OperationCanceledException]::new('Broker inventory retrieval cancelled.')
    }
    $fallbackMachines = @(Get-BrokerMachine -MaxRecordCount $MaximumRecordCount -ErrorAction Stop)
    if ($fallbackMachines.Count -ge $MaximumRecordCount) {
        throw "The legacy Broker query returned the $MaximumRecordCount-record cap, so completeness cannot be proven without -Skip paging. No potentially truncated table will be published."
    }
    return $fallbackMachines
}

<#
.SYNOPSIS
    Converts pasted text into an ordered, unique list of requested server names.
.DESCRIPTION
    Splits on new lines, commas, or semicolons. Matching is case-insensitive for
    duplicate removal, but qualified names remain distinct until they are
    resolved against the Citrix inventory.
.OUTPUTS
    System.String[]
#>
function ConvertTo-ServerNameList {
    param(
        [AllowEmptyString()]
        [string]$InputText,
        [ValidateRange(1, 1048576)]
        [int]$MaximumCharacters = 262144,
        [ValidateRange(1, 10000)]
        [int]$MaximumNames = 5000,
        [ValidateRange(1, 1024)]
        [int]$MaximumNameLength = 512
    )

    $names = [System.Collections.Generic.List[string]]::new()
    $seenNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if ([string]::IsNullOrWhiteSpace($InputText)) { return @() }
    if ($InputText.Length -gt $MaximumCharacters) {
        throw "Server input exceeds the $MaximumCharacters-character safety limit. Use smaller Server List batches."
    }

    foreach ($value in [regex]::Split($InputText, '[,;\r\n]+')) {
        $trimmed = $value.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        if ($trimmed.Length -gt $MaximumNameLength) {
            throw "A server name exceeds the $MaximumNameLength-character safety limit."
        }

        # Keep qualified names distinct; alias equivalence is resolved against inventory later.
        $normalizedName = $trimmed.TrimEnd('.')
        if ($seenNames.Add($normalizedName)) {
            [void]$names.Add($trimmed)
            if ($names.Count -gt $MaximumNames) {
                throw "Server input exceeds the $MaximumNames-name safety limit. Use smaller Server List batches."
            }
        }
    }

    return @($names)
}

<#
.SYNOPSIS
    Builds the aliases used to compare user input with Citrix inventory names.
.DESCRIPTION
    Returns the supplied value, the value without a DOMAIN\ prefix, and the
    short host portion of an FQDN. The caller decides whether an alias is unique.
.OUTPUTS
    System.String[]
#>
function Get-ServerNameAliases {
    param(
        [AllowEmptyCollection()]
        [string[]]$Names
    )

    $aliases = [System.Collections.Generic.List[string]]::new()
    $seenAliases = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($name in @($Names)) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $trimmed = $name.Trim().TrimEnd('.')
        if ($seenAliases.Add($trimmed)) { [void]$aliases.Add($trimmed) }

        $withoutDomain = $trimmed -replace '^.*\\', ''
        if ($seenAliases.Add($withoutDomain)) { [void]$aliases.Add($withoutDomain) }

        if ($withoutDomain.Contains('.')) {
            $shortName = $withoutDomain.Split('.')[0]
            if ($seenAliases.Add($shortName)) { [void]$aliases.Add($shortName) }
        }
    }

    return @($aliases)
}

<#
.SYNOPSIS
    Matches requested server names to Citrix inventory records.
.DESCRIPTION
    Builds exact-name and alias lookup tables once, then processes user input in
    request order. Exact matches take priority. An alias is accepted only when
    it resolves to one record; otherwise it is reported as ambiguous. A record
    already selected through another alias is reported as a duplicate.
.PARAMETER Items
    The full PVS device or Broker machine inventory.
.PARAMETER ServerNames
    The ordered names parsed from the user's input.
.PARAMETER NameProperties
    Inventory properties that can identify a record, such as Name, MachineName,
    or DNSName.
.OUTPUTS
    PSCustomObject with Items, MissingNames, AmbiguousNames, and DuplicateNames.
#>
function Select-ItemsByServerList {
    param(
        [object[]]$Items,
        [string[]]$ServerNames,
        [string[]]$NameProperties
    )

    $exactLookup = @{}
    $aliasLookup = @{}

    $addLookupItem = {
        param(
            [hashtable]$Lookup,
            [string]$Key,
            [object]$Item
        )

        if ([string]::IsNullOrWhiteSpace($Key)) { return }
        if (-not $Lookup.ContainsKey($Key)) {
            $Lookup[$Key] = [System.Collections.Generic.List[object]]::new()
        }

        $matches = [System.Collections.Generic.List[object]]$Lookup[$Key]
        if (-not $matches.Contains($Item)) {
            [void]$matches.Add($Item)
        }
    }

    foreach ($item in @($Items)) {
        foreach ($propertyName in $NameProperties) {
            if ($item.PSObject.Properties.Match($propertyName).Count -gt 0 -and $item.$propertyName) {
                $exactName = ([string]$item.$propertyName).Trim().TrimEnd('.')
                & $addLookupItem -Lookup $exactLookup -Key $exactName -Item $item

                foreach ($alias in @(Get-ServerNameAliases -Names $exactName)) {
                    & $addLookupItem -Lookup $aliasLookup -Key $alias -Item $item
                }
            }
        }
    }

    $selectedItems = [System.Collections.Generic.List[object]]::new()
    $missingNames = [System.Collections.Generic.List[string]]::new()
    $ambiguousNames = [System.Collections.Generic.List[string]]::new()
    $duplicateNames = [System.Collections.Generic.List[string]]::new()
    $addedItems = [System.Collections.Generic.HashSet[object]]::new()

    foreach ($serverName in @($ServerNames)) {
        $candidateMatches = [System.Collections.Generic.HashSet[object]]::new()
        $exactName = $serverName.Trim().TrimEnd('.')

        # An exact DOMAIN\HOST, FQDN, or inventory-name match always takes priority.
        if ($exactLookup.ContainsKey($exactName)) {
            foreach ($candidate in @($exactLookup[$exactName])) {
                [void]$candidateMatches.Add($candidate)
            }
        } else {
            foreach ($alias in @(Get-ServerNameAliases -Names $serverName)) {
                if ($aliasLookup.ContainsKey($alias)) {
                    foreach ($candidate in @($aliasLookup[$alias])) {
                        [void]$candidateMatches.Add($candidate)
                    }
                }
            }
        }

        if ($candidateMatches.Count -eq 0) {
            [void]$missingNames.Add($serverName)
        } elseif ($candidateMatches.Count -gt 1) {
            [void]$ambiguousNames.Add($serverName)
        } else {
            $matchedItem = @($candidateMatches)[0]
            if ($addedItems.Add($matchedItem)) {
                [void]$selectedItems.Add($matchedItem)
            } else {
                [void]$duplicateNames.Add($serverName)
            }
        }
    }

    return [PSCustomObject]@{
        Items = @($selectedItems)
        MissingNames = @($missingNames)
        AmbiguousNames = @($ambiguousNames)
        DuplicateNames = @($duplicateNames)
    }
}

# Builds a stable collection identity from the PVS device object. CollectionId
# is preferred because collection names can repeat in different PVS sites. Older
# SDK shapes that omit the ID fall back to SiteId/SiteName plus CollectionName.
function Get-PvsDeviceCollectionChoiceKey {
    param(
        [object]$Device
    )

    if ($null -eq $Device) { return '' }
    $collectionName = if ($Device.PSObject.Properties.Match('CollectionName').Count -gt 0) { ([string]$Device.CollectionName).Trim() } else { '' }
    if ([string]::IsNullOrWhiteSpace($collectionName)) { return '' }

    $collectionId = if ($Device.PSObject.Properties.Match('CollectionId').Count -gt 0) { ([string]$Device.CollectionId).Trim() } else { '' }
    $parsedCollectionId = [guid]::Empty
    if ([guid]::TryParse($collectionId, [ref]$parsedCollectionId) -and $parsedCollectionId -ne [guid]::Empty) {
        return 'ID:' + $parsedCollectionId.ToString('D').ToUpperInvariant()
    }

    $normalizedCollectionName = $collectionName.ToUpperInvariant()
    $siteId = if ($Device.PSObject.Properties.Match('SiteId').Count -gt 0) { ([string]$Device.SiteId).Trim() } else { '' }
    $parsedSiteId = [guid]::Empty
    if ([guid]::TryParse($siteId, [ref]$parsedSiteId) -and $parsedSiteId -ne [guid]::Empty) {
        # Length-prefix the free-text component so embedded delimiters/control
        # characters cannot make two distinct collection tuples share a key.
        return 'SITEID:' + $parsedSiteId.ToString('D').ToUpperInvariant() + '|COLLECTION:' + $normalizedCollectionName.Length + ':' + $normalizedCollectionName
    }

    $siteName = if ($Device.PSObject.Properties.Match('SiteName').Count -gt 0) { ([string]$Device.SiteName).Trim() } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($siteName)) {
        $normalizedSiteName = $siteName.ToUpperInvariant()
        return 'SITENAME:' + $normalizedSiteName.Length + ':' + $normalizedSiteName + '|COLLECTION:' + $normalizedCollectionName.Length + ':' + $normalizedCollectionName
    }

    # A collection name alone is not a safe farm-wide identity because the same
    # name may exist in more than one PVS site. Exclude that row from the picker
    # instead of silently merging unrelated collections.
    return ''
}

# Derives collection choices from a supplied PVS farm inventory without issuing
# another query inside this helper. Duplicate names in different sites remain
# separate choices.
function Get-PvsDeviceCollectionChoices {
    param(
        [object[]]$Devices
    )

    $choiceMap = @{}
    foreach ($device in @($Devices)) {
        $choiceKey = Get-PvsDeviceCollectionChoiceKey -Device $device
        if ([string]::IsNullOrWhiteSpace($choiceKey)) { continue }

        if (-not $choiceMap.ContainsKey($choiceKey)) {
            $collectionName = ([string]$device.CollectionName).Trim()
            $siteName = if ($device.PSObject.Properties.Match('SiteName').Count -gt 0) { ([string]$device.SiteName).Trim() } else { '' }
            $rawSiteId = if ($device.PSObject.Properties.Match('SiteId').Count -gt 0) { ([string]$device.SiteId).Trim() } else { '' }
            $parsedSiteId = [guid]::Empty
            $siteId = if ([guid]::TryParse($rawSiteId, [ref]$parsedSiteId) -and $parsedSiteId -ne [guid]::Empty) { $parsedSiteId.ToString('D') } else { '' }
            $rawCollectionId = if ($device.PSObject.Properties.Match('CollectionId').Count -gt 0) { ([string]$device.CollectionId).Trim() } else { '' }
            $parsedCollectionId = [guid]::Empty
            $collectionId = if ([guid]::TryParse($rawCollectionId, [ref]$parsedCollectionId) -and $parsedCollectionId -ne [guid]::Empty) { $parsedCollectionId.ToString('D') } else { '' }
            $choiceMap[$choiceKey] = [PSCustomObject]@{
                ChoiceKey = $choiceKey
                CollectionId = $collectionId
                CollectionName = $collectionName
                SiteId = $siteId
                SiteName = $siteName
                DeviceCount = 0
                LocationName = ''
                DisplayName = ''
            }
        }
        $choiceMap[$choiceKey].DeviceCount = [int]$choiceMap[$choiceKey].DeviceCount + 1
    }

    $choices = @($choiceMap.Values)
    $locationCounts = @{}
    foreach ($choice in $choices) {
        $siteLabel = if (-not [string]::IsNullOrWhiteSpace([string]$choice.SiteName)) {
            [string]$choice.SiteName
        } elseif (-not [string]::IsNullOrWhiteSpace([string]$choice.SiteId)) {
            "Site $($choice.SiteId)"
        } else {
            ''
        }
        $choice.LocationName = if ([string]::IsNullOrWhiteSpace($siteLabel)) {
            [string]$choice.CollectionName
        } else {
            "$siteLabel \ $($choice.CollectionName)"
        }
        $locationKey = ([string]$choice.LocationName).ToUpperInvariant()
        $locationCounts[$locationKey] = if ($locationCounts.ContainsKey($locationKey)) { [int]$locationCounts[$locationKey] + 1 } else { 1 }
    }

    foreach ($choice in $choices) {
        $location = [string]$choice.LocationName
        if ([int]$locationCounts[$location.ToUpperInvariant()] -gt 1) {
            $identityLabel = if (-not [string]::IsNullOrWhiteSpace([string]$choice.CollectionId)) {
                "Collection ID: $($choice.CollectionId)"
            } else {
                "Site ID: $($choice.SiteId)"
            }
            $location = "$location [$identityLabel]"
        }
        $deviceLabel = if ([int]$choice.DeviceCount -eq 1) { 'device' } else { 'devices' }
        $choice.DisplayName = "$location ($($choice.DeviceCount) $deviceLabel)"
    }
    return @($choices | Sort-Object SiteName, SiteId, CollectionName, CollectionId)
}

# Filters a supplied farm inventory locally by the opaque choice key returned by
# Get-PvsDeviceCollectionChoices. Inventory order is preserved.
function Select-PvsDevicesByCollection {
    param(
        [object[]]$Devices,
        [string]$ChoiceKey
    )

    if ([string]::IsNullOrWhiteSpace($ChoiceKey)) { return @() }
    return @(
        @($Devices) | Where-Object {
            (Get-PvsDeviceCollectionChoiceKey -Device $_) -ieq $ChoiceKey
        }
    )
}

# Extracts one XDC Machine Catalog or Delivery Group identity from a Broker
# machine. A valid Broker UUID or positive UID is required so incomplete SDK
# metadata can never merge unrelated same-named groups.
function Get-DdcNamedScopeRecord {
    param(
        [object]$Machine,
        [ValidateSet('MachineCatalog','DeliveryGroup')]
        [string]$ScopeType
    )

    $scopePrefix = if ($ScopeType -eq 'MachineCatalog') { 'MC' } else { 'DG' }
    $nameProperty = if ($ScopeType -eq 'MachineCatalog') { 'CatalogName' } else { 'DesktopGroupName' }
    $uuidProperty = if ($ScopeType -eq 'MachineCatalog') { 'CatalogUUID' } else { 'DesktopGroupUUID' }
    $uidProperty = if ($ScopeType -eq 'MachineCatalog') { 'CatalogUid' } else { 'DesktopGroupUid' }

    $scopeName = ''
    if ($null -ne $Machine) {
        $nameMatch = @($Machine.PSObject.Properties.Match($nameProperty))
        if ($nameMatch.Count -gt 0 -and $null -ne $nameMatch[0].Value) {
            $scopeName = ([string]$nameMatch[0].Value).Trim()
        }
    }
    $normalizedName = if ([string]::IsNullOrWhiteSpace($scopeName)) { '' } else { $scopeName.ToUpperInvariant() }

    $uuidText = ''
    $uuidMatch = if ($null -eq $Machine) { @() } else { @($Machine.PSObject.Properties.Match($uuidProperty)) }
    if ($uuidMatch.Count -gt 0 -and $null -ne $uuidMatch[0].Value) {
        $uuidText = ([string]$uuidMatch[0].Value).Trim()
    }
    $uidText = ''
    $uidMatch = if ($null -eq $Machine) { @() } else { @($Machine.PSObject.Properties.Match($uidProperty)) }
    if ($uidMatch.Count -gt 0 -and $null -ne $uidMatch[0].Value) {
        $uidText = ([string]$uidMatch[0].Value).Trim()
    }
    $hasIdentifierMetadata = -not [string]::IsNullOrWhiteSpace($uuidText) -or -not [string]::IsNullOrWhiteSpace($uidText)

    $parsedUuid = [guid]::Empty
    if (-not [string]::IsNullOrWhiteSpace($scopeName) -and [guid]::TryParse($uuidText, [ref]$parsedUuid) -and $parsedUuid -ne [guid]::Empty) {
        $normalizedUuid = $parsedUuid.ToString('D').ToUpperInvariant()
        return [PSCustomObject]@{
            ScopeName = $scopeName
            NormalizedName = $normalizedName
            ExplicitChoiceKey = "$scopePrefix|UUID:$normalizedUuid"
            IdentifierType = 'UUID'
            IdentifierValue = $parsedUuid.ToString('D')
            IsUnassigned = $false
            HasIdentifierMetadata = $true
        }
    }

    $parsedUid = 0
    if (-not [string]::IsNullOrWhiteSpace($scopeName) -and [int]::TryParse($uidText, [ref]$parsedUid) -and $parsedUid -gt 0) {
        return [PSCustomObject]@{
            ScopeName = $scopeName
            NormalizedName = $normalizedName
            ExplicitChoiceKey = "$scopePrefix|UID:$parsedUid"
            IdentifierType = 'UID'
            IdentifierValue = [string]$parsedUid
            IsUnassigned = $false
            HasIdentifierMetadata = $true
        }
    }

    return [PSCustomObject]@{
        ScopeName = $scopeName
        NormalizedName = $normalizedName
        ExplicitChoiceKey = ''
        IdentifierType = ''
        IdentifierValue = ''
        IsUnassigned = ($ScopeType -eq 'DeliveryGroup' -and [string]::IsNullOrWhiteSpace($scopeName) -and -not $hasIdentifierMetadata)
        HasIdentifierMetadata = $hasIdentifierMetadata
    }
}

# Builds XDC picker choices and a per-inventory-row membership map from a supplied
# Broker inventory. Unassigned Delivery Group rows are tracked separately from
# malformed or incomplete identity metadata; neither enters a picker.
function Get-DdcNamedScopeAnalysis {
    param(
        [object[]]$Machines,
        [ValidateSet('MachineCatalog','DeliveryGroup')]
        [string]$ScopeType
    )

    $inventory = @($Machines)
    $choiceMap = @{}
    $membershipKeys = [System.Collections.Generic.List[string]]::new()
    $unassignedMachineCount = 0
    $unsafeMetadataMachineCount = 0
    foreach ($machine in $inventory) {
        $record = Get-DdcNamedScopeRecord -Machine $machine -ScopeType $ScopeType
        if ([bool]$record.IsUnassigned) {
            $unassignedMachineCount++
            [void]$membershipKeys.Add('')
            continue
        }
        $resolvedKey = [string]$record.ExplicitChoiceKey
        if ([string]::IsNullOrWhiteSpace([string]$record.ScopeName) -or [string]::IsNullOrWhiteSpace($resolvedKey)) {
            $unsafeMetadataMachineCount++
            [void]$membershipKeys.Add('')
            continue
        }

        [void]$membershipKeys.Add($resolvedKey)
        if (-not $choiceMap.ContainsKey($resolvedKey)) {
            $choiceMap[$resolvedKey] = [PSCustomObject]@{
                ChoiceKey = $resolvedKey
                ScopeType = $ScopeType
                ScopeName = [string]$record.ScopeName
                IdentifierType = [string]$record.IdentifierType
                IdentifierValue = [string]$record.IdentifierValue
                MachineCount = 0
                DisplayName = ''
            }
        }
        $choiceMap[$resolvedKey].MachineCount = [int]$choiceMap[$resolvedKey].MachineCount + 1
    }

    $choices = @($choiceMap.Values)
    $choiceNameCounts = @{}
    foreach ($choice in $choices) {
        $nameKey = ([string]$choice.ScopeName).ToUpperInvariant()
        $choiceNameCounts[$nameKey] = if ($choiceNameCounts.ContainsKey($nameKey)) { [int]$choiceNameCounts[$nameKey] + 1 } else { 1 }
    }

    $scopeLabel = if ($ScopeType -eq 'MachineCatalog') { 'Catalog' } else { 'Delivery Group' }
    foreach ($choice in $choices) {
        $displayName = [string]$choice.ScopeName
        if ([int]$choiceNameCounts[$displayName.ToUpperInvariant()] -gt 1) {
            $displayName += " [$scopeLabel $($choice.IdentifierType): $($choice.IdentifierValue)]"
        }
        $machineLabel = if ([int]$choice.MachineCount -eq 1) { 'machine' } else { 'machines' }
        $choice.DisplayName = "$displayName ($($choice.MachineCount) $machineLabel)"
    }

    return [PSCustomObject]@{
        Choices = @($choices | Sort-Object ScopeName, IdentifierType, IdentifierValue)
        MembershipKeys = @($membershipKeys.ToArray())
        UnassignedMachineCount = $unassignedMachineCount
        UnsafeMetadataMachineCount = $unsafeMetadataMachineCount
        UnselectableMachineCount = $unassignedMachineCount + $unsafeMetadataMachineCount
    }
}

# Filters the original Broker inventory by an opaque picker key while preserving
# the order returned by Get-BrokerMachineInventory.
function Select-DdcMachinesByNamedScope {
    param(
        [object[]]$Machines,
        [ValidateSet('MachineCatalog','DeliveryGroup')]
        [string]$ScopeType,
        [string]$ChoiceKey
    )

    if ([string]::IsNullOrWhiteSpace($ChoiceKey)) { return @() }
    $inventory = @($Machines)
    $analysis = Get-DdcNamedScopeAnalysis -Machines $inventory -ScopeType $ScopeType
    $selected = [System.Collections.Generic.List[object]]::new()
    for ($index = 0; $index -lt $inventory.Count; $index++) {
        if ([System.StringComparer]::OrdinalIgnoreCase.Equals([string]$analysis.MembershipKeys[$index], $ChoiceKey)) {
            [void]$selected.Add($inventory[$index])
        }
    }
    return @($selected.ToArray())
}

# Keeps message-box and summary text readable when a long server list has misses.
function Format-ServerNamePreview {
    param(
        [string[]]$Names,
        [int]$MaximumNames = 20
    )

    $nameList = @($Names)
    if ($nameList.Count -eq 0) { return 'None' }

    $preview = @($nameList | Select-Object -First $MaximumNames) -join ', '
    if ($nameList.Count -gt $MaximumNames) {
        $preview += " (+$($nameList.Count - $MaximumNames) more)"
    }
    return $preview
}

# Formats unresolved and duplicate list entries for warnings and persistent summaries.
function Format-ServerSelectionIssues {
    param(
        [string[]]$MissingNames,
        [string[]]$AmbiguousNames,
        [string[]]$DuplicateNames
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    if (@($MissingNames).Count -gt 0) {
        [void]$lines.Add("Not found: $(Format-ServerNamePreview -Names $MissingNames)")
    }
    if (@($AmbiguousNames).Count -gt 0) {
        [void]$lines.Add("Ambiguous (use DOMAIN\SERVER or FQDN): $(Format-ServerNamePreview -Names $AmbiguousNames)")
    }
    if (@($DuplicateNames).Count -gt 0) {
        [void]$lines.Add("Duplicate aliases ignored: $(Format-ServerNamePreview -Names $DuplicateNames)")
    }

    return ($lines -join "`r`n")
}

# Builds DDC health totals for maintenance, registration, and logon state.
function Get-DdcHealthSummary {
    param(
        [object[]]$Machines
    )

    $machineList = @($Machines)
    if ($machineList.Count -eq 0) {
        return [PSCustomObject]@{
            MaintenanceCount = 0
            UnregisteredCount = 0
            LogonDisabledCount = 0
            LogonUnavailableCount = 0
        }
    }

    $maintenanceCount = @($machineList | Where-Object { $_.InMaintenanceMode -eq $true }).Count
    $unregisteredCount = @($machineList | Where-Object { [string]$_.RegistrationState -eq 'Unregistered' }).Count

    $logonDisabledCount = 0
    $logonUnavailableCount = 0
    foreach ($machine in $machineList) {
        if ($machine.PSObject.Properties.Match('WindowsConnectionSetting').Count -gt 0 -and $null -ne $machine.WindowsConnectionSetting) {
            if ([string]$machine.WindowsConnectionSetting -ne 'LogonEnabled') {
                $logonDisabledCount++
            }
        } else {
            $logonUnavailableCount++
        }
    }

    return [PSCustomObject]@{
        MaintenanceCount = $maintenanceCount
        UnregisteredCount = $unregisteredCount
        LogonDisabledCount = $logonDisabledCount
        LogonUnavailableCount = $logonUnavailableCount
    }
}

# Produces grouped count summaries for a selected property.
function Get-GroupedCountSummary {
    param(
        [object[]]$Items,
        [string]$PropertyName,
        [string]$Label
    )

    $itemList = @($Items)
    if ($itemList.Count -eq 0) {
        return [PSCustomObject]@{
            Inline = "${Label}: none"
            Detail = "${Label}: none"
        }
    }

    $groups = @(
        $itemList |
            Group-Object -Property $PropertyName |
            Sort-Object -Property Name
    )

    if ($groups.Count -eq 0) {
        return [PSCustomObject]@{
            Inline = "${Label}: none"
            Detail = "${Label}: none"
        }
    }

    $pairs = foreach ($group in $groups) {
        $groupName = [string]$group.Name
        if ([string]::IsNullOrWhiteSpace($groupName)) { $groupName = 'N/A' }
        "$groupName=$($group.Count)"
    }

    return [PSCustomObject]@{
        Inline = "${Label}: " + ($pairs -join ', ')
        Detail = "${Label}: " + ($pairs -join ' | ')
    }
}

# -------------------------------------------------------------
# Region: Background Citrix Retrieval
# -------------------------------------------------------------
# Normal PVS and XDC work runs in one dedicated runspace. Citrix calls remain
# serial inside that runspace; the purpose is UI responsiveness and cooperative
# cancellation, not higher query concurrency against the shared farm/site.
$script:PvsWorkloadConfirmationThreshold = 1000
# Inline choices must represent a reasonably current farm/site preview. The UI
# refreshes old previews after this bounded age; Retrieve still performs one
# fresh inventory read so membership and live values represent execution time.
$script:NamedScopeSnapshotMaximumAgeMinutes = 5

# Makes one required Citrix command available in the current runspace. Implicit
# module auto-loading stays disabled: already loaded commands are source-checked,
# then the exact registered snap-in or a protected machine-installed module is
# loaded explicitly. This is deliberately command-specific; the presence of an
# unrelated Citrix snap-in must not be mistaken for SDK readiness.
function Test-TrustedInstallPath {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    try {
        $candidate = [IO.Path]::GetFullPath($Path)
        # Derive machine-owned installation roots from operating-system APIs.
        # Process environment variables are intentionally excluded because a
        # caller can replace ProgramFiles/SystemRoot before starting PowerShell.
        $trustedRoots = [System.Collections.Generic.List[string]]::new()
        foreach ($specialFolder in @(
            [Environment+SpecialFolder]::ProgramFiles,
            [Environment+SpecialFolder]::ProgramFilesX86,
            [Environment+SpecialFolder]::CommonProgramFiles,
            [Environment+SpecialFolder]::CommonProgramFilesX86
        )) {
            $folderPath = [Environment]::GetFolderPath($specialFolder)
            if (-not [string]::IsNullOrWhiteSpace($folderPath)) {
                [void]$trustedRoots.Add($folderPath)
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($PSHOME)) {
            [void]$trustedRoots.Add($PSHOME)
        }

        # A registered snap-in assembly can be loaded from the protected GAC.
        # Trust only the two GAC roots, never the full Windows directory (which
        # contains writable locations such as Windows\Temp).
        $windowsPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::Windows)
        if (-not [string]::IsNullOrWhiteSpace($windowsPath)) {
            [void]$trustedRoots.Add((Join-Path $windowsPath 'assembly'))
            [void]$trustedRoots.Add((Join-Path $windowsPath 'Microsoft.NET\assembly'))
        }

        foreach ($root in @($trustedRoots | Select-Object -Unique)) {
            $fullRoot = [IO.Path]::GetFullPath([string]$root).TrimEnd([char[]]@('\','/'))
            if ($candidate -ieq $fullRoot) { return $true }
            $rootWithSeparator = $fullRoot + [IO.Path]::DirectorySeparatorChar
            if ($candidate.StartsWith($rootWithSeparator, [StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
    } catch { }
    return $false
}

# CWxPVS is an environment-specific module wrapper that can expose genuine PVS
# cmdlets without preserving the Citrix.PVS.SnapIn name in CommandInfo. Accept
# that one wrapper only when the cmdlet is implemented by the exact Citrix PVS
# console DLL installed under the OS-derived Program Files root. This verifies
# the binary origin rather than trusting a module name from the current session.
function Test-TrustedCwxPvsCmdletAssembly {
    param(
        [System.Management.Automation.CommandInfo]$Command
    )

    if ($null -eq $Command -or [string]$Command.CommandType -ne 'Cmdlet' -or
        $null -eq $Command.ImplementingType) {
        return $false
    }

    try {
        $assembly = $Command.ImplementingType.Assembly
        if ($null -eq $assembly -or [string]$assembly.GetName().Name -ine 'Citrix.PVS.SnapIn') {
            return $false
        }

        $programFiles = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)
        if ([string]::IsNullOrWhiteSpace($programFiles)) { return $false }

        $expectedAssemblyPath = [IO.Path]::GetFullPath(
            (Join-Path $programFiles 'Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll'))
        $actualAssemblyPath = [IO.Path]::GetFullPath([string]$assembly.Location)
        if ($actualAssemblyPath -ine $expectedAssemblyPath) { return $false }

        return (Test-TrustedInstallPath -Path $actualAssemblyPath)
    } catch {
        return $false
    }
}

function Test-TrustedCitrixCommandInfo {
    param(
        [System.Management.Automation.CommandInfo]$Command,
        [ValidateSet('Get-PVSDevice','Get-PVSDiskInfo','Get-PVSDevicePersonality','Get-PvsServer','Get-BrokerMachine')]
        [string]$CommandName
    )

    if ($null -eq $Command -or $Command.Name -ine $CommandName -or
        [string]$Command.CommandType -notin @('Cmdlet','Function')) {
        return $false
    }

    $rule = if ($CommandName -in @('Get-PVSDevice','Get-PVSDiskInfo','Get-PVSDevicePersonality','Get-PvsServer')) {
        @{ SnapInName = 'Citrix.PVS.SnapIn'; ModulePattern = '^Citrix\.PVS(\.|$)' }
    } else {
        @{ SnapInName = 'Citrix.Broker.Admin.V2'; ModulePattern = '^Citrix\.Broker(\.|$)' }
    }

    $snapInName = ''
    if ($Command.PSObject.Properties.Match('PSSnapIn').Count -gt 0 -and $null -ne $Command.PSSnapIn) {
        $snapInName = [string]$Command.PSSnapIn.Name
    }
    if ([string]$Command.CommandType -eq 'Cmdlet' -and $snapInName -ieq $rule.SnapInName) {
        # An exact registered name is not sufficient by itself. Validate the
        # implementing assembly (or registered module path) against a protected
        # machine installation root before accepting the snap-in command.
        $snapInSourcePath = ''
        if ($null -ne $Command.ImplementingType) {
            $snapInSourcePath = [string]$Command.ImplementingType.Assembly.Location
        }
        if ([string]::IsNullOrWhiteSpace($snapInSourcePath) -and
            $Command.PSSnapIn.PSObject.Properties.Match('ModuleName').Count -gt 0 -and
            -not [string]::IsNullOrWhiteSpace([string]$Command.PSSnapIn.ModuleName)) {
            $registeredModule = [string]$Command.PSSnapIn.ModuleName
            if ([IO.Path]::IsPathRooted($registeredModule)) {
                $snapInSourcePath = $registeredModule
            } elseif ($Command.PSSnapIn.PSObject.Properties.Match('ApplicationBase').Count -gt 0 -and
                      -not [string]::IsNullOrWhiteSpace([string]$Command.PSSnapIn.ApplicationBase)) {
                $snapInSourcePath = Join-Path ([string]$Command.PSSnapIn.ApplicationBase) $registeredModule
            }
        }
        return (Test-TrustedInstallPath -Path $snapInSourcePath)
    }

    $moduleName = [string]$Command.ModuleName
    if ([string]::IsNullOrWhiteSpace($moduleName)) { $moduleName = [string]$Command.Source }

    # Some supported PVS hosts expose the canonical Citrix PVS cmdlets through
    # CWxPVS. The wrapper name alone is never trusted; its implementing assembly
    # must pass the exact path and assembly-identity checks above.
    if ($CommandName -in @('Get-PVSDevice','Get-PVSDiskInfo','Get-PVSDevicePersonality','Get-PvsServer') -and
        $moduleName -ieq 'CWxPVS') {
        return (Test-TrustedCwxPvsCmdletAssembly -Command $Command)
    }

    if ($moduleName -notmatch $rule.ModulePattern) { return $false }

    $sourcePath = ''
    if ($null -ne $Command.Module -and $Command.Module.Path) {
        $sourcePath = [string]$Command.Module.Path
    } elseif ($null -ne $Command.ImplementingType) {
        $sourcePath = [string]$Command.ImplementingType.Assembly.Location
    }
    return (Test-TrustedInstallPath -Path $sourcePath)
}

function Initialize-CitrixWorkerCommand {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Get-PVSDevice','Get-PVSDiskInfo','Get-PVSDevicePersonality','Get-PvsServer','Get-BrokerMachine')]
        [string]$CommandName
    )

    $script:TrustedCitrixWorkerCommandInfo = $null
    # Prevent exact-name discovery from auto-importing a user-writable module
    # before its source can be validated.
    $PSModuleAutoLoadingPreference = 'None'

    $rule = if ($CommandName -in @('Get-PVSDevice','Get-PVSDiskInfo','Get-PVSDevicePersonality','Get-PvsServer')) {
        @{ SnapInName = 'Citrix.PVS.SnapIn'; ModulePattern = '^Citrix\.PVS(\.|$)' }
    } else {
        @{ SnapInName = 'Citrix.Broker.Admin.V2'; ModulePattern = '^Citrix\.Broker(\.|$)' }
    }

    $setTrustedMarker = {
        param([System.Management.Automation.CommandInfo]$Command)
        $script:TrustedCitrixWorkerCommandInfo = [PSCustomObject]@{
            CommandName = [string]$Command.Name
            CommandType = [string]$Command.CommandType
            Source = [string]$Command.Source
            ModuleName = [string]$Command.ModuleName
            ModulePath = if ($null -ne $Command.Module) { [string]$Command.Module.Path } else { '' }
            SnapInName = if ($Command.PSObject.Properties.Match('PSSnapIn').Count -gt 0 -and $null -ne $Command.PSSnapIn) { [string]$Command.PSSnapIn.Name } else { '' }
        }
    }

    $resolved = @(Get-Command -Name $CommandName -All -ErrorAction SilentlyContinue)
    if ($resolved.Count -gt 0) {
        if (Test-TrustedCitrixCommandInfo -Command $resolved[0] -CommandName $CommandName) {
            & $setTrustedMarker $resolved[0]
            return
        }
        throw "Command [$CommandName] is shadowed by an untrusted source: Type=$($resolved[0].CommandType), Source=$($resolved[0].Source). Run the tool from a dedicated powershell.exe -NoProfile process and remove the conflicting command."
    }

    $loadErrors = [System.Collections.Generic.List[string]]::new()
    try {
        $snapIn = Get-PSSnapin -Registered -Name $rule.SnapInName -ErrorAction Stop | Select-Object -First 1
        if ($null -ne $snapIn) {
            $isLoaded = @(Get-PSSnapin -Name $snapIn.Name -ErrorAction SilentlyContinue).Count -gt 0
            if (-not $isLoaded) { Add-PSSnapin -Name $snapIn.Name -ErrorAction Stop }
        }
    } catch {
        [void]$loadErrors.Add("Snap-in [$($rule.SnapInName)]: $($_.Exception.Message)")
    }

    $resolved = @(Get-Command -Name $CommandName -All -ErrorAction SilentlyContinue)
    if ($resolved.Count -gt 0) {
        if (Test-TrustedCitrixCommandInfo -Command $resolved[0] -CommandName $CommandName) {
            & $setTrustedMarker $resolved[0]
            return
        }
        throw "Command [$CommandName] resolved to an untrusted source after loading snap-in [$($rule.SnapInName)]: $($resolved[0].Source)."
    }

    $candidateModules = @()
    try {
        $candidateModules = @(
            Get-Module -ListAvailable -ErrorAction Stop |
                Where-Object {
                    $_.Name -match $rule.ModulePattern -and
                    (Test-TrustedInstallPath -Path ([string]$_.Path)) -and
                    ($null -eq $_.ExportedCommands -or $_.ExportedCommands.Count -eq 0 -or $_.ExportedCommands.ContainsKey($CommandName))
                } |
                Sort-Object Version -Descending
        )
    } catch {
        $loadErrors.Add("Citrix module discovery failed: $($_.Exception.Message)") | Out-Null
    }

    foreach ($module in $candidateModules) {
        try {
            Import-Module -Name ([string]$module.Path) -ErrorAction Stop
        } catch {
            $loadErrors.Add("Module [$($module.Name)]: $($_.Exception.Message)") | Out-Null
            continue
        }

        $resolved = @(Get-Command -Name $CommandName -All -ErrorAction SilentlyContinue)
        if ($resolved.Count -gt 0) {
            if (Test-TrustedCitrixCommandInfo -Command $resolved[0] -CommandName $CommandName) {
                & $setTrustedMarker $resolved[0]
                return
            }
            throw "Command [$CommandName] resolved to an untrusted source after module loading: $($resolved[0].Source)."
        }
    }

    $errorSuffix = if ($loadErrors.Count -gt 0) {
        ' Load attempts: ' + ($loadErrors -join ' | ')
    } else {
        ' No approved Citrix snap-in or trusted machine-installed module exposed the command.'
    }
    throw "Required Citrix command [$CommandName] is unavailable in this runspace.$errorSuffix"
}

function Send-RetrievalWorkerMessage {
    param(
        [object]$MessageQueue,
        [ValidateSet('Progress','Log','WorkloadConfirmation')]
        [string]$Kind,
        [string]$Message,
        [int]$Current = 0,
        [int]$Total = 0,
        [bool]$IsIndeterminate = $false,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [hashtable]$Details,
        [string]$Component = 'RetrievalWorker',
        [string]$Operation = '',
        [string]$Role = '',
        [string]$Retryable = '',
        [string]$SuggestedAction = ''
    )

    if ($null -eq $MessageQueue) { return }
    if ([string]::IsNullOrWhiteSpace($Role)) { $Role = [string]$script:CurrentRetrievalWorkerRole }
    if ($Component -eq 'RetrievalWorker' -and -not [string]::IsNullOrWhiteSpace($Role)) { $Component = "$Role Retrieval" }
    if ([string]::IsNullOrWhiteSpace($Operation)) { $Operation = $Kind }
    [void]$MessageQueue.Enqueue([PSCustomObject]@{
        Kind = $Kind
        Message = $Message
        Current = $Current
        Total = $Total
        IsIndeterminate = $IsIndeterminate
        Level = $Level
        Details = $Details
        Component = $Component
        Operation = $Operation
        Role = $Role
        Retryable = $Retryable
        SuggestedAction = $SuggestedAction
    })
}

# Writes queued worker log entries through one contract, keeping PVS and XDC
# diagnostics comparable without changing their user-facing progress text.
function Write-RetrievalWorkerMessageLog {
    param(
        [object]$Message,
        [object]$Context
    )

    if ($null -eq $Message) { return }
    $workerRunId = if ($null -ne $Context -and $Context.PSObject.Properties.Match('RunId').Count -gt 0) { [string]$Context.RunId } else { '' }
    $role = if (-not [string]::IsNullOrWhiteSpace([string]$Message.Role)) { [string]$Message.Role } elseif ($null -ne $Context) { [string]$Context.Role } else { '' }
    Write-Log -Message ([string]$Message.Message) -Level ([string]$Message.Level) -Component ([string]$Message.Component) -Operation ([string]$Message.Operation) -Role $role -WorkerRunId $workerRunId -Retryable ([string]$Message.Retryable) -SuggestedAction ([string]$Message.SuggestedAction)
}

function Test-RetrievalWorkerCancellation {
    param(
        [System.Threading.CancellationToken]$CancellationToken,
        [string]$Message = 'Retrieval cancelled by the user.'
    )

    if ($CancellationToken.IsCancellationRequested) {
        throw [System.OperationCanceledException]::new($Message)
    }
}

# Returns one consistent failed-worker envelope without changing successful or
# cancelled result contracts. The UI can retain previous output while logs state
# the failing phase and the practical next action.
function New-RetrievalWorkerFailureResult {
    param(
        [ValidateSet('PVS','DDC')]
        [string]$Role,
        [object]$ErrorRecord,
        [string]$Phase,
        [string]$SuggestedAction = 'Review the execution log, correct the reported prerequisite or input issue, then retry the retrieval.'
    )

    return [PSCustomObject]@{
        Kind = 'Result'
        Role = $Role
        Success = $false
        Cancelled = $false
        ErrorMessage = [string]$ErrorRecord.Exception.Message
        ErrorDetail = Get-ErrorDiagnosticText -ErrorRecord $ErrorRecord -Phase $Phase
        FailurePhase = $Phase
        Retryable = $true
        SuggestedAction = $SuggestedAction
    }
}

function Invoke-PvsRetrievalWorker {
    param(
        [hashtable]$Request,
        [object]$MessageQueue,
        [System.Threading.CancellationToken]$CancellationToken,
        [System.Threading.ManualResetEventSlim]$ApprovalGate,
        [hashtable]$ApprovalState
    )

    try {
        $script:CurrentRetrievalWorkerRole = 'PVS'
        $scopeInventoryOnly = [bool]$Request.ScopeInventoryOnly
        $wholeFarmMode = [bool]$Request.WholeFarmMode
        $deviceCollectionMode = [bool]$Request.DeviceCollectionMode
        $singleServerMode = [bool]$Request.SingleServerMode
        $serverListMode = [bool]$Request.ServerListMode
        $calculatedSelectedServerMode = $singleServerMode -or $serverListMode
        $selectedCollectionChoiceKey = [string]$Request.SelectedCollectionChoiceKey
        $selectedCollectionKeySupplied = -not [string]::IsNullOrWhiteSpace($selectedCollectionChoiceKey)
        if (-not $scopeInventoryOnly) {
            # Fail closed: a malformed request must never fall through to the
            # default full-farm dataset and silently broaden a targeted query.
            $activeScopeCount = [int]$wholeFarmMode + [int]$deviceCollectionMode + [int]$singleServerMode + [int]$serverListMode
            if ($activeScopeCount -ne 1) {
                throw 'Invalid PVS retrieval scope state. Select exactly one scope.'
            }
            if ([bool]$Request.SelectedServerMode -ne $calculatedSelectedServerMode) {
                throw 'The PVS selected-server scope state is inconsistent with the Single Server and Server List flags.'
            }
            if ($deviceCollectionMode -xor $selectedCollectionKeySupplied) {
                throw 'The PVS Device Collection scope state is inconsistent with the selected collection key.'
            }

            $requestedNames = @($Request.RequestedServers)
            if ($calculatedSelectedServerMode) {
                if ($requestedNames.Count -eq 0 -or $requestedNames.Count -gt 5000) {
                    throw 'The PVS selected-server request must contain between 1 and 5,000 names.'
                }
                if ($singleServerMode -and $requestedNames.Count -ne 1) {
                    throw 'The PVS Single Server request must contain exactly one name.'
                }
                foreach ($requestedName in $requestedNames) {
                    if ([string]::IsNullOrWhiteSpace([string]$requestedName) -or ([string]$requestedName).Length -gt 512) {
                        throw 'The PVS selected-server request contains an empty or oversized name.'
                    }
                }
            } elseif ($requestedNames.Count -gt 0) {
                throw 'Server names were supplied for a PVS scope that does not accept server-name input.'
            }

            $dhcpEnabled = $null -ne $Request.Dhcp -and [bool]$Request.Dhcp.Enabled
            if ([bool]$Request.WantDhcpCheck -ne $dhcpEnabled) {
                throw 'The PVS DHCP checkbox state is inconsistent with the DHCP request envelope.'
            }
            if ($dhcpEnabled) {
                $expectedRetrievalScope = if ($wholeFarmMode) { 'WholeFarm' } elseif ($deviceCollectionMode) { 'DeviceCollection' } elseif ($singleServerMode) { 'SingleServer' } else { 'ServerList' }
                if ([string]$Request.Dhcp.RetrievalScope -ne $expectedRetrievalScope) {
                    throw 'The DHCP retrieval scope is inconsistent with the selected PVS scope.'
                }
                $dhcpMode = [string]$Request.Dhcp.Mode
                if (($wholeFarmMode -and $dhcpMode -notin @('Baseline','DeepAudit')) -or
                    (-not $wholeFarmMode -and $dhcpMode -ne 'Targeted')) {
                    throw 'The DHCP validation mode is not valid for the selected PVS scope.'
                }
            }
        }

        $cacheMaterialSupplied = -not [string]::IsNullOrWhiteSpace([string]$Request.InventorySnapshotId) -or
            ($Request.ContainsKey('CachedInventory') -and @($Request.CachedInventory).Count -gt 0)
        if (-not $scopeInventoryOnly -and $cacheMaterialSupplied) {
            throw 'Cached PVS inventory data is not accepted for retrieval. Retrieve uses a fresh farm inventory and only the selected opaque collection key.'
        }

        # Keep implicit module auto-loading disabled throughout this worker.
        # Every Citrix command used by a selected phase is resolved and source-
        # validated explicitly before the first invocation.
        $PSModuleAutoLoadingPreference = 'None'
        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
        Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message 'Loading the Citrix PVS SDK...' -IsIndeterminate $true
        Initialize-CitrixWorkerCommand -CommandName 'Get-PVSDevice'
        if (-not $scopeInventoryOnly) {
            if ([bool]$Request.WantVDisk) {
                Initialize-CitrixWorkerCommand -CommandName 'Get-PVSDiskInfo'
            }
            if ([bool]$Request.WantReboot -or [bool]$Request.WantCSR -or [bool]$Request.WantXDC) {
                Initialize-CitrixWorkerCommand -CommandName 'Get-PVSDevicePersonality'
            }
        }

        Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message 'Retrieving the shared PVS farm inventory...' -IsIndeterminate $true
        Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message $(if ($scopeInventoryOnly) { 'Retrieving PVS devices for the inline Device Collection list' } else { 'Retrieving fresh PVS devices for requested output' })
        $allDevices = @(Get-PVSDevice -ErrorAction Stop)
        if ($allDevices.Count -gt 100000) {
            throw 'PVS inventory exceeds the 100,000-record production safety limit. No table was published; review the farm scope before raising this limit.'
        }
        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken

        # Radio-button selection preloads only the collection choices. Retrieve
        # later sends the opaque key and performs its own current farm read so
        # membership, IP, and DHCP inputs represent execution time.
        if ($scopeInventoryOnly) {
            # Return plain choice-source records rather than proprietary SDK
            # objects. Records without a device name are unusable and excluded.
            $missingNameDeviceCount = 0
            $snapshotDevices = @(
                foreach ($device in $allDevices) {
                    $deviceName = [string]$device.Name
                    if ([string]::IsNullOrWhiteSpace($deviceName)) {
                        $missingNameDeviceCount++
                        continue
                    }
                    [PSCustomObject]@{
                        Name = $deviceName
                        CollectionName = [string]$device.CollectionName
                        CollectionId = if ($device.PSObject.Properties.Match('CollectionId').Count -gt 0) { [string]$device.CollectionId } else { '' }
                        SiteId = if ($device.PSObject.Properties.Match('SiteId').Count -gt 0) { [string]$device.SiteId } else { '' }
                        SiteName = if ($device.PSObject.Properties.Match('SiteName').Count -gt 0) { [string]$device.SiteName } else { '' }
                    }
                }
            )
            if ($missingNameDeviceCount -gt 0) {
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Level WARN -Message "$missingNameDeviceCount PVS inventory record(s) without a device name were excluded from the inline collection list."
            }
            $collectionChoices = @(Get-PvsDeviceCollectionChoices -Devices $snapshotDevices)
            $selectableDeviceCount = [int](($collectionChoices | Measure-Object -Property DeviceCount -Sum).Sum)
            $unselectableDeviceCount = [Math]::Max(0, $snapshotDevices.Count - $selectableDeviceCount)
            if ($unselectableDeviceCount -gt 0) {
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Level WARN -Message "$unselectableDeviceCount PVS inventory record(s) were excluded from the inline collection list because CollectionName or a safe collection/site identity was unavailable."
            }
            return [PSCustomObject]@{
                Kind = 'Result'
                Role = 'PVS'
                Success = $true
                Cancelled = $false
                ErrorMessage = ''
                InventoryOnly = $true
                InventorySnapshotId = [guid]::NewGuid().ToString('N')
                Inventory = @($snapshotDevices)
                InventoryCount = $snapshotDevices.Count
                MissingNameDeviceCount = $missingNameDeviceCount
                Collections = @($collectionChoices)
                UnselectableDeviceCount = $unselectableDeviceCount
            }
        }

        $devices = $allDevices
        $selectedCollectionChoice = $null
        $missingServers = @()
        $ambiguousServers = @()
        $duplicateServers = @()
        if ([bool]$Request.DeviceCollectionMode) {
            if ($allDevices.Count -eq 0) {
                throw 'No PVS devices were returned by the shared farm inventory, so no Device Collection can be selected.'
            }
            $namedScopeDevices = @($allDevices | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Name) })
            $missingCurrentDeviceNameCount = $allDevices.Count - $namedScopeDevices.Count
            if ($missingCurrentDeviceNameCount -gt 0) {
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Level WARN -Message "$missingCurrentDeviceNameCount PVS inventory record(s) without a device name were excluded from the selected Device Collection retrieval."
            }
            $collectionChoices = @(Get-PvsDeviceCollectionChoices -Devices $namedScopeDevices)
            $selectableDeviceCount = [int](($collectionChoices | Measure-Object -Property DeviceCount -Sum).Sum)
            $unselectableDeviceCount = [Math]::Max(0, $namedScopeDevices.Count - $selectableDeviceCount)
            if ($unselectableDeviceCount -gt 0) {
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Level WARN -Message "$unselectableDeviceCount PVS inventory record(s) were excluded from Device Collection retrieval because CollectionName or a safe collection/site identity was unavailable."
            }
            if ($collectionChoices.Count -eq 0) {
                throw 'No selectable PVS device collections were returned. The inventory did not provide CollectionName together with CollectionId, SiteId, or SiteName.'
            }

            $selectedCollectionChoice = @($collectionChoices | Where-Object { $_.ChoiceKey -ieq $selectedCollectionChoiceKey } | Select-Object -First 1)
            if ($selectedCollectionChoice.Count -ne 1) {
                throw 'The selected PVS device collection was not present in the retrieved farm inventory.'
            }
            $selectedCollectionChoice = $selectedCollectionChoice[0]
            $devices = @(Select-PvsDevicesByCollection -Devices $namedScopeDevices -ChoiceKey $selectedCollectionChoiceKey)
            Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "Selected PVS device collection [$($selectedCollectionChoice.DisplayName)] with $($devices.Count) devices"
        } elseif ([bool]$Request.SelectedServerMode) {
            $scopeResult = Select-ItemsByServerList -Items $allDevices -ServerNames @($Request.RequestedServers) -NameProperties @('Name')
            $devices = @($scopeResult.Items)
            $missingServers = @($scopeResult.MissingNames)
            $ambiguousServers = @($scopeResult.AmbiguousNames)
            $duplicateServers = @($scopeResult.DuplicateNames)
        }

        $totalDevices = @($devices).Count
        Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message "Matched $totalDevices PVS devices." -Current 0 -Total $totalDevices -IsIndeterminate $false

        if ($totalDevices -gt 0) {
            $personalitySelected = [bool]$Request.WantReboot -or [bool]$Request.WantCSR -or [bool]$Request.WantXDC
            $vDiskCalls = if ([bool]$Request.WantVDisk) { $totalDevices } else { 0 }
            $personalityCalls = if ($personalitySelected) { $totalDevices } else { 0 }
            $pingAttempts = if ([bool]$Request.WantPingReachability) { $totalDevices } else { 0 }
            $maximumDnsFallbacks = if ([bool]$Request.WantServerIP) { $totalDevices } else { 0 }
            $estimatedOptionalOperations = $vDiskCalls + $personalityCalls + $pingAttempts + $maximumDnsFallbacks

            if ($estimatedOptionalOperations -ge [int]$Request.WorkloadConfirmationThreshold) {
                $ApprovalState.Approved = $false
                $ApprovalGate.Reset()
                $workloadDetails = @{
                    DeviceCount = $totalDevices
                    VDiskCalls = $vDiskCalls
                    PersonalityCalls = $personalityCalls
                    PingAttempts = $pingAttempts
                    MaximumDnsFallbacks = $maximumDnsFallbacks
                    AdditionalPvsSdkCalls = ($vDiskCalls + $personalityCalls)
                    EstimatedOptionalOperations = $estimatedOptionalOperations
                }
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind WorkloadConfirmation -Message 'Large PVS workload confirmation required.' -Details $workloadDetails

                while (-not $ApprovalGate.Wait(100)) {
                    Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken -Message 'Retrieval cancelled before optional PVS processing.'
                }
                if (-not [bool]$ApprovalState.Approved) {
                    throw [System.OperationCanceledException]::new('Retrieval cancelled before optional PVS processing.')
                }
            }
        }

        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken

        $pingResults = @{}
        if ([bool]$Request.WantPingReachability -and $totalDevices -gt 0) {
            Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message "Checking ping reachability for $totalDevices devices..." -Current 0 -Total $totalDevices -IsIndeterminate $false
            $pingResults = Get-BulkPingResults -ComputerNames @($devices | ForEach-Object { $_.Name }) -MessageQueue $MessageQueue -CancellationToken $CancellationToken

            $reachableCount = @($pingResults.Values | Where-Object { $_ -eq 'Reachable' }).Count
            $noReplyCount = @($pingResults.Values | Where-Object { $_ -eq 'No ICMP Reply' }).Count
            $dnsFailedCount = @($pingResults.Values | Where-Object { $_ -eq 'DNS Resolution Failed' }).Count
            $pingErrorCount = @($pingResults.Values | Where-Object { $_ -eq 'Ping Error' }).Count
            Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "Ping reachability complete: Reachable=$reachableCount, NoReply=$noReplyCount, DnsFailed=$dnsFailedCount, Errors=$pingErrorCount"
        }

        $deviceInfo = [System.Collections.Generic.List[object]]::new($totalDevices)
        $sourceItems = [System.Collections.Generic.List[object]]::new($totalDevices)
        $dnsCache = @{}
        $warningCount = 0
        $counter = 0

        foreach ($device in $devices) {
            Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
            $counter++
            $serverName = [string]$device.Name
            $deviceCollection = [string]$device.CollectionName
            Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message "Processing PVS device $counter of $totalDevices..." -Current $counter -Total $totalDevices -IsIndeterminate $false

            $vDiskValue = $rebootValue = $csrValue = $xdcValue = $serverIpValue = 'N/A'
            if ([bool]$Request.WantVDisk) {
                try {
                    $diskInfo = @(Get-PVSDiskInfo -DeviceName $serverName -ErrorAction Stop)
                    if ($diskInfo) {
                        $vDiskValue = ($diskInfo | ForEach-Object { $_.Name }) -join ', '
                    }
                } catch {
                    $warningCount++
                    Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "vDisk fetch failed for ${serverName}: $_" -Level WARN
                }
            }

            Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
            if ([bool]$Request.WantReboot -or [bool]$Request.WantCSR -or [bool]$Request.WantXDC) {
                try {
                    $personality = @(Get-PVSDevicePersonality -DeviceName $serverName -ErrorAction Stop |
                        Select-Object -ExpandProperty DevicePersonality)
                } catch {
                    $personality = @()
                    $warningCount++
                    Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "Personality fetch failed for ${serverName}: $_" -Level WARN
                }

                if ($personality) {
                    if ([bool]$Request.WantReboot) {
                        $value = ($personality | Where-Object { $_.Name -eq 'Reboot' }).Value
                        if ($value) { $rebootValue = $value -join ', ' }
                    }
                    if ([bool]$Request.WantCSR) {
                        $value = ($personality | Where-Object { $_.Name -eq 'CSAServer' }).Value
                        if ($value) { $csrValue = $value -join ', ' }
                    }
                    if ([bool]$Request.WantXDC) {
                        $value = ($personality | Where-Object { $_.Name -eq 'XDC_LIST' }).Value
                        if ($value) { $xdcValue = $value -join ', ' }
                    }
                }
            }

            Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
            if ([bool]$Request.WantServerIP) {
                $serverIp = $null
                $cacheKey = $serverName.ToLowerInvariant()
                foreach ($propertyName in @('IPAddress','IpAddress','IP','DeviceIP','DeviceIPAddress')) {
                    if ($device.PSObject.Properties.Match($propertyName).Count -gt 0 -and $device.$propertyName) {
                        $rawValue = $device.$propertyName
                        if ($rawValue -is [System.Array]) {
                            $serverIp = @($rawValue | Where-Object { $_ })[0]
                        } else {
                            $serverIp = $rawValue
                        }
                        if ($serverIp) { break }
                    }
                }

                if (-not $serverIp) {
                    if ($dnsCache.ContainsKey($cacheKey)) {
                        $serverIp = $dnsCache[$cacheKey]
                    } else {
                        $serverIp = Resolve-IPv4Address -Name $serverName -MessageQueue $MessageQueue -CancellationToken $CancellationToken
                        $dnsCache[$cacheKey] = $serverIp
                    }
                }
                if ($serverIp) { $serverIpValue = [string]$serverIp }
            }

            $row = [ordered]@{ SrNumber = $counter }
            if ([bool]$Request.WantDeviceName) { $row['ServerName'] = $serverName }
            $row['DeviceCollection'] = $deviceCollection
            if ([bool]$Request.WantVDisk) { $row['vDisk'] = $vDiskValue }
            if ([bool]$Request.WantPingReachability) {
                $row['PingReachability'] = if ($pingResults.ContainsKey($serverName)) { [string]$pingResults[$serverName] } else { 'Ping Error' }
            }
            if ([bool]$Request.WantReboot) { $row['RebootDay'] = $rebootValue }
            if ([bool]$Request.WantCSR) { $row['CSRServer'] = $csrValue }
            if ([bool]$Request.WantXDC) { $row['XDCServer'] = $xdcValue }
            if ([bool]$Request.WantServerIP) { $row['ServerIP'] = $serverIpValue }

            [void]$deviceInfo.Add([PSCustomObject]$row)
            [void]$sourceItems.Add([PSCustomObject]@{
                Name = $serverName
                CollectionName = $deviceCollection
                CollectionId = if ($device.PSObject.Properties.Match('CollectionId').Count -gt 0) { [string]$device.CollectionId } else { '' }
                SiteId = if ($device.PSObject.Properties.Match('SiteId').Count -gt 0) { [string]$device.SiteId } else { '' }
                SiteName = if ($device.PSObject.Properties.Match('SiteName').Count -gt 0) { [string]$device.SiteName } else { '' }
            })
        }

        # DHCP validation runs inside the same background retrieval worker so the
        # normal PVS table and DHCP Details tab are published together. A DHCP
        # failure is isolated from the normal PVS result; user cancellation still
        # cancels the combined operation.
        $dhcpResult = [PSCustomObject]@{
            Requested = $false
            Mode = 'None'
            RetrievalScope = ''
            Outcome = 'NotRequested'
            CompletionStatus = 'NotRun'
            IsComplete = $false
            ValidationStatus = 'Not run'
            MastersDiscovered = 0
            MastersComplete = 0
            MasterResults = @()
            TargetResults = @()
            TargetCount = 0
            ResolvedTargetCount = 0
            UnresolvedTargetCount = 0
            ReservationCount = 0
            Findings = @()
            FindingCount = 0
            FindingsOmitted = 0
            RemoteCallCount = 0
            DurationMs = 0
            ErrorMessage = ''
        }
        if ($null -ne $Request.Dhcp -and [bool]$Request.Dhcp.Enabled) {
            try {
                Initialize-CitrixWorkerCommand -CommandName 'Get-PvsServer'
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message "Starting $($Request.Dhcp.Mode) DHCP validation across all PVS Masters..." -IsIndeterminate $true
                $dhcpResult = Invoke-PvsDhcpValidation -Mode ([string]$Request.Dhcp.Mode) -Devices @($devices) -Config ([hashtable]$Request.Dhcp) -MessageQueue $MessageQueue -CancellationToken $CancellationToken -ApprovalGate $ApprovalGate -ApprovalState $ApprovalState
                $dhcpResult | Add-Member -NotePropertyName RetrievalScope -NotePropertyValue ([string]$Request.Dhcp.RetrievalScope) -Force
            } catch [System.OperationCanceledException] {
                throw
            } catch {
                $dhcpError = $_.Exception.Message
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "DHCP validation failed without discarding the normal PVS result: $dhcpError" -Level ERROR
                $dhcpResult = [PSCustomObject]@{
                    Requested = $true
                    Mode = [string]$Request.Dhcp.Mode
                    RetrievalScope = [string]$Request.Dhcp.RetrievalScope
                    Outcome = 'Unavailable'
                    CompletionStatus = 'Unavailable'
                    IsComplete = $false
                    ValidationStatus = 'Not comparable'
                    MastersDiscovered = 0
                    MastersComplete = 0
                    MasterResults = @()
                    TargetResults = @()
                    TargetCount = $totalDevices
                    ResolvedTargetCount = 0
                    UnresolvedTargetCount = $totalDevices
                    ReservationCount = 0
                    Findings = @("DHCP validation failed: $dhcpError")
                    FindingCount = 1
                    FindingsOmitted = 0
                    RemoteCallCount = 0
                    DurationMs = 0
                    ErrorMessage = $dhcpError
                }
            }
        }

        return [PSCustomObject]@{
            Kind = 'Result'
            Role = 'PVS'
            Success = $true
            Cancelled = $false
            ErrorMessage = ''
            AllInventoryCount = @($allDevices).Count
            Data = @($deviceInfo.ToArray())
            SourceItems = @($sourceItems.ToArray())
            MissingNames = @($missingServers)
            AmbiguousNames = @($ambiguousServers)
            DuplicateNames = @($duplicateServers)
            WarningCount = $warningCount
            DhcpResult = $dhcpResult
            SelectedCollectionChoice = $selectedCollectionChoice
        }
    } catch [System.OperationCanceledException] {
        return [PSCustomObject]@{
            Kind = 'Result'
            Role = 'PVS'
            Success = $false
            Cancelled = $true
            ErrorMessage = $_.Exception.Message
        }
    } catch {
        return New-RetrievalWorkerFailureResult -Role 'PVS' -ErrorRecord $_ -Phase 'PVS retrieval worker' -SuggestedAction 'Review the PVS retrieval error in the execution log, correct the Citrix SDK, scope, or connectivity issue, then retry.'
    }
}

function Invoke-DdcRetrievalWorker {
    param(
        [hashtable]$Request,
        [object]$MessageQueue,
        [System.Threading.CancellationToken]$CancellationToken,
        [System.Threading.ManualResetEventSlim]$ApprovalGate,
        [hashtable]$ApprovalState
    )

    try {
        $script:CurrentRetrievalWorkerRole = 'DDC'
        $scopeInventoryOnly = [bool]$Request.ScopeInventoryOnly
        $wholeFarmMode = [bool]$Request.WholeFarmMode
        $machineCatalogMode = [bool]$Request.MachineCatalogMode
        $deliveryGroupMode = [bool]$Request.DeliveryGroupMode
        $singleServerMode = [bool]$Request.SingleServerMode
        $serverListMode = [bool]$Request.ServerListMode
        $calculatedSelectedServerMode = $singleServerMode -or $serverListMode
        $namedScopeMode = $machineCatalogMode -or $deliveryGroupMode
        $selectedDdcScopeChoiceKey = [string]$Request.SelectedDdcScopeChoiceKey
        $selectedDdcKeySupplied = -not [string]::IsNullOrWhiteSpace($selectedDdcScopeChoiceKey)
        $cacheMaterialSupplied = -not [string]::IsNullOrWhiteSpace([string]$Request.InventorySnapshotId) -or
            ($Request.ContainsKey('CachedInventory') -and @($Request.CachedInventory).Count -gt 0)

        if (-not $scopeInventoryOnly) {
            $activeScopeCount = [int]$wholeFarmMode + [int]$machineCatalogMode + [int]$deliveryGroupMode + [int]$singleServerMode + [int]$serverListMode
            if ($activeScopeCount -ne 1) {
                throw 'Invalid XDC retrieval scope state. Select exactly one scope.'
            }
            if ([bool]$Request.SelectedServerMode -ne $calculatedSelectedServerMode) {
                throw 'The XDC selected-server scope state is inconsistent with the Single Server and Server List flags.'
            }
            if ($namedScopeMode -xor $selectedDdcKeySupplied) {
                throw 'The XDC named-scope state is inconsistent with the selected Machine Catalog or Delivery Group key.'
            }
            $requestedNames = @($Request.RequestedServers)
            if ($calculatedSelectedServerMode) {
                if ($requestedNames.Count -eq 0 -or $requestedNames.Count -gt 5000) {
                    throw 'The XDC selected-server request must contain between 1 and 5,000 names.'
                }
                if ($singleServerMode -and $requestedNames.Count -ne 1) {
                    throw 'The XDC Single Server request must contain exactly one name.'
                }
                foreach ($requestedName in $requestedNames) {
                    if ([string]::IsNullOrWhiteSpace([string]$requestedName) -or ([string]$requestedName).Length -gt 512) {
                        throw 'The XDC selected-server request contains an empty or oversized name.'
                    }
                }
            } elseif ($requestedNames.Count -gt 0) {
                throw 'Server names were supplied for an XDC scope that does not accept server-name input.'
            }
            if ($cacheMaterialSupplied) {
                throw 'Cached Broker inventory data is not accepted for retrieval. Retrieve uses a fresh site inventory and only the selected opaque MC/DG key.'
            }
        }

        $PSModuleAutoLoadingPreference = 'None'
        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
        Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message 'Loading the Citrix Broker SDK...' -IsIndeterminate $true
        Initialize-CitrixWorkerCommand -CommandName 'Get-BrokerMachine'

        Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message 'Retrieving the shared Broker site inventory...' -IsIndeterminate $true
        Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message $(if ($scopeInventoryOnly) { 'Retrieving Broker machines for the inline MC/DG lists' } else { 'Retrieving fresh Broker machines for requested output' })
        $allMachines = @(Get-BrokerMachineInventory -MessageQueue $MessageQueue -CancellationToken $CancellationToken)
        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken

        # Machine Catalog and Delivery Group choices come from the same Broker
        # snapshot. Loading both lists at radio-button selection time lets the UI
        # switch between them without another preview query. Retrieve uses the
        # opaque key against a fresh site inventory so volatile Broker state is current.
        if ($scopeInventoryOnly) {
            # Keep only the identity properties needed to display and validate
            # MC/DG choices. Records without a machine name are unusable.
            $missingMachineNameCount = 0
            $snapshotMachines = @(
                foreach ($machine in $allMachines) {
                    $machineName = [string]$machine.MachineName
                    if ([string]::IsNullOrWhiteSpace($machineName)) {
                        $missingMachineNameCount++
                        continue
                    }
                    [PSCustomObject]@{
                        MachineName = $machineName
                        CatalogName = [string]$machine.CatalogName
                        CatalogUUID = if ($machine.PSObject.Properties.Match('CatalogUUID').Count -gt 0) { [string]$machine.CatalogUUID } else { '' }
                        CatalogUid = if ($machine.PSObject.Properties.Match('CatalogUid').Count -gt 0) { $machine.CatalogUid } else { 0 }
                        DesktopGroupName = [string]$machine.DesktopGroupName
                        DesktopGroupUUID = if ($machine.PSObject.Properties.Match('DesktopGroupUUID').Count -gt 0) { [string]$machine.DesktopGroupUUID } else { '' }
                        DesktopGroupUid = if ($machine.PSObject.Properties.Match('DesktopGroupUid').Count -gt 0) { $machine.DesktopGroupUid } else { 0 }
                    }
                }
            )
            if ($missingMachineNameCount -gt 0) {
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Level WARN -Message "$missingMachineNameCount Broker inventory record(s) without MachineName were excluded from the inline MC/DG lists."
            }
            $machineCatalogAnalysis = Get-DdcNamedScopeAnalysis -Machines $snapshotMachines -ScopeType MachineCatalog
            $deliveryGroupAnalysis = Get-DdcNamedScopeAnalysis -Machines $snapshotMachines -ScopeType DeliveryGroup
            return [PSCustomObject]@{
                Kind = 'Result'
                Role = 'DDC'
                Success = $true
                Cancelled = $false
                ErrorMessage = ''
                InventoryOnly = $true
                InventorySnapshotId = [guid]::NewGuid().ToString('N')
                Inventory = @($snapshotMachines)
                InventoryCount = $snapshotMachines.Count
                MissingMachineNameCount = $missingMachineNameCount
                MachineCatalogChoices = @($machineCatalogAnalysis.Choices)
                MachineCatalogUnselectableMachineCount = [int]$machineCatalogAnalysis.UnselectableMachineCount
                MachineCatalogUnsafeMetadataMachineCount = [int]$machineCatalogAnalysis.UnsafeMetadataMachineCount
                DeliveryGroupChoices = @($deliveryGroupAnalysis.Choices)
                DeliveryGroupUnselectableMachineCount = [int]$deliveryGroupAnalysis.UnselectableMachineCount
                DeliveryGroupUnassignedMachineCount = [int]$deliveryGroupAnalysis.UnassignedMachineCount
                DeliveryGroupUnsafeMetadataMachineCount = [int]$deliveryGroupAnalysis.UnsafeMetadataMachineCount
            }
        }

        $machines = $allMachines
        $selectedDdcScopeChoice = $null
        $unselectableScopeMachineCount = 0
        $unassignedScopeMachineCount = 0
        $unsafeScopeMetadataMachineCount = 0
        $scopeChoiceCount = 0
        $missingServers = @()
        $ambiguousServers = @()
        $duplicateServers = @()
        if ($namedScopeMode) {
            if ($allMachines.Count -eq 0) {
                throw 'No Broker machines were returned by the shared site inventory, so no Machine Catalog or Delivery Group can be selected.'
            }

            $scopeType = if ($machineCatalogMode) { 'MachineCatalog' } else { 'DeliveryGroup' }
            $scopeLabel = if ($machineCatalogMode) { 'Machine Catalog' } else { 'Delivery Group' }
            $namedScopeMachines = @($allMachines | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.MachineName) })
            $missingCurrentMachineNameCount = $allMachines.Count - $namedScopeMachines.Count
            if ($missingCurrentMachineNameCount -gt 0) {
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Level WARN -Message "$missingCurrentMachineNameCount Broker inventory record(s) without MachineName were excluded from the selected $scopeLabel retrieval."
            }
            $scopeAnalysis = Get-DdcNamedScopeAnalysis -Machines $namedScopeMachines -ScopeType $scopeType
            $scopeChoices = @($scopeAnalysis.Choices)
            $scopeChoiceCount = $scopeChoices.Count
            $unselectableScopeMachineCount = [int]$scopeAnalysis.UnselectableMachineCount
            $unassignedScopeMachineCount = [int]$scopeAnalysis.UnassignedMachineCount
            $unsafeScopeMetadataMachineCount = [int]$scopeAnalysis.UnsafeMetadataMachineCount
            if ($unassignedScopeMachineCount -gt 0) {
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "$unassignedScopeMachineCount Broker machine record(s) were not assigned to a Delivery Group and therefore do not appear in the inline Delivery Group list."
            }
            if ($unsafeScopeMetadataMachineCount -gt 0) {
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Level WARN -Message "$unsafeScopeMetadataMachineCount Broker machine record(s) did not expose a group name plus a valid UUID or positive UID and were excluded from the inline $scopeLabel list."
            }
            if ($scopeChoices.Count -eq 0) {
                if ($deliveryGroupMode -and $namedScopeMachines.Count -gt 0 -and $unassignedScopeMachineCount -eq $namedScopeMachines.Count) {
                    throw 'No assigned Delivery Groups were represented in the Broker machine inventory.'
                }
                throw "No selectable $scopeLabel choices were returned by the Broker machine inventory."
            }

            $selectedChoiceMatches = @($scopeChoices | Where-Object { $_.ChoiceKey -ieq $selectedDdcScopeChoiceKey } | Select-Object -First 2)
            if ($selectedChoiceMatches.Count -ne 1) {
                throw "The selected $scopeLabel was not present in the retrieved Broker site inventory."
            }
            $selectedDdcScopeChoice = $selectedChoiceMatches[0]
            Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken -Message "$scopeLabel retrieval cancelled before local filtering completed."
            $machines = @(Select-DdcMachinesByNamedScope -Machines $namedScopeMachines -ScopeType $scopeType -ChoiceKey $selectedDdcScopeChoiceKey)
            Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken -Message "$scopeLabel retrieval cancelled before selected rows were processed."
            if ($machines.Count -ne [int]$selectedDdcScopeChoice.MachineCount) {
                throw "The selected $scopeLabel membership changed while the Broker snapshot was being filtered."
            }
            Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "Selected XDC $scopeLabel [$($selectedDdcScopeChoice.ScopeName)] | Identity=$($selectedDdcScopeChoice.IdentifierType):$($selectedDdcScopeChoice.IdentifierValue) | Machines=$($machines.Count) | SiteInventory=$($allMachines.Count)"
        } elseif ([bool]$Request.SelectedServerMode) {
            $scopeResult = Select-ItemsByServerList -Items $allMachines -ServerNames @($Request.RequestedServers) -NameProperties @('MachineName','DNSName')
            $machines = @($scopeResult.Items)
            $missingServers = @($scopeResult.MissingNames)
            $ambiguousServers = @($scopeResult.AmbiguousNames)
            $duplicateServers = @($scopeResult.DuplicateNames)
        }

        $totalMachines = @($machines).Count
        $machineInfo = [System.Collections.Generic.List[object]]::new($totalMachines)
        $sourceItems = [System.Collections.Generic.List[object]]::new($totalMachines)
        $dnsCache = @{}
        $counter = 0

        foreach ($machine in $machines) {
            Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
            $counter++
            Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message "Processing Broker machine $counter of $totalMachines..." -Current $counter -Total $totalMachines -IsIndeterminate $false

            $row = [ordered]@{ SrNumber = $counter }
            if ([bool]$Request.WantName) { $row['MachineName'] = ([string]$machine.MachineName -replace '^.*\\', '') }
            if ([bool]$Request.WantCatalog) { $row['CatalogName'] = $machine.CatalogName }
            if ([bool]$Request.WantDeliveryGroup) { $row['DeliveryGroup'] = $machine.DesktopGroupName }
            if ([bool]$Request.WantRegistrationState) { $row['RegistrationState'] = $machine.RegistrationState }
            if ([bool]$Request.WantMaintenanceMode) { $row['MaintenanceMode'] = $machine.InMaintenanceMode }

            $windowsConnectionSetting = $null
            if ($machine.PSObject.Properties.Match('WindowsConnectionSetting').Count -gt 0) {
                $windowsConnectionSetting = $machine.WindowsConnectionSetting
            }
            if ([bool]$Request.WantLogonMode) {
                if ($null -ne $windowsConnectionSetting) {
                    $mode = [string]$windowsConnectionSetting
                    $row['LogonMode'] = if ($mode -eq 'LogonEnabled') { 'Enabled' } else { 'Disabled' }
                } else {
                    # Maintenance mode and Broker logon mode are independent
                    # states. Do not infer one from the other when the explicit
                    # Broker property is unavailable on an older SDK.
                    $row['LogonMode'] = 'Unavailable'
                }
            }

            if ([bool]$Request.WantServerIP) {
                $serverIp = $null
                $lookupName = if ($machine.DNSName) { [string]$machine.DNSName } else { [string]($machine.MachineName -replace '^.*\\', '') }
                $cacheKey = $lookupName.ToLowerInvariant()
                foreach ($propertyName in @('IPAddress','IPv4Address','IPV4Address')) {
                    if ($machine.PSObject.Properties.Match($propertyName).Count -gt 0 -and $machine.$propertyName) {
                        $rawValue = $machine.$propertyName
                        if ($rawValue -is [System.Array]) {
                            $serverIp = @($rawValue | Where-Object { $_ })[0]
                        } else {
                            $serverIp = $rawValue
                        }
                        if ($serverIp) { break }
                    }
                }
                if (-not $serverIp) {
                    if ($dnsCache.ContainsKey($cacheKey)) {
                        $serverIp = $dnsCache[$cacheKey]
                    } else {
                        $serverIp = Resolve-IPv4Address -Name $lookupName -MessageQueue $MessageQueue -CancellationToken $CancellationToken
                        $dnsCache[$cacheKey] = $serverIp
                    }
                }
                $row['ServerIP'] = if ($serverIp) { [string]$serverIp } else { 'N/A' }
            }

            [void]$machineInfo.Add([PSCustomObject]$row)
            [void]$sourceItems.Add([PSCustomObject]@{
                MachineName = [string]$machine.MachineName
                DNSName = [string]$machine.DNSName
                CatalogName = [string]$machine.CatalogName
                DesktopGroupName = [string]$machine.DesktopGroupName
                RegistrationState = [string]$machine.RegistrationState
                InMaintenanceMode = [bool]$machine.InMaintenanceMode
                WindowsConnectionSetting = $windowsConnectionSetting
            })
        }

        return [PSCustomObject]@{
            Kind = 'Result'
            Role = 'DDC'
            Success = $true
            Cancelled = $false
            ErrorMessage = ''
            AllInventoryCount = @($allMachines).Count
            Data = @($machineInfo.ToArray())
            SourceItems = @($sourceItems.ToArray())
            MissingNames = @($missingServers)
            AmbiguousNames = @($ambiguousServers)
            DuplicateNames = @($duplicateServers)
            SelectedScopeChoice = $selectedDdcScopeChoice
            ScopeChoiceCount = $scopeChoiceCount
            UnselectableScopeMachineCount = $unselectableScopeMachineCount
            UnassignedScopeMachineCount = $unassignedScopeMachineCount
            UnsafeScopeMetadataMachineCount = $unsafeScopeMetadataMachineCount
            WarningCount = 0
        }
    } catch [System.OperationCanceledException] {
        return [PSCustomObject]@{
            Kind = 'Result'
            Role = 'DDC'
            Success = $false
            Cancelled = $true
            ErrorMessage = $_.Exception.Message
        }
    } catch {
        return New-RetrievalWorkerFailureResult -Role 'DDC' -ErrorRecord $_ -Phase 'XDC retrieval worker' -SuggestedAction 'Review the XDC retrieval error in the execution log, correct the Citrix Broker SDK, scope, or connectivity issue, then retry.'
    }
}

function Get-RetrievalWorkerFunctionText {
    param(
        [string[]]$FunctionNames
    )

    $builder = [System.Text.StringBuilder]::new()
    foreach ($functionName in $FunctionNames) {
        $command = Get-Command -Name $functionName -CommandType Function -ErrorAction Stop
        [void]$builder.AppendLine("function $functionName {")
        [void]$builder.AppendLine($command.Definition)
        [void]$builder.AppendLine('}')
    }
    return $builder.ToString()
}

# Defines every function that is serialized into a background runspace. Keeping
# this manifest in one place prevents foreground/worker drift when a helper is
# added, renamed, or moved. It is intentionally data-only: callers retain their
# current runspace construction and function serialization behavior.
function Get-CitrixWorkerDependencyManifest {
    return [ordered]@{
        Common = @(
            'Get-ErrorDiagnosticText',
            'Test-TrustedInstallPath',
            'Test-TrustedCwxPvsCmdletAssembly',
            'Test-TrustedCitrixCommandInfo',
            'Initialize-CitrixWorkerCommand',
            'Send-RetrievalWorkerMessage',
            'Test-RetrievalWorkerCancellation',
            'New-RetrievalWorkerFailureResult',
            'Resolve-IPv4Address',
            'Get-ObjectIPv4AddressValue',
            'Get-ServerNameAliases',
            'Select-ItemsByServerList'
        )
        PVS = @(
            'Get-PvsDeviceCollectionChoiceKey',
            'Get-PvsDeviceCollectionChoices',
            'Select-PvsDevicesByCollection',
            'Get-BulkPingResults',
            'ConvertTo-DhcpOptionValueText',
            'ConvertTo-DhcpClientIdKey',
            'Get-DhcpValidationOptionMap',
            'Initialize-TrustedDhcpServerModule',
            'Convert-DhcpIPv4ToUInt32',
            'Test-IPv4InDhcpScope',
            'New-DhcpOptionCells',
            'Get-DhcpOptionSet',
            'Get-DhcpEffectiveOptionSet',
            'Resolve-PvsMastersForDhcp',
            'Resolve-DhcpTargetSet',
            'Get-PvsDhcpMasterSnapshot',
            'Get-CitrixWorkerDependencyManifest',
            'Get-DhcpMasterWorkerScriptText',
            'New-DhcpTimedOutMasterResult',
            'Invoke-DhcpMasterCollection',
            'Get-DhcpOptionCellSignature',
            'Get-DhcpOptionCellDisplayText',
            'Invoke-PvsDhcpValidation',
            'Invoke-PvsRetrievalWorker'
        )
        DDC = @(
            'Get-DdcNamedScopeRecord',
            'Get-DdcNamedScopeAnalysis',
            'Select-DdcMachinesByNamedScope',
            'Get-BrokerMachineInventory',
            'Invoke-DdcRetrievalWorker'
        )
        DhcpMaster = @(
            'Test-TrustedInstallPath',
            'Initialize-TrustedDhcpServerModule',
            'Test-RetrievalWorkerCancellation',
            'ConvertTo-DhcpOptionValueText',
            'ConvertTo-DhcpClientIdKey',
            'Convert-DhcpIPv4ToUInt32',
            'Test-IPv4InDhcpScope',
            'New-DhcpOptionCells',
            'Get-DhcpOptionSet',
            'Get-DhcpEffectiveOptionSet',
            'Get-PvsDhcpMasterSnapshot'
        )
        Probe = @(
            'Test-TrustedInstallPath',
            'Test-TrustedCwxPvsCmdletAssembly',
            'Test-TrustedCitrixCommandInfo',
            'Initialize-CitrixWorkerCommand'
        )
    }
}

function Start-CitrixRetrievalWorker {
    param(
        [ValidateSet('PVS','DDC')]
        [string]$Role,
        [hashtable]$Request
    )

    $messageQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $cancellationSource = [System.Threading.CancellationTokenSource]::new()
    $approvalGate = [System.Threading.ManualResetEventSlim]::new($false)
    $approvalState = [hashtable]::Synchronized(@{
        Approved = $false
    })
    $runspace = $null
    $pipeline = $null

    try {
        $workerDependencies = Get-CitrixWorkerDependencyManifest
        $commonFunctions = @($workerDependencies.Common)
        if ($Role -eq 'PVS') {
            $functionNames = $commonFunctions + @($workerDependencies.PVS)
            $workerFunction = 'Invoke-PvsRetrievalWorker'
        } else {
            $functionNames = $commonFunctions + @($workerDependencies.DDC)
            $workerFunction = 'Invoke-DdcRetrievalWorker'
        }

        $functionText = Get-RetrievalWorkerFunctionText -FunctionNames $functionNames
        $workerScript = @(
            'param($Request, $MessageQueue, $CancellationToken, $ApprovalGate, $ApprovalState)'
            $functionText
            "$workerFunction -Request `$Request -MessageQueue `$MessageQueue -CancellationToken `$CancellationToken -ApprovalGate `$ApprovalGate -ApprovalState `$ApprovalState"
        ) -join "`r`n"

        $runspace = [RunspaceFactory]::CreateRunspace()
        $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
        $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $runspace.Open()

        $pipeline = [PowerShell]::Create()
        $pipeline.Runspace = $runspace
        [void]$pipeline.AddScript($workerScript).AddArgument($Request).AddArgument($messageQueue).AddArgument($cancellationSource.Token).AddArgument($approvalGate).AddArgument($approvalState)
        $handle = $pipeline.BeginInvoke()

        return [PSCustomObject]@{
            RunId = [guid]::NewGuid().ToString('N')
            Role = $Role
            Request = $Request
            MessageQueue = $messageQueue
            CancellationSource = $cancellationSource
            ApprovalGate = $approvalGate
            ApprovalState = $approvalState
            Runspace = $runspace
            Pipeline = $pipeline
            Handle = $handle
            StopHandle = $null
            CancellationRequested = $false
            Completed = $false
        }
    } catch {
        if ($null -ne $pipeline) { try { $pipeline.Dispose() } catch { } }
        if ($null -ne $runspace) {
            try { $runspace.Close() } catch { }
            try { $runspace.Dispose() } catch { }
        }
        try { $approvalGate.Dispose() } catch { }
        try { $cancellationSource.Dispose() } catch { }
        throw
    }
}

function Receive-CitrixRetrievalMessages {
    param(
        $Context,
        [ValidateRange(1, 1000)]
        [int]$MaxMessages = 200
    )

    $messages = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $Context -or $null -eq $Context.MessageQueue) { return @() }

    $message = $null
    while ($messages.Count -lt $MaxMessages -and $Context.MessageQueue.TryDequeue([ref]$message)) {
        [void]$messages.Add($message)
        $message = $null
    }
    return @($messages.ToArray())
}

function Request-CitrixRetrievalCancellation {
    param(
        $Context
    )

    if ($null -eq $Context -or [bool]$Context.Completed -or [bool]$Context.CancellationRequested) { return }
    # The worker may finish during the interval between dispatcher polls. In
    # that case the successful result wins and the next poll publishes it.
    if ($null -ne $Context.Handle -and $Context.Handle.IsCompleted) { return }
    $Context.CancellationRequested = $true
    $Context.CancellationSource.Cancel()
    $Context.ApprovalGate.Set()

    if ($null -ne $Context.Handle -and -not $Context.Handle.IsCompleted) {
        try {
            $Context.StopHandle = $Context.Pipeline.BeginStop($null, $null)
        } catch {
            # The cooperative token remains set if the pipeline rejects BeginStop.
        }
    }
}

function Complete-CitrixRetrievalWorker {
    param(
        $Context
    )

    if ($null -eq $Context) {
        return [PSCustomObject]@{ Outcome = 'Failed'; Result = $null; ErrorMessage = 'Retrieval context is unavailable.'; ErrorDetail = 'Phase=Worker completion | Context=Unavailable' }
    }

    $result = $null
    $errorMessage = ''
    $errorDetail = ''
    $outcome = 'Failed'
    try {
        $output = @($Context.Pipeline.EndInvoke($Context.Handle))
        $result = $output | Where-Object { $_.PSObject.Properties.Match('Kind').Count -gt 0 -and $_.Kind -eq 'Result' } | Select-Object -Last 1
        if ($null -ne $result -and [bool]$result.Success) {
            $outcome = 'Completed'
        } elseif ([bool]$Context.CancellationRequested -or ($null -ne $result -and [bool]$result.Cancelled)) {
            $outcome = 'Cancelled'
        } elseif ($null -ne $result) {
            $errorMessage = [string]$result.ErrorMessage
            if ($result.PSObject.Properties.Match('ErrorDetail').Count -gt 0) {
                $errorDetail = [string]$result.ErrorDetail
            }
        } else {
            $streamErrors = @($Context.Pipeline.Streams.Error)
            if ($streamErrors.Count -gt 0) {
                $errorMessage = [string]$streamErrors[-1].Exception.Message
                $errorDetail = Get-ErrorDiagnosticText -ErrorRecord $streamErrors[-1] -Phase "$($Context.Role) worker error stream"
            }
            if ([string]::IsNullOrWhiteSpace($errorMessage)) { $errorMessage = 'The background worker returned no result.' }
            if ([string]::IsNullOrWhiteSpace($errorDetail)) { $errorDetail = "Phase=$($Context.Role) worker completion | Result=Missing" }
        }
    } catch {
        if ([bool]$Context.CancellationRequested -or $_.Exception -is [System.Management.Automation.PipelineStoppedException]) {
            $outcome = 'Cancelled'
        } else {
            $errorMessage = $_.Exception.Message
            $errorDetail = Get-ErrorDiagnosticText -ErrorRecord $_ -Phase "$($Context.Role) worker EndInvoke"
        }
    } finally {
        $Context.Completed = $true
        if ($null -ne $Context.StopHandle -and $Context.StopHandle.IsCompleted) {
            try { $Context.Pipeline.EndStop($Context.StopHandle) } catch { }
        }
        try { $Context.Pipeline.Dispose() } catch {
            Write-Log "Worker pipeline cleanup warning | Role=$($Context.Role) | RunId=$($Context.RunId) | $($_.Exception.Message)" -Level WARN
        }
        try { $Context.Runspace.Close() } catch { }
        try { $Context.Runspace.Dispose() } catch {
            Write-Log "Worker runspace cleanup warning | Role=$($Context.Role) | RunId=$($Context.RunId) | $($_.Exception.Message)" -Level WARN
        }
        try { $Context.ApprovalGate.Dispose() } catch {
            Write-Log "Worker approval-gate cleanup warning | Role=$($Context.Role) | RunId=$($Context.RunId) | $($_.Exception.Message)" -Level WARN
        }
        try { $Context.CancellationSource.Dispose() } catch {
            Write-Log "Worker cancellation-source cleanup warning | Role=$($Context.Role) | RunId=$($Context.RunId) | $($_.Exception.Message)" -Level WARN
        }
    }

    return [PSCustomObject]@{
        Outcome = $outcome
        Result = $result
        ErrorMessage = $errorMessage
        ErrorDetail = $errorDetail
    }
}

# -------------------------------------------------------------
# Region: Production-safe PVS DHCP Validation
#
# DHCP mode is derived from the PVS retrieval scope:
#   Baseline  - Whole Farm: Master/service, exact scopes, and server/scope options.
#   Targeted  - Device Collection, Server List, or Single Server: only the matched targets on all Masters.
#   DeepAudit - Explicit Whole Farm audit: Baseline plus every reservation override.
#
# The outer PVS worker owns orchestration and cancellation. Independent Master
# snapshots run through a nested pool capped at two by default. Calls remain
# serial inside each Master. Every snapshot is flattened before it leaves its
# runspace, and option cells distinguish Value, Unset, and ReadFailed so failed
# reads can never be reported as clean matches.
# -------------------------------------------------------------

function ConvertTo-DhcpOptionValueText {
    param([object]$Value)

    if ($null -eq $Value) { return '' }
    # Preserve DHCP array order. Router and DNS-server order can be operationally
    # significant, so the comparison must not sort multi-valued options.
    $values = @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($values.Count -eq 0) { return '' }
    return ($values -join ', ')
}

function ConvertTo-DhcpClientIdKey {
    param([object]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [byte[]]) {
        $hex = (($Value | ForEach-Object { $_.ToString('X2') }) -join '')
    } else {
        $hex = ([string]$Value -replace '[^0-9A-Fa-f]', '')
    }
    # Ethernet DHCP client identifiers commonly include a leading hardware-type
    # byte (01). Normalize that form to the same 12-hex-digit MAC key used by PVS.
    if ($hex.Length -eq 14 -and $hex.StartsWith('01', [System.StringComparison]::OrdinalIgnoreCase)) {
        $hex = $hex.Substring(2)
    }
    return $hex.ToUpperInvariant()
}

function Get-DhcpValidationOptionMap {
    return @{
        3  = 'Router'
        6  = 'DNS Servers'
        11 = 'Resource Location Servers'
        15 = 'DNS Domain Name'
        66 = 'Boot Server Host Name'
        67 = 'Bootfile Name'
    }
}

# Loads only the inbox/RSAT DHCPServer module installed beneath the 64-bit
# Windows PowerShell home. An identically named user module is never imported.
function Initialize-TrustedDhcpServerModule {
    $PSModuleAutoLoadingPreference = 'None'
    $moduleRoot = Join-Path $PSHOME 'Modules\DhcpServer'
    $manifestPath = Join-Path $moduleRoot 'DhcpServer.psd1'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "The trusted DHCPServer module manifest was not found at [$manifestPath]. Run the tool in 64-bit Windows PowerShell with the DHCP management tools installed."
    }

    $fullModuleRoot = [IO.Path]::GetFullPath($moduleRoot).TrimEnd([char[]]@('\','/')) + [IO.Path]::DirectorySeparatorChar
    foreach ($loadedModule in @(Get-Module -Name DhcpServer)) {
        $loadedPath = if ($loadedModule.Path) { [IO.Path]::GetFullPath([string]$loadedModule.Path) } else { '' }
        if (-not (Test-TrustedInstallPath -Path $loadedPath) -or
            -not $loadedPath.StartsWith($fullModuleRoot, [StringComparison]::OrdinalIgnoreCase)) {
            throw "An untrusted DHCPServer module is already loaded from [$loadedPath]. Start a dedicated powershell.exe -NoProfile process."
        }
    }

    $trustedLoadedModule = @(
        Get-Module -Name DhcpServer |
            Where-Object {
                $_.Path -and ([IO.Path]::GetFullPath([string]$_.Path)).StartsWith($fullModuleRoot, [StringComparison]::OrdinalIgnoreCase)
            }
    ) | Select-Object -First 1
    if ($null -eq $trustedLoadedModule) {
        Import-Module -Name $manifestPath -ErrorAction Stop
    }

    foreach ($commandName in @('Get-DhcpServerv4Scope','Get-DhcpServerv4OptionValue','Get-DhcpServerv4Reservation')) {
        $resolvedCommands = @(Get-Command -Name $commandName -All -ErrorAction SilentlyContinue)
        $trustedCommand = $resolvedCommands | Select-Object -First 1
        $trustedCommandPath = if ($null -ne $trustedCommand -and $null -ne $trustedCommand.Module) { [string]$trustedCommand.Module.Path } else { '' }
        $trustedCommandPathIsApproved = $false
        if (-not [string]::IsNullOrWhiteSpace($trustedCommandPath)) {
            try {
                $trustedCommandPathIsApproved = ([IO.Path]::GetFullPath($trustedCommandPath)).StartsWith($fullModuleRoot, [StringComparison]::OrdinalIgnoreCase)
            } catch { }
        }
        if ($null -eq $trustedCommand -or
            [string]$trustedCommand.CommandType -notin @('Cmdlet','Function') -or
            $trustedCommand.ModuleName -ine 'DhcpServer' -or
            -not $trustedCommandPathIsApproved) {
            if ($null -ne $trustedCommand) {
                throw "DHCP command [$commandName] is shadowed by an untrusted source: Type=$($trustedCommand.CommandType), Source=$($trustedCommand.Source). Start a dedicated powershell.exe -NoProfile process and remove the conflicting command."
            }
            throw "Trusted DHCP command [$commandName] is unavailable from [$manifestPath]."
        }
    }
}

function Get-ObjectIPv4AddressValue {
    param(
        [object]$InputObject,
        [string[]]$PropertyNames = @('IPAddress','IpAddress','IPv4Address','IPV4Address','IP','Address','DeviceIP','DeviceIPAddress')
    )

    if ($null -eq $InputObject) { return $null }
    foreach ($propertyName in $PropertyNames) {
        if ($InputObject.PSObject.Properties.Match($propertyName).Count -eq 0) { continue }
        foreach ($candidate in @($InputObject.$propertyName)) {
            if ($null -eq $candidate) { continue }
            try {
                $ip = [System.Net.IPAddress]::Parse(([string]$candidate).Trim())
                if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
                    return $ip.ToString()
                }
            } catch { }
        }
    }
    return $null
}

function Convert-DhcpIPv4ToUInt32 {
    param([string]$IpAddress)

    try {
        $ip = [System.Net.IPAddress]::Parse($IpAddress)
        if ($ip.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { return $null }
        $bytes = $ip.GetAddressBytes()
        [array]::Reverse($bytes)
        return [BitConverter]::ToUInt32($bytes, 0)
    } catch {
        return $null
    }
}

function Test-IPv4InDhcpScope {
    param(
        [string]$IpAddress,
        [object]$Scope
    )

    try {
        $ip = [System.Net.IPAddress]::Parse($IpAddress)
        $scopeId = [System.Net.IPAddress]::Parse([string]$Scope.ScopeId)
        $mask = [System.Net.IPAddress]::Parse([string]$Scope.SubnetMask)
        if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork -and
            $scopeId.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork -and
            $mask.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
            $ipBytes = $ip.GetAddressBytes()
            $scopeBytes = $scopeId.GetAddressBytes()
            $maskBytes = $mask.GetAddressBytes()
            for ($i = 0; $i -lt 4; $i++) {
                if (($ipBytes[$i] -band $maskBytes[$i]) -ne ($scopeBytes[$i] -band $maskBytes[$i])) {
                    return $false
                }
            }
            return $true
        }
    } catch { }

    $ipNumber = Convert-DhcpIPv4ToUInt32 -IpAddress $IpAddress
    $startNumber = Convert-DhcpIPv4ToUInt32 -IpAddress ([string]$Scope.StartRange)
    $endNumber = Convert-DhcpIPv4ToUInt32 -IpAddress ([string]$Scope.EndRange)
    return ($null -ne $ipNumber -and $null -ne $startNumber -and $null -ne $endNumber -and
        $ipNumber -ge $startNumber -and $ipNumber -le $endNumber)
}

function New-DhcpOptionCells {
    param(
        [int[]]$OptionIds,
        [ValidateSet('Server','Scope','Reservation')]
        [string]$Source,
        [ValidateSet('Unset','ReadFailed')]
        [string]$InitialState = 'Unset'
    )

    $cells = @{}
    foreach ($optionId in $OptionIds) {
        $cells[[string]$optionId] = [PSCustomObject]@{
            State = $InitialState
            Value = ''
            Source = $Source
        }
    }
    return $cells
}

function Get-DhcpOptionSet {
    param(
        [string]$ComputerName,
        [int[]]$OptionIds,
        [ValidateSet('Server','Scope','Reservation')]
        [string]$Level,
        [string]$ScopeId,
        [string]$ReservedIP,
        [System.Threading.CancellationToken]$CancellationToken = [System.Threading.CancellationToken]::None
    )

    Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken -Message "DHCP validation cancelled while reading $Level options on $ComputerName."
    $cells = New-DhcpOptionCells -OptionIds $OptionIds -Source $Level
    $parameters = @{
        ComputerName = $ComputerName
        ErrorAction = 'Stop'
    }
    if ($Level -eq 'Scope') { $parameters.ScopeId = [System.Net.IPAddress]::Parse($ScopeId) }
    if ($Level -eq 'Reservation') { $parameters.ReservedIP = [System.Net.IPAddress]::Parse($ReservedIP) }

    $command = Get-Command Get-DhcpServerv4OptionValue -ErrorAction Stop
    if ($command.Parameters.ContainsKey('Brief')) { $parameters.Brief = $true }
    if ($command.Parameters.ContainsKey('ThrottleLimit')) { $parameters.ThrottleLimit = 1 }

    try {
        # Omitting OptionId retrieves all configured standard values in one call.
        # This is compatible with older in-box DHCP modules and avoids six calls.
        $rows = @(Get-DhcpServerv4OptionValue @parameters)
        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken -Message "DHCP validation cancelled while reading $Level options on $ComputerName."
        foreach ($row in $rows) {
            if ($row.PSObject.Properties.Match('OptionId').Count -eq 0) { continue }
            $optionId = [int]$row.OptionId
            $key = [string]$optionId
            if (-not $cells.ContainsKey($key)) { continue }
            $valueText = ConvertTo-DhcpOptionValueText -Value $row.Value
            if (-not [string]::IsNullOrWhiteSpace($valueText)) {
                $cells[$key] = [PSCustomObject]@{
                    State = 'Value'
                    Value = $valueText
                    Source = $Level
                }
            }
        }
        return [PSCustomObject]@{
            ReadSucceeded = $true
            Values = $cells
            Error = ''
            RemoteCallCount = 1
        }
    } catch [System.OperationCanceledException] {
        throw
    } catch {
        if ($CancellationToken.IsCancellationRequested) {
            throw [System.OperationCanceledException]::new("DHCP validation cancelled while reading $Level options on $ComputerName.")
        }
        return [PSCustomObject]@{
            ReadSucceeded = $false
            Values = (New-DhcpOptionCells -OptionIds $OptionIds -Source $Level -InitialState ReadFailed)
            Error = $_.Exception.Message
            RemoteCallCount = 1
        }
    }
}

function Get-DhcpEffectiveOptionSet {
    param(
        [int[]]$OptionIds,
        [hashtable]$ReservationOptions,
        [hashtable]$ScopeOptions,
        [hashtable]$ServerOptions
    )

    $effective = @{}
    foreach ($optionId in $OptionIds) {
        $key = [string]$optionId
        $selected = $null
        foreach ($tier in @($ReservationOptions, $ScopeOptions, $ServerOptions)) {
            if ($null -eq $tier -or -not $tier.ContainsKey($key)) { continue }
            $cell = $tier[$key]
            if ([string]$cell.State -eq 'Value') {
                $selected = [PSCustomObject]@{ State = 'Value'; Value = [string]$cell.Value; Source = [string]$cell.Source }
                break
            }
            if ([string]$cell.State -eq 'ReadFailed') {
                $selected = [PSCustomObject]@{ State = 'ReadFailed'; Value = ''; Source = [string]$cell.Source }
                break
            }
        }
        if ($null -eq $selected) {
            $selected = [PSCustomObject]@{ State = 'Unset'; Value = ''; Source = '' }
        }
        $effective[$key] = $selected
    }
    return $effective
}

function Resolve-PvsMastersForDhcp {
    param(
        [object]$MessageQueue,
        [datetime]$DeadlineUtc = [datetime]::MaxValue,
        [System.Threading.CancellationToken]$CancellationToken = [System.Threading.CancellationToken]::None
    )

    $rawMasters = [System.Collections.Generic.List[object]]::new()
    $enumerationError = ''
    $timedOut = $false
    try {
        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
        $servers = @(Get-PvsServer -ErrorAction Stop)
        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
        foreach ($server in $servers) {
            Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
            if ([datetime]::UtcNow -ge $DeadlineUtc) { $timedOut = $true; break }
            $name = ''
            foreach ($propertyName in @('DNSName','FQDN','ServerName','Name','HostName')) {
                if ($server.PSObject.Properties.Match($propertyName).Count -gt 0 -and
                    -not [string]::IsNullOrWhiteSpace([string]$server.$propertyName)) {
                    $name = ([string]$server.$propertyName).Trim().TrimEnd('.')
                    if ($name -match '\\') { $name = ($name -replace '^.*\\', '') }
                    break
                }
            }
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            [void]$rawMasters.Add([PSCustomObject]@{
                Name = $name
                ExpectedIPv4 = Get-ObjectIPv4AddressValue -InputObject $server
            })
        }
    } catch [System.OperationCanceledException] {
        throw
    } catch {
        if ($CancellationToken.IsCancellationRequested) {
            throw [System.OperationCanceledException]::new('DHCP validation cancelled during PVS Master discovery.')
        }
        $enumerationError = $_.Exception.Message
        Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "Unable to enumerate all PVS Masters for DHCP validation: $enumerationError" -Level WARN
    }

    Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
    if ([datetime]::UtcNow -ge $DeadlineUtc) { $timedOut = $true }
    $localName = [string]$env:COMPUTERNAME
    $localShort = ($localName -split '\.')[0]
    $hasLocalAlias = @($rawMasters | Where-Object { (([string]$_.Name -split '\.')[0]) -ieq $localShort }).Count -gt 0
    if (-not $hasLocalAlias -and -not [string]::IsNullOrWhiteSpace($localName)) {
        [void]$rawMasters.Insert(0, [PSCustomObject]@{ Name = $localName; ExpectedIPv4 = $null })
    }

    $seenExact = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $masters = [System.Collections.Generic.List[object]]::new()
    $ordinal = 0
    foreach ($master in $rawMasters) {
        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
        if ([datetime]::UtcNow -ge $DeadlineUtc) { $timedOut = $true; break }
        $name = [string]$master.Name
        if ([string]::IsNullOrWhiteSpace($name) -or -not $seenExact.Add($name)) { continue }
        $ordinal++
        $expectedIp = [string]$master.ExpectedIPv4
        if ([string]::IsNullOrWhiteSpace($expectedIp)) {
            $expectedIp = Resolve-IPv4Address -Name $name -MessageQueue $MessageQueue -DeadlineUtc $DeadlineUtc -CancellationToken $CancellationToken
        }
        [void]$masters.Add([PSCustomObject]@{
            Ordinal = $ordinal
            Name = $name
            ShortName = (($name -split '\.')[0])
            ExpectedIPv4 = if ($expectedIp) { [string]$expectedIp } else { '' }
        })
    }

    return [PSCustomObject]@{
        Masters = @($masters.ToArray())
        EnumerationError = $enumerationError
        TimedOut = $timedOut -or [datetime]::UtcNow -ge $DeadlineUtc
    }
}

function Resolve-DhcpTargetSet {
    param(
        [object[]]$Devices,
        [object]$MessageQueue,
        [datetime]$DeadlineUtc = [datetime]::MaxValue,
        [System.Threading.CancellationToken]$CancellationToken = [System.Threading.CancellationToken]::None
    )

    $targets = [System.Collections.Generic.List[object]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $dnsCache = @{}
    $ipOwners = @{}
    $ordinal = 0
    $timedOut = $false
    $totalDeviceCount = @($Devices).Count

    foreach ($device in $Devices) {
        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
        if ([datetime]::UtcNow -ge $DeadlineUtc) {
            $timedOut = $true
            break
        }
        $ordinal++
        $name = [string]$device.Name
        $ip = Get-ObjectIPv4AddressValue -InputObject $device
        $source = if ($ip) { 'PVS inventory' } else { '' }
        if (-not $ip -and -not [string]::IsNullOrWhiteSpace($name)) {
            $cacheKey = $name.ToLowerInvariant()
            if ($dnsCache.ContainsKey($cacheKey)) {
                $ip = $dnsCache[$cacheKey]
            } else {
                $ip = Resolve-IPv4Address -Name $name -MessageQueue $MessageQueue -DeadlineUtc $DeadlineUtc -CancellationToken $CancellationToken
                $dnsCache[$cacheKey] = $ip
            }
            if ($ip) { $source = 'DNS fallback' }
        }

        $resolutionState = if ($ip) { 'Resolved' } else { 'Failed' }
        if (-not $ip) { [void]$warnings.Add("Target $name could not be resolved to IPv4.") }
        if ($ip) {
            if ($ipOwners.ContainsKey($ip)) {
                [void]$warnings.Add("Targets $($ipOwners[$ip]) and $name resolve to the same IPv4 address $ip.")
            } else {
                $ipOwners[$ip] = $name
            }
        }
        $expectedClientId = ''
        foreach ($propertyName in @('DeviceMac','DeviceMAC','MacAddress','MAC','ClientId')) {
            if ($device.PSObject.Properties.Match($propertyName).Count -eq 0 -or $null -eq $device.$propertyName) { continue }
            $expectedClientId = ConvertTo-DhcpClientIdKey -Value $device.$propertyName
            if ($expectedClientId) { break }
        }
        if ([string]::IsNullOrWhiteSpace($expectedClientId)) {
            $targetLabel = if ([string]::IsNullOrWhiteSpace($name)) { "target ordinal $ordinal" } else { "Target $name" }
            [void]$warnings.Add("$targetLabel has no usable PVS MAC/client-ID metadata; reservation client-ID correctness cannot be proven.")
        }
        [void]$targets.Add([PSCustomObject]@{
            Ordinal = $ordinal
            Key = $name.ToLowerInvariant()
            Name = $name
            IPAddress = if ($ip) { [string]$ip } else { '' }
            AddressSource = $source
            ResolutionState = $resolutionState
            ExpectedClientId = $expectedClientId
        })
        if ([datetime]::UtcNow -ge $DeadlineUtc) {
            $timedOut = $true
            break
        }
    }

    if ($timedOut) {
        [void]$warnings.Add("Target IPv4 resolution stopped at the overall DHCP soft deadline after $($targets.Count) of $totalDeviceCount targets.")
    }

    return [PSCustomObject]@{
        Targets = @($targets.ToArray())
        Warnings = @($warnings.ToArray())
        ResolvedCount = @($targets | Where-Object { $_.ResolutionState -eq 'Resolved' }).Count
        UnresolvedCount = @($targets | Where-Object { $_.ResolutionState -ne 'Resolved' }).Count
        TimedOut = $timedOut
        TotalRequestedCount = $totalDeviceCount
    }
}

function Get-PvsDhcpMasterSnapshot {
    param(
        [object]$Master,
        [ValidateSet('Baseline','Targeted','DeepAudit')]
        [string]$Mode,
        [object[]]$Targets,
        [int[]]$OptionIds,
        [System.Threading.CancellationToken]$CancellationToken = [System.Threading.CancellationToken]::None
    )

    $started = [datetime]::UtcNow
    $errors = [System.Collections.Generic.List[string]]::new()
    $remoteCalls = 0
    $serviceState = 'Unavailable'
    $scopesOut = [System.Collections.Generic.List[object]]::new()
    $targetResults = [System.Collections.Generic.List[object]]::new()
    $deepReservations = [System.Collections.Generic.List[object]]::new()
    $completionStatus = 'Complete'
    $serviceQuerySucceeded = $true
    $targetBatchQueryFailed = $false

    try {
        $PSModuleAutoLoadingPreference = 'None'
        Initialize-TrustedDhcpServerModule

        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
        try {
            $remoteCalls++
            $service = Get-Service -ComputerName $Master.Name -Name 'DHCPServer' -ErrorAction Stop
            $serviceState = [string]$service.Status
        } catch {
            if ($CancellationToken.IsCancellationRequested) { throw [System.OperationCanceledException]::new('DHCP validation cancelled during service query.') }
            $message = if ([string]$_.FullyQualifiedErrorId -like 'NoServiceFoundForGivenName*') {
                'DHCP Server service is not installed.'
            } else {
                "DHCP service query failed: $($_.Exception.Message)"
            }
            $serviceState = 'Unknown'
            $serviceQuerySucceeded = $false
            [void]$errors.Add($message)
            # SCM/RPC access can fail even when the DHCP CIM cmdlets are usable.
            # Continue collecting configuration and report service state separately.
        }

        Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
        try {
            $remoteCalls++
            $rawScopes = @(Get-DhcpServerv4Scope -ComputerName $Master.Name -ErrorAction Stop)
        } catch {
            if ($CancellationToken.IsCancellationRequested) { throw [System.OperationCanceledException]::new('DHCP validation cancelled during scope inventory.') }
            [void]$errors.Add("Scope inventory read failed: $($_.Exception.Message)")
            return [PSCustomObject]@{
                Kind = 'DhcpMasterResult'; Ordinal = [int]$Master.Ordinal; MasterName = [string]$Master.Name
                ExpectedIPv4 = [string]$Master.ExpectedIPv4; CompletionStatus = 'Partial'; ServiceState = $serviceState
                ServiceQuerySucceeded = $serviceQuerySucceeded
                ServerOptions = @{}; Scopes = @(); TargetResults = @(); DeepReservations = @(); Errors = @($errors.ToArray())
                RemoteCallCount = $remoteCalls; DurationMs = [int]([datetime]::UtcNow - $started).TotalMilliseconds
            }
        }

        $serverOptionResult = Get-DhcpOptionSet -ComputerName $Master.Name -OptionIds $OptionIds -Level Server -CancellationToken $CancellationToken
        $remoteCalls += [int]$serverOptionResult.RemoteCallCount
        if (-not [bool]$serverOptionResult.ReadSucceeded) {
            $completionStatus = 'Partial'
            [void]$errors.Add("Server option read failed: $($serverOptionResult.Error)")
        }

        $scopeLookup = @{}
        foreach ($rawScope in $rawScopes) {
            $scopeKey = [string]$rawScope.ScopeId
            $scopeLookup[$scopeKey] = [PSCustomObject]@{
                ScopeId = $scopeKey
                StartRange = [string]$rawScope.StartRange
                EndRange = [string]$rawScope.EndRange
                SubnetMask = [string]$rawScope.SubnetMask
                State = [string]$rawScope.State
                Raw = $rawScope
                IsRelevant = ($Mode -ne 'Targeted')
            }
        }

        $targetScopeByKey = @{}
        if ($Mode -eq 'Targeted') {
            foreach ($target in $Targets) {
                if ([string]$target.ResolutionState -ne 'Resolved') { continue }
                foreach ($scopeKey in @($scopeLookup.Keys)) {
                    if (Test-IPv4InDhcpScope -IpAddress ([string]$target.IPAddress) -Scope $scopeLookup[$scopeKey]) {
                        $scopeLookup[$scopeKey].IsRelevant = $true
                        $targetScopeByKey[[string]$target.Key] = $scopeKey
                        break
                    }
                }
            }
        }

        foreach ($scopeKey in @($scopeLookup.Keys | Sort-Object)) {
            Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
            $scopeRecord = $scopeLookup[$scopeKey]
            $scopeOptions = New-DhcpOptionCells -OptionIds $OptionIds -Source Scope
            $scopeOptionSucceeded = $true
            $reservationReadSucceeded = $true
            $reservations = @()

            if ([bool]$scopeRecord.IsRelevant) {
                $scopeOptionResult = Get-DhcpOptionSet -ComputerName $Master.Name -OptionIds $OptionIds -Level Scope -ScopeId $scopeKey -CancellationToken $CancellationToken
                $remoteCalls += [int]$scopeOptionResult.RemoteCallCount
                $scopeOptions = $scopeOptionResult.Values
                $scopeOptionSucceeded = [bool]$scopeOptionResult.ReadSucceeded
                if (-not $scopeOptionSucceeded) {
                    $completionStatus = 'Partial'
                    [void]$errors.Add("Scope $scopeKey option read failed: $($scopeOptionResult.Error)")
                }

                if ($Mode -eq 'DeepAudit') {
                    try {
                        $remoteCalls++
                        $reservations = @(Get-DhcpServerv4Reservation -ComputerName $Master.Name -ScopeId ([System.Net.IPAddress]::Parse($scopeKey)) -ErrorAction Stop)
                    } catch {
                        if ($CancellationToken.IsCancellationRequested) { throw [System.OperationCanceledException]::new('DHCP validation cancelled during reservation inventory.') }
                        $reservationReadSucceeded = $false
                        $completionStatus = 'Partial'
                        [void]$errors.Add("Scope $scopeKey reservation read failed: $($_.Exception.Message)")
                    }
                }
            }

            [void]$scopesOut.Add([PSCustomObject]@{
                ScopeId = $scopeKey
                StartRange = [string]$scopeRecord.StartRange
                EndRange = [string]$scopeRecord.EndRange
                SubnetMask = [string]$scopeRecord.SubnetMask
                State = [string]$scopeRecord.State
                IsRelevant = [bool]$scopeRecord.IsRelevant
                Options = $scopeOptions
                OptionReadSucceeded = $scopeOptionSucceeded
                ReservationReadSucceeded = $reservationReadSucceeded
                Reservations = @($reservations | ForEach-Object {
                    $clientIdKey = ConvertTo-DhcpClientIdKey -Value $_.ClientId
                    [PSCustomObject]@{
                        IPAddress = [string]$_.IPAddress
                        Name = [string]$_.Name
                        ScopeId = $scopeKey
                        ClientId = $clientIdKey
                        ClientIdKey = $clientIdKey
                    }
                })
            })
        }

        $scopesById = @{}
        foreach ($scope in $scopesOut) { $scopesById[[string]$scope.ScopeId] = $scope }

        if ($Mode -eq 'Targeted') {
            # The normal path queries only selected target IPs in bounded
            # batches. A failed batch is marked incomplete and never broadened
            # into a full-scope reservation enumeration; this is a safety boundary
            # for large production scopes and for authorization/transport errors.
            $targetReservationsByIp = @{}
            $failedTargetReservationIps = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $resolvedTargetIps = @($Targets | Where-Object { $_.ResolutionState -eq 'Resolved' -and $_.IPAddress } | ForEach-Object { [string]$_.IPAddress } | Sort-Object -Unique)
            $reservationCommand = Get-Command Get-DhcpServerv4Reservation -ErrorAction Stop
            for ($offset = 0; $offset -lt $resolvedTargetIps.Count; $offset += 200) {
                Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
                $lastIndex = [Math]::Min($offset + 199, $resolvedTargetIps.Count - 1)
                $chunkStrings = @($resolvedTargetIps[$offset..$lastIndex])
                $reservationParameters = @{
                    ComputerName = $Master.Name
                    IPAddress = [System.Net.IPAddress[]]@($chunkStrings | ForEach-Object { [System.Net.IPAddress]::Parse($_) })
                    ErrorAction = 'Stop'
                }
                if ($reservationCommand.Parameters.ContainsKey('ThrottleLimit')) { $reservationParameters.ThrottleLimit = 1 }
                try {
                    $remoteCalls++
                    foreach ($reservation in @(Get-DhcpServerv4Reservation @reservationParameters)) {
                        $reservationIp = [string]$reservation.IPAddress
                        if ([string]::IsNullOrWhiteSpace($reservationIp)) { continue }
                        $reservationScopeId = if ($reservation.PSObject.Properties.Match('ScopeId').Count -gt 0 -and $reservation.ScopeId) {
                            [string]$reservation.ScopeId
                        } else {
                            $matchingTarget = $Targets | Where-Object { [string]$_.IPAddress -eq $reservationIp } | Select-Object -First 1
                            if ($matchingTarget -and $targetScopeByKey.ContainsKey([string]$matchingTarget.Key)) { [string]$targetScopeByKey[[string]$matchingTarget.Key] } else { '' }
                        }
                        $clientIdKey = ConvertTo-DhcpClientIdKey -Value $reservation.ClientId
                        $targetReservationsByIp[$reservationIp] = [PSCustomObject]@{
                            IPAddress = $reservationIp
                            Name = [string]$reservation.Name
                            ScopeId = $reservationScopeId
                            ClientId = $clientIdKey
                            ClientIdKey = $clientIdKey
                        }
                    }
                } catch {
                    if ($CancellationToken.IsCancellationRequested) { throw [System.OperationCanceledException]::new('DHCP validation cancelled during targeted reservation lookup.') }
                    $targetBatchQueryFailed = $true
                    $completionStatus = 'Partial'
                    foreach ($ip in $chunkStrings) { [void]$failedTargetReservationIps.Add([string]$ip) }
                    [void]$errors.Add("Target reservation batch query failed for $($chunkStrings.Count) addresses; no scope-wide fallback was attempted: $($_.Exception.Message)")
                }
            }

            foreach ($target in $Targets) {
                Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
                $scope = $null
                if ([string]$target.ResolutionState -eq 'Resolved' -and $targetScopeByKey.ContainsKey([string]$target.Key)) {
                    $scope = $scopesById[[string]$targetScopeByKey[[string]$target.Key]]
                }

                $status = 'Unresolved'
                $reservationName = ''
                $reservationIp = ''
                $nameMatchesTarget = $false
                $reservationClientId = ''
                $reservationClientIdKey = ''
                $clientIdMatchesTarget = $null
                $reservationOptions = New-DhcpOptionCells -OptionIds $OptionIds -Source Reservation
                $effectiveOptions = Get-DhcpEffectiveOptionSet -OptionIds $OptionIds -ReservationOptions $reservationOptions -ScopeOptions @{} -ServerOptions $serverOptionResult.Values

                if ([string]$target.ResolutionState -eq 'Resolved' -and $null -eq $scope) {
                    $status = 'ScopeNotFound'
                } elseif ($null -ne $scope -and $failedTargetReservationIps.Contains([string]$target.IPAddress)) {
                    $status = 'Incomplete'
                    $reservationOptions = New-DhcpOptionCells -OptionIds $OptionIds -Source Reservation -InitialState ReadFailed
                    $effectiveOptions = Get-DhcpEffectiveOptionSet -OptionIds $OptionIds -ReservationOptions $reservationOptions -ScopeOptions $scope.Options -ServerOptions $serverOptionResult.Values
                } elseif ($null -ne $scope) {
                    $reservation = if ($targetReservationsByIp.ContainsKey([string]$target.IPAddress)) { $targetReservationsByIp[[string]$target.IPAddress] } else { $null }
                    if (-not $reservation) {
                        $status = 'Missing'
                        $effectiveOptions = Get-DhcpEffectiveOptionSet -OptionIds $OptionIds -ReservationOptions $reservationOptions -ScopeOptions $scope.Options -ServerOptions $serverOptionResult.Values
                    } else {
                        $status = 'Found'
                        $reservationName = [string]$reservation.Name
                        $reservationIp = [string]$reservation.IPAddress
                        $reservationClientId = [string]$reservation.ClientId
                        $reservationClientIdKey = if ($reservation.PSObject.Properties.Match('ClientIdKey').Count -gt 0) { [string]$reservation.ClientIdKey } else { ConvertTo-DhcpClientIdKey -Value $reservation.ClientId }
                        $targetShort = ((([string]$target.Name -replace '^.*\\', '') -split '\.')[0])
                        $reservationShort = ((([string]$reservationName -replace '^.*\\', '') -split '\.')[0])
                        $nameMatchesTarget = ([string]$reservationName -ieq [string]$target.Name) -or ($reservationShort -ieq $targetShort)
                        if (-not [string]::IsNullOrWhiteSpace([string]$target.ExpectedClientId)) {
                            $clientIdMatchesTarget = ($reservationClientIdKey -eq [string]$target.ExpectedClientId)
                        }
                        $reservationOptionResult = Get-DhcpOptionSet -ComputerName $Master.Name -OptionIds $OptionIds -Level Reservation -ReservedIP $reservationIp -CancellationToken $CancellationToken
                        $remoteCalls += [int]$reservationOptionResult.RemoteCallCount
                        $reservationOptions = $reservationOptionResult.Values
                        if (-not [bool]$reservationOptionResult.ReadSucceeded) {
                            $completionStatus = 'Partial'
                            $status = 'Incomplete'
                            [void]$errors.Add("Reservation $reservationIp option read failed: $($reservationOptionResult.Error)")
                        }
                        $effectiveOptions = Get-DhcpEffectiveOptionSet -OptionIds $OptionIds -ReservationOptions $reservationOptions -ScopeOptions $scope.Options -ServerOptions $serverOptionResult.Values
                    }
                }

                [void]$targetResults.Add([PSCustomObject]@{
                    TargetOrdinal = [int]$target.Ordinal
                    TargetKey = [string]$target.Key
                    TargetName = [string]$target.Name
                    TargetIPAddress = [string]$target.IPAddress
                    AddressSource = [string]$target.AddressSource
                    MasterName = [string]$Master.Name
                    Status = $status
                    ScopeId = if ($null -ne $scope) { [string]$scope.ScopeId } else { '' }
                    ReservationName = $reservationName
                    ReservationIPAddress = $reservationIp
                    NameMatchesTarget = $nameMatchesTarget
                    ReservationClientId = $reservationClientId
                    ReservationClientIdKey = $reservationClientIdKey
                    ClientIdMatchesTarget = $clientIdMatchesTarget
                    ExpectedClientId = [string]$target.ExpectedClientId
                    EffectiveOptions = $effectiveOptions
                })
            }
        } elseif ($Mode -eq 'DeepAudit') {
            foreach ($scope in $scopesOut) {
                if (-not [bool]$scope.ReservationReadSucceeded) { continue }
                foreach ($reservation in $scope.Reservations) {
                    Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken
                    $reservationOptionResult = Get-DhcpOptionSet -ComputerName $Master.Name -OptionIds $OptionIds -Level Reservation -ReservedIP ([string]$reservation.IPAddress) -CancellationToken $CancellationToken
                    $remoteCalls += [int]$reservationOptionResult.RemoteCallCount
                    if (-not [bool]$reservationOptionResult.ReadSucceeded) {
                        $completionStatus = 'Partial'
                        [void]$errors.Add("Reservation $($reservation.IPAddress) option read failed: $($reservationOptionResult.Error)")
                    }
                    $effectiveOptions = Get-DhcpEffectiveOptionSet -OptionIds $OptionIds -ReservationOptions $reservationOptionResult.Values -ScopeOptions $scope.Options -ServerOptions $serverOptionResult.Values
                    [void]$deepReservations.Add([PSCustomObject]@{
                        IPAddress = [string]$reservation.IPAddress
                        Name = [string]$reservation.Name
                        ScopeId = [string]$scope.ScopeId
                        ClientId = [string]$reservation.ClientId
                        ClientIdKey = if ($reservation.PSObject.Properties.Match('ClientIdKey').Count -gt 0) { [string]$reservation.ClientIdKey } else { ConvertTo-DhcpClientIdKey -Value $reservation.ClientId }
                        EffectiveOptions = $effectiveOptions
                        OptionReadSucceeded = [bool]$reservationOptionResult.ReadSucceeded
                    })
                }
            }
        }

        if ([string]::IsNullOrWhiteSpace([string]$Master.ExpectedIPv4)) {
            $completionStatus = 'Partial'
            [void]$errors.Add('Master IPv4 address could not be resolved; option 66 validation is unavailable.')
        }

        return [PSCustomObject]@{
            Kind = 'DhcpMasterResult'
            Ordinal = [int]$Master.Ordinal
            MasterName = [string]$Master.Name
            ExpectedIPv4 = [string]$Master.ExpectedIPv4
            CompletionStatus = $completionStatus
            ServiceState = $serviceState
            ServiceQuerySucceeded = $serviceQuerySucceeded
            TargetBatchQueryFailed = $targetBatchQueryFailed
            ServerOptions = $serverOptionResult.Values
            Scopes = @($scopesOut.ToArray())
            TargetResults = @($targetResults.ToArray())
            DeepReservations = @($deepReservations.ToArray())
            Errors = @($errors.ToArray())
            RemoteCallCount = $remoteCalls
            DurationMs = [int]([datetime]::UtcNow - $started).TotalMilliseconds
        }
    } catch [System.OperationCanceledException] {
        throw
    } catch {
        [void]$errors.Add($_.Exception.Message)
        return [PSCustomObject]@{
            Kind = 'DhcpMasterResult'; Ordinal = [int]$Master.Ordinal; MasterName = [string]$Master.Name
            ExpectedIPv4 = [string]$Master.ExpectedIPv4; CompletionStatus = 'Unavailable'; ServiceState = $serviceState
            ServiceQuerySucceeded = $serviceQuerySucceeded
            ServerOptions = @{}; Scopes = @($scopesOut.ToArray()); TargetResults = @($targetResults.ToArray())
            DeepReservations = @($deepReservations.ToArray()); Errors = @($errors.ToArray())
            RemoteCallCount = $remoteCalls; DurationMs = [int]([datetime]::UtcNow - $started).TotalMilliseconds
        }
    }
}

function Get-DhcpMasterWorkerScriptText {
    $functionNames = @(Get-CitrixWorkerDependencyManifest).DhcpMaster
    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.AppendLine('param($Master, $Mode, $Targets, $OptionIds, $CancellationToken)')
    foreach ($functionName in $functionNames) {
        $command = Get-Command -Name $functionName -CommandType Function -ErrorAction Stop
        [void]$builder.AppendLine("function $functionName {")
        [void]$builder.AppendLine($command.Definition)
        [void]$builder.AppendLine('}')
    }
    [void]$builder.AppendLine('Get-PvsDhcpMasterSnapshot -Master $Master -Mode $Mode -Targets $Targets -OptionIds $OptionIds -CancellationToken $CancellationToken')
    return $builder.ToString()
}

function New-DhcpTimedOutMasterResult {
    param(
        [object]$Master,
        [int]$DurationMs,
        [string]$Message
    )

    return [PSCustomObject]@{
        Kind = 'DhcpMasterResult'
        Ordinal = [int]$Master.Ordinal
        MasterName = [string]$Master.Name
        ExpectedIPv4 = [string]$Master.ExpectedIPv4
        CompletionStatus = 'TimedOut'
        ServiceState = 'Unknown'
        ServiceQuerySucceeded = $false
        ServerOptions = @{}
        Scopes = @()
        TargetResults = @()
        DeepReservations = @()
        Errors = @($Message)
        RemoteCallCount = $null
        DurationMs = $DurationMs
    }
}

function Invoke-DhcpMasterCollection {
    param(
        [object[]]$Masters,
        [ValidateSet('Baseline','Targeted','DeepAudit')]
        [string]$Mode,
        [object[]]$Targets,
        [int[]]$OptionIds,
        [ValidateRange(1,3)]
        [int]$MasterThrottle = 2,
        [ValidateRange(10,1800)]
        [int]$MasterSoftTimeoutSeconds = 90,
        [ValidateRange(1,3600)]
        [int]$OverallSoftTimeoutSeconds = 900,
        [object]$MessageQueue,
        [System.Threading.CancellationToken]$CancellationToken = [System.Threading.CancellationToken]::None
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $pending = [System.Collections.Generic.Queue[object]]::new()
    foreach ($master in @($Masters | Sort-Object Ordinal)) { $pending.Enqueue($master) }
    $active = [System.Collections.Generic.List[object]]::new()
    $pool = $null
    $overallTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $overallTimedOut = $false
    $workerScript = Get-DhcpMasterWorkerScriptText

    try {
        $pool = [RunspaceFactory]::CreateRunspacePool(1, $MasterThrottle)
        $pool.ApartmentState = [System.Threading.ApartmentState]::MTA
        $pool.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $pool.Open()

        while ($pending.Count -gt 0 -or $active.Count -gt 0) {
            Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken -Message 'DHCP validation cancelled by the user.'

            if (-not $overallTimedOut -and $overallTimer.Elapsed.TotalSeconds -ge $OverallSoftTimeoutSeconds) {
                $overallTimedOut = $true
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "DHCP overall soft deadline of $OverallSoftTimeoutSeconds seconds reached; stop requested for active Master queries." -Level WARN
                while ($pending.Count -gt 0) {
                    $notStarted = $pending.Dequeue()
                    [void]$results.Add((New-DhcpTimedOutMasterResult -Master $notStarted -DurationMs 0 -Message 'Not started before the overall DHCP deadline.'))
                }
            }

            while (-not $overallTimedOut -and $pending.Count -gt 0 -and $active.Count -lt $MasterThrottle) {
                $master = $pending.Dequeue()
                $pipeline = [PowerShell]::Create()
                $pipeline.RunspacePool = $pool
                $workerCancellationSource = [System.Threading.CancellationTokenSource]::CreateLinkedTokenSource(
                    $CancellationToken,
                    [System.Threading.CancellationToken]::None
                )
                try {
                    [void]$pipeline.AddScript($workerScript).AddArgument($master).AddArgument($Mode).AddArgument(@($Targets)).AddArgument([int[]]$OptionIds).AddArgument($workerCancellationSource.Token)
                    $handle = $pipeline.BeginInvoke()
                } catch {
                    $workerCancellationSource.Dispose()
                    $pipeline.Dispose()
                    throw
                }
                [void]$active.Add([PSCustomObject]@{
                    Master = $master
                    Pipeline = $pipeline
                    Handle = $handle
                    StopHandle = $null
                    WorkerCancellationSource = $workerCancellationSource
                    Timer = [System.Diagnostics.Stopwatch]::StartNew()
                    TimeoutRequested = $false
                })
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message "DHCP ${Mode}: querying $($master.Name) ($($results.Count + $active.Count) of $($Masters.Count))..." -Current $results.Count -Total $Masters.Count -IsIndeterminate $false
            }

            for ($index = $active.Count - 1; $index -ge 0; $index--) {
                $context = $active[$index]
                if (($overallTimedOut -or $context.Timer.Elapsed.TotalSeconds -ge $MasterSoftTimeoutSeconds) -and -not [bool]$context.TimeoutRequested) {
                    $context.TimeoutRequested = $true
                    $deadlineLabel = if ($overallTimedOut) { 'overall deadline' } else { "$MasterSoftTimeoutSeconds-second Master deadline" }
                    Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "DHCP $deadlineLabel reached for $($context.Master.Name); cooperative stop requested." -Level WARN
                    try { $context.WorkerCancellationSource.Cancel() } catch { }
                    try { $context.StopHandle = $context.Pipeline.BeginStop($null, $null) } catch { }
                }

                if (-not $context.Handle.IsCompleted) { continue }

                $masterResult = $null
                try {
                    $output = @($context.Pipeline.EndInvoke($context.Handle))
                    if (-not [bool]$context.TimeoutRequested) {
                        $masterResult = $output | Where-Object { $_.Kind -eq 'DhcpMasterResult' } | Select-Object -Last 1
                    }
                } catch {
                    if (-not [bool]$context.TimeoutRequested) {
                        $masterResult = [PSCustomObject]@{
                            Kind = 'DhcpMasterResult'; Ordinal = [int]$context.Master.Ordinal; MasterName = [string]$context.Master.Name
                            ExpectedIPv4 = [string]$context.Master.ExpectedIPv4; CompletionStatus = 'Unavailable'; ServiceState = 'Unavailable'
                            ServerOptions = @{}; Scopes = @(); TargetResults = @(); DeepReservations = @()
                            Errors = @($_.Exception.Message); RemoteCallCount = 0; DurationMs = [int]$context.Timer.ElapsedMilliseconds
                        }
                    }
                }

                if ([bool]$context.TimeoutRequested) {
                    $masterResult = New-DhcpTimedOutMasterResult -Master $context.Master -DurationMs ([int]$context.Timer.ElapsedMilliseconds) -Message 'The DHCP query exceeded its soft deadline; stop was requested.'
                } elseif ($null -eq $masterResult) {
                    $masterResult = [PSCustomObject]@{
                        Kind = 'DhcpMasterResult'; Ordinal = [int]$context.Master.Ordinal; MasterName = [string]$context.Master.Name
                        ExpectedIPv4 = [string]$context.Master.ExpectedIPv4; CompletionStatus = 'Unavailable'; ServiceState = 'Unavailable'
                        ServerOptions = @{}; Scopes = @(); TargetResults = @(); DeepReservations = @()
                        Errors = @('The Master worker returned no result.'); RemoteCallCount = 0; DurationMs = [int]$context.Timer.ElapsedMilliseconds
                    }
                }

                [void]$results.Add($masterResult)
                Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Log -Message "DHCP Master $($masterResult.MasterName) completed with status $($masterResult.CompletionStatus), calls=$($masterResult.RemoteCallCount), durationMs=$($masterResult.DurationMs)" -Level $(if ($masterResult.CompletionStatus -eq 'Complete') { 'INFO' } else { 'WARN' })

                if ($null -ne $context.StopHandle -and $context.StopHandle.IsCompleted) {
                    try { $context.Pipeline.EndStop($context.StopHandle) } catch { }
                }
                $context.Pipeline.Dispose()
                $context.WorkerCancellationSource.Dispose()
                $context.Timer.Stop()
                $active.RemoveAt($index)
            }

            if ($pending.Count -gt 0 -or $active.Count -gt 0) { Start-Sleep -Milliseconds 100 }
        }
    } finally {
        foreach ($context in @($active)) {
            try { $context.WorkerCancellationSource.Cancel() } catch { }
            if (-not $context.Handle.IsCompleted) {
                try { [void]$context.Pipeline.BeginStop($null, $null) } catch { }
            }
        }
        if ($null -ne $pool) {
            try { $pool.Close() } catch { }
        }
        foreach ($context in @($active)) {
            if ($context.Handle.IsCompleted) {
                try { [void]$context.Pipeline.EndInvoke($context.Handle) } catch { }
            }
            $context.Pipeline.Dispose()
            $context.WorkerCancellationSource.Dispose()
            $context.Timer.Stop()
        }
        if ($null -ne $pool) { $pool.Dispose() }
        $overallTimer.Stop()
    }

    return [PSCustomObject]@{
        Results = @($results.ToArray() | Sort-Object Ordinal)
        OverallTimedOut = $overallTimedOut
        DurationMs = [int]$overallTimer.ElapsedMilliseconds
    }
}

function Get-DhcpOptionCellSignature {
    param([object]$Cell)

    if ($null -eq $Cell) { return 'ReadFailed:' }
    if ([string]$Cell.State -eq 'Value') { return "Value:$([string]$Cell.Value)" }
    return "$([string]$Cell.State):"
}

function Get-DhcpOptionCellDisplayText {
    param([object]$Cell)

    if ($null -eq $Cell) { return 'ReadFailed' }
    switch ([string]$Cell.State) {
        'Value' { return [string]$Cell.Value }
        'Unset' { return 'Unset' }
        default { return 'ReadFailed' }
    }
}

function Invoke-PvsDhcpValidation {
    param(
        [ValidateSet('Baseline','Targeted','DeepAudit')]
        [string]$Mode,
        [object[]]$Devices,
        [hashtable]$Config,
        [object]$MessageQueue,
        [System.Threading.CancellationToken]$CancellationToken,
        [System.Threading.ManualResetEventSlim]$ApprovalGate,
        [hashtable]$ApprovalState
    )

    $overallSoftTimeout = if ($Config.ContainsKey('OverallSoftTimeoutSeconds')) { [int]$Config.OverallSoftTimeoutSeconds } else { 900 }
    $overallBudgetTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $initialOverallDeadlineUtc = [datetime]::UtcNow.AddSeconds($overallSoftTimeout)
    $optionMap = Get-DhcpValidationOptionMap
    $optionIds = [int[]]@($optionMap.Keys | Sort-Object)
    $findings = [System.Collections.Generic.List[string]]::new()
    $findingState = @{ Count = 0; ValidationIssues = 0; CoverageIssues = 0 }
    $maxFindings = if ($Config.ContainsKey('MaxFindings')) { [int]$Config.MaxFindings } else { 250 }
    $addFinding = {
        param([string]$Text, [ValidateSet('Validation','Coverage')] [string]$Category = 'Validation')
        $findingState.Count++
        if ($Category -eq 'Validation') { $findingState.ValidationIssues++ } else { $findingState.CoverageIssues++ }
        if ($findings.Count -lt $maxFindings) { [void]$findings.Add($Text) }
    }

    Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind Progress -Message "Preparing DHCP $Mode validation..." -IsIndeterminate $true
    $masterResolution = Resolve-PvsMastersForDhcp -MessageQueue $MessageQueue -DeadlineUtc $initialOverallDeadlineUtc -CancellationToken $CancellationToken
    $masters = @($masterResolution.Masters)
    if ([bool]$masterResolution.TimedOut -or $overallBudgetTimer.Elapsed.TotalSeconds -ge $overallSoftTimeout) {
        $overallBudgetTimer.Stop()
        return [PSCustomObject]@{
            Requested = $true; Mode = $Mode; Outcome = 'TimedOut'; CompletionStatus = 'TimedOut'; IsComplete = $false
            ValidationStatus = 'Not comparable'; MastersDiscovered = $masters.Count; MastersComplete = 0; MasterResults = @(); TargetResults = @()
            TargetCount = 0; ResolvedTargetCount = 0; UnresolvedTargetCount = 0; ReservationCount = 0
            Findings = @("DHCP validation reached its $overallSoftTimeout-second overall soft deadline during PVS Master discovery.")
            FindingCount = 1; FindingsOmitted = 0; RemoteCallCount = 0; CallCountIsLowerBound = $false
            DurationMs = [int]$overallBudgetTimer.ElapsedMilliseconds
            ErrorMessage = 'The overall DHCP soft deadline was reached during PVS Master discovery.'
        }
    }
    if ($masters.Count -eq 0) {
        return [PSCustomObject]@{
            Requested = $true; Mode = $Mode; Outcome = 'Unavailable'; CompletionStatus = 'Unavailable'; IsComplete = $false
            ValidationStatus = 'Not comparable'; MastersDiscovered = 0; MastersComplete = 0; MasterResults = @(); TargetResults = @()
            TargetCount = 0; ResolvedTargetCount = 0; UnresolvedTargetCount = 0; ReservationCount = 0
            Findings = @('No PVS Master could be resolved for DHCP validation.'); FindingCount = 1; FindingsOmitted = 0
            RemoteCallCount = 0; DurationMs = [int]$overallBudgetTimer.ElapsedMilliseconds; ErrorMessage = 'No PVS Master resolved.'
        }
    }

    $targetSet = [PSCustomObject]@{ Targets = @(); Warnings = @(); ResolvedCount = 0; UnresolvedCount = 0; TimedOut = $false }
    if ($Mode -eq 'Targeted') {
        $matchedTargetCount = @($Devices).Count
        if ($matchedTargetCount -eq 0) {
            return [PSCustomObject]@{
                Requested = $true; Mode = $Mode; Outcome = 'NotRun'; CompletionStatus = 'NotRun'; IsComplete = $false
                ValidationStatus = 'Not comparable'; MastersDiscovered = $masters.Count; MastersComplete = 0; MasterResults = @(); TargetResults = @()
                TargetCount = 0; ResolvedTargetCount = 0; UnresolvedTargetCount = 0; ReservationCount = 0
                Findings = @('No matched PVS target was available for targeted DHCP validation.'); FindingCount = 1; FindingsOmitted = 0
                RemoteCallCount = 0; DurationMs = [int]$overallBudgetTimer.ElapsedMilliseconds; ErrorMessage = ''
            }
        }
        $maximumTargetCount = if ($Config.ContainsKey('MaximumTargetCount')) { [int]$Config.MaximumTargetCount } else { 1000 }
        if ($matchedTargetCount -gt $maximumTargetCount) {
            return [PSCustomObject]@{
                Requested = $true; Mode = $Mode; Outcome = 'NotRun'; CompletionStatus = 'NotRun'; IsComplete = $false
                ValidationStatus = 'Not run'; MastersDiscovered = $masters.Count; MastersComplete = 0; MasterResults = @(); TargetResults = @()
                TargetCount = $matchedTargetCount; ResolvedTargetCount = 0; UnresolvedTargetCount = 0; ReservationCount = 0
                Findings = @("Targeted DHCP validation is limited to $maximumTargetCount matched servers per run. Use smaller Server List batches, or run the separately approved Whole Farm Deep Audit.")
                FindingCount = 1; FindingsOmitted = 0; RemoteCallCount = 0
                DurationMs = [int]$overallBudgetTimer.ElapsedMilliseconds; ErrorMessage = ''
            }
        }

        # Ask before any per-target DNS fallback. This keeps a declined large
        # request cheap even when many device addresses are absent from PVS.
        $targetMasterCells = $matchedTargetCount * $masters.Count
        $maximumTargetMasterCells = if ($Config.ContainsKey('MaximumTargetMasterCells')) { [int]$Config.MaximumTargetMasterCells } else { 5000 }
        if ($targetMasterCells -gt $maximumTargetMasterCells) {
            return [PSCustomObject]@{
                Requested = $true; Mode = $Mode; Outcome = 'NotRun'; CompletionStatus = 'NotRun'; IsComplete = $false
                ValidationStatus = 'Not run'; MastersDiscovered = $masters.Count; MastersComplete = 0; MasterResults = @(); TargetResults = @()
                TargetCount = $matchedTargetCount; ResolvedTargetCount = 0; UnresolvedTargetCount = 0; ReservationCount = 0
                Findings = @("Targeted DHCP validation was not started because $matchedTargetCount targets across $($masters.Count) PVS Masters create $targetMasterCells target-by-Master checks; the production safety limit is $maximumTargetMasterCells. Use smaller Server List or Device Collection batches.")
                FindingCount = 1; FindingsOmitted = 0; RemoteCallCount = 0
                DurationMs = [int]$overallBudgetTimer.ElapsedMilliseconds; ErrorMessage = ''
            }
        }
        $confirmationThreshold = if ($Config.ContainsKey('TargetConfirmationThreshold')) { [int]$Config.TargetConfirmationThreshold } else { 1000 }
        if ($targetMasterCells -ge $confirmationThreshold) {
            $ApprovalState.Approved = $false
            $ApprovalGate.Reset()
            # Operator decision time is not query execution time and therefore
            # does not consume the DHCP preparation/collection budget.
            $overallBudgetTimer.Stop()
            Send-RetrievalWorkerMessage -MessageQueue $MessageQueue -Kind WorkloadConfirmation -Message 'Large targeted DHCP validation confirmation required.' -Details @{
                ConfirmationType = 'DhcpTargeted'
                TargetCount = $matchedTargetCount
                MasterCount = $masters.Count
                TargetMasterCells = $targetMasterCells
            }
            try {
                while (-not $ApprovalGate.Wait(100)) {
                    Test-RetrievalWorkerCancellation -CancellationToken $CancellationToken -Message 'Retrieval cancelled before targeted DHCP validation.'
                }
            } finally {
                $overallBudgetTimer.Start()
            }
            if (-not [bool]$ApprovalState.Approved) {
                return [PSCustomObject]@{
                    Requested = $true; Mode = $Mode; Outcome = 'Declined'; CompletionStatus = 'NotRun'; IsComplete = $false
                    ValidationStatus = 'Not run'; MastersDiscovered = $masters.Count; MastersComplete = 0; MasterResults = @(); TargetResults = @()
                    TargetCount = $matchedTargetCount; ResolvedTargetCount = 0; UnresolvedTargetCount = 0
                    ReservationCount = 0; Findings = @('The operator declined the large targeted DHCP validation.'); FindingCount = 1; FindingsOmitted = 0
                    RemoteCallCount = 0; DurationMs = [int]$overallBudgetTimer.ElapsedMilliseconds; ErrorMessage = ''
                }
            }
        }

        $resolutionSecondsRemaining = [Math]::Max(0, $overallSoftTimeout - $overallBudgetTimer.Elapsed.TotalSeconds)
        $resolutionDeadlineUtc = [datetime]::UtcNow.AddSeconds($resolutionSecondsRemaining)
        $targetSet = Resolve-DhcpTargetSet -Devices $Devices -MessageQueue $MessageQueue -DeadlineUtc $resolutionDeadlineUtc -CancellationToken $CancellationToken
        foreach ($warning in @($targetSet.Warnings)) { & $addFinding $warning 'Coverage' }
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$masterResolution.EnumerationError)) {
        & $addFinding "PVS Master enumeration was incomplete: $($masterResolution.EnumerationError)" 'Coverage'
    }

    $masterThrottle = if ($Config.ContainsKey('MasterThrottle')) { [int]$Config.MasterThrottle } else { 2 }
    $masterTimeout = if ($Config.ContainsKey('MasterSoftTimeoutSeconds')) { [int]$Config.MasterSoftTimeoutSeconds } else { if ($Mode -eq 'DeepAudit') { 300 } else { 90 } }
    if ($Mode -eq 'Targeted') {
        $recommendedMasterTimeout = [Math]::Min(600, [Math]::Max(90, 60 + (2 * @($targetSet.Targets).Count)))
        $masterTimeout = [Math]::Max($masterTimeout, $recommendedMasterTimeout)
    }

    $remainingOverallSeconds = [int][Math]::Floor($overallSoftTimeout - $overallBudgetTimer.Elapsed.TotalSeconds)
    if ([bool]$targetSet.TimedOut -or $remainingOverallSeconds -lt 1) {
        & $addFinding "DHCP validation reached its $overallSoftTimeout-second overall soft deadline during preparation; no Master comparison was started." 'Coverage'
        $overallBudgetTimer.Stop()
        return [PSCustomObject]@{
            Requested = $true; Mode = $Mode; Outcome = 'TimedOut'; CompletionStatus = 'TimedOut'; IsComplete = $false
            ValidationStatus = 'Not comparable'; MastersDiscovered = $masters.Count; MastersComplete = 0; MasterResults = @(); TargetResults = @()
            TargetCount = if ($Mode -eq 'Targeted') { $matchedTargetCount } else { 0 }
            ResolvedTargetCount = [int]$targetSet.ResolvedCount; UnresolvedTargetCount = [int]$targetSet.UnresolvedCount
            ReservationCount = 0; Findings = @($findings.ToArray()); FindingCount = [int]$findingState.Count
            FindingsOmitted = [Math]::Max(0, [int]$findingState.Count - $findings.Count); RemoteCallCount = 0
            CallCountIsLowerBound = $false; DurationMs = [int]$overallBudgetTimer.ElapsedMilliseconds
            ErrorMessage = 'The overall DHCP soft deadline was reached during preparation.'
        }
    }

    $collection = Invoke-DhcpMasterCollection -Masters $masters -Mode $Mode -Targets @($targetSet.Targets) -OptionIds $optionIds -MasterThrottle $masterThrottle -MasterSoftTimeoutSeconds $masterTimeout -OverallSoftTimeoutSeconds $remainingOverallSeconds -MessageQueue $MessageQueue -CancellationToken $CancellationToken
    $snapshots = @($collection.Results)

    foreach ($snapshot in $snapshots) {
        if ([string]$snapshot.CompletionStatus -ne 'Complete') {
            $snapshotErrors = @($snapshot.Errors)
            $errorText = @($snapshotErrors | Select-Object -First 5) -join ' | '
            if ($snapshotErrors.Count -gt 5) { $errorText += " | plus $($snapshotErrors.Count - 5) additional errors" }
            if ([string]::IsNullOrWhiteSpace($errorText)) { $errorText = 'Required data was incomplete.' }
            & $addFinding "$($snapshot.MasterName): $($snapshot.CompletionStatus) - $errorText" 'Coverage'
        }
        if ($snapshot.PSObject.Properties.Match('ServiceQuerySucceeded').Count -gt 0 -and -not [bool]$snapshot.ServiceQuerySucceeded) {
            & $addFinding "$($snapshot.MasterName): DHCP Server service state could not be read; configuration data is compared separately." 'Coverage'
        } elseif ([string]$snapshot.ServiceState -ne 'Running') {
            & $addFinding "$($snapshot.MasterName): DHCP Server service state is '$($snapshot.ServiceState)'." 'Validation'
        }
    }

    $completeSnapshots = @($snapshots | Where-Object { [string]$_.CompletionStatus -eq 'Complete' })
    $comparableSnapshots = $completeSnapshots
    # A cross-Master clean statement requires every discovered Master and at
    # least two readable Masters. A readable subset can still expose issues, but
    # it cannot authoritatively prove consistency.
    $notComparable = ($comparableSnapshots.Count -lt $masters.Count -or $comparableSnapshots.Count -lt 2)

    if ($Mode -in @('Baseline','DeepAudit')) {
        $scopeKeys = @($comparableSnapshots | ForEach-Object { $_.Scopes | ForEach-Object { [string]$_.ScopeId } } | Sort-Object -Unique)
        foreach ($scopeKey in $scopeKeys) {
            $scopeEntries = [System.Collections.Generic.List[object]]::new()
            foreach ($snapshot in $comparableSnapshots) {
                $scope = $snapshot.Scopes | Where-Object { [string]$_.ScopeId -eq $scopeKey } | Select-Object -First 1
                if ($null -eq $scope) {
                    & $addFinding "Scope $scopeKey is missing on $($snapshot.MasterName)." 'Validation'
                } else {
                    [void]$scopeEntries.Add([PSCustomObject]@{ Master = $snapshot; Scope = $scope })
                }
            }

            foreach ($propertyName in @('StartRange','EndRange','SubnetMask','State')) {
                $pairs = [System.Collections.Generic.List[string]]::new()
                $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($entry in $scopeEntries) {
                    $value = [string]$entry.Scope.$propertyName
                    [void]$pairs.Add("$($entry.Master.MasterName)='$value'")
                    [void]$values.Add($value)
                }
                if ($values.Count -gt 1) {
                    & $addFinding "Scope $scopeKey $propertyName mismatch: $($pairs -join ' | ')" 'Validation'
                }
            }

            foreach ($optionId in @(3,6,11,15,67)) {
                $pairs = [System.Collections.Generic.List[string]]::new()
                $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($entry in $scopeEntries) {
                    $cell = $entry.Scope.Options[[string]$optionId]
                    $signature = Get-DhcpOptionCellSignature -Cell $cell
                    [void]$pairs.Add("$($entry.Master.MasterName)='$(Get-DhcpOptionCellDisplayText -Cell $cell)'")
                    [void]$values.Add($signature)
                }
                if ($values.Count -gt 1) {
                    & $addFinding "Scope $scopeKey option $optionId ($($optionMap[$optionId])) mismatch: $($pairs -join ' | ')" 'Validation'
                }
            }

            foreach ($entry in $scopeEntries) {
                $effectiveScopeOptions = Get-DhcpEffectiveOptionSet -OptionIds $optionIds -ReservationOptions (New-DhcpOptionCells -OptionIds $optionIds -Source Reservation) -ScopeOptions $entry.Scope.Options -ServerOptions $entry.Master.ServerOptions
                $cell66 = $effectiveScopeOptions['66']
                if ([string]$cell66.State -ne 'Value' -or [string]$cell66.Value -ne [string]$entry.Master.ExpectedIPv4) {
                    & $addFinding "Scope $scopeKey effective default option 66 on $($entry.Master.MasterName) expected '$($entry.Master.ExpectedIPv4)' but found '$(Get-DhcpOptionCellDisplayText -Cell $cell66)'." 'Validation'
                }
            }
        }

        foreach ($optionId in @(3,6,11,15,67)) {
            $pairs = [System.Collections.Generic.List[string]]::new()
            $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($snapshot in $comparableSnapshots) {
                $cell = $snapshot.ServerOptions[[string]$optionId]
                [void]$pairs.Add("$($snapshot.MasterName)='$(Get-DhcpOptionCellDisplayText -Cell $cell)'")
                [void]$values.Add((Get-DhcpOptionCellSignature -Cell $cell))
            }
            if ($values.Count -gt 1) {
                & $addFinding "Server option $optionId ($($optionMap[$optionId])) mismatch: $($pairs -join ' | ')" 'Validation'
            }
        }
        foreach ($snapshot in $comparableSnapshots) {
            $cell66 = $snapshot.ServerOptions['66']
            if ([string]$cell66.State -eq 'Value' -and [string]$cell66.Value -ne [string]$snapshot.ExpectedIPv4) {
                & $addFinding "Server option 66 on $($snapshot.MasterName) expected '$($snapshot.ExpectedIPv4)' but found '$($cell66.Value)'." 'Validation'
            } elseif (@($snapshot.Scopes).Count -eq 0 -and [string]$cell66.State -ne 'Value') {
                & $addFinding "Default option 66 is not configured on $($snapshot.MasterName), and no scopes were available for an effective-value check." 'Validation'
            }
        }
    }

    if ($Mode -eq 'Targeted') {
        $targetResultsByMaster = @{}
        $masterByName = @{}
        foreach ($snapshot in $comparableSnapshots) {
            $masterKey = ([string]$snapshot.MasterName).ToLowerInvariant()
            $masterByName[$masterKey] = $snapshot
            $targetMap = @{}
            foreach ($targetResult in @($snapshot.TargetResults)) {
                $targetMap[[string]$targetResult.TargetKey] = $targetResult
            }
            $targetResultsByMaster[$masterKey] = $targetMap
        }

        foreach ($target in @($targetSet.Targets | Sort-Object Ordinal)) {
            $entries = [System.Collections.Generic.List[object]]::new()
            foreach ($snapshot in $comparableSnapshots) {
                $masterKey = ([string]$snapshot.MasterName).ToLowerInvariant()
                if ($targetResultsByMaster[$masterKey].ContainsKey([string]$target.Key)) {
                    [void]$entries.Add($targetResultsByMaster[$masterKey][[string]$target.Key])
                } else {
                    $notComparable = $true
                    & $addFinding "$($target.Name): no target result was returned by $($snapshot.MasterName)." 'Coverage'
                }
            }
            if ([string]$target.ResolutionState -ne 'Resolved') {
                $notComparable = $true
                continue
            }
            foreach ($entry in $entries) {
                switch ([string]$entry.Status) {
                    'Missing' {
                        & $addFinding "$($target.Name): reservation $($target.IPAddress) is missing on $($entry.MasterName)." 'Validation'
                    }
                    'ScopeNotFound' { & $addFinding "$($target.Name): no containing DHCP scope was found on $($entry.MasterName)." 'Validation' }
                    'Incomplete' { & $addFinding "$($target.Name): result is incomplete on $($entry.MasterName)." 'Coverage' }
                    'Unresolved' { & $addFinding "$($target.Name): IPv4 resolution failed." 'Coverage' }
                }
            }

            $foundEntries = @($entries | Where-Object { [string]$_.Status -eq 'Found' })
            if ($foundEntries.Count -eq 0) {
                continue
            }

            $reservationNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $namePairs = [System.Collections.Generic.List[string]]::new()
            $clientIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $clientIdPairs = [System.Collections.Generic.List[string]]::new()
            foreach ($entry in $foundEntries) {
                [void]$reservationNames.Add([string]$entry.ReservationName)
                [void]$namePairs.Add("$($entry.MasterName)='$($entry.ReservationName)'")
                [void]$clientIds.Add([string]$entry.ReservationClientIdKey)
                [void]$clientIdPairs.Add("$($entry.MasterName)='$($entry.ReservationClientId)'")
                if ($entry.PSObject.Properties.Match('NameMatchesTarget').Count -gt 0 -and -not [bool]$entry.NameMatchesTarget) {
                    & $addFinding "$($target.Name): reservation $($target.IPAddress) on $($entry.MasterName) is named '$($entry.ReservationName)', which does not match the target name." 'Validation'
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$entry.ExpectedClientId) -and -not [bool]$entry.ClientIdMatchesTarget) {
                    & $addFinding "$($target.Name): reservation client ID on $($entry.MasterName) does not match the PVS device MAC/client ID." 'Validation'
                }
            }
            if ($reservationNames.Count -gt 1) {
                & $addFinding "$($target.Name): reservation name mismatch: $($namePairs -join ' | ')" 'Validation'
            }
            if ($clientIds.Count -gt 1) {
                & $addFinding "$($target.Name): reservation client ID mismatch: $($clientIdPairs -join ' | ')" 'Validation'
            }

            foreach ($optionId in @(3,6,11,15,67)) {
                $pairs = [System.Collections.Generic.List[string]]::new()
                $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($entry in $foundEntries) {
                    $cell = $entry.EffectiveOptions[[string]$optionId]
                    [void]$pairs.Add("$($entry.MasterName)='$(Get-DhcpOptionCellDisplayText -Cell $cell)'")
                    [void]$values.Add((Get-DhcpOptionCellSignature -Cell $cell))
                }
                if ($values.Count -gt 1) {
                    & $addFinding "$($target.Name): effective default option $optionId ($($optionMap[$optionId])) mismatch: $($pairs -join ' | ')" 'Validation'
                }
            }
            foreach ($entry in $foundEntries) {
                $master = $masterByName[([string]$entry.MasterName).ToLowerInvariant()]
                $cell66 = $entry.EffectiveOptions['66']
                if ([string]$cell66.State -ne 'Value' -or [string]$cell66.Value -ne [string]$master.ExpectedIPv4) {
                    & $addFinding "$($target.Name): effective default option 66 on $($entry.MasterName) expected '$($master.ExpectedIPv4)' but found '$(Get-DhcpOptionCellDisplayText -Cell $cell66)'." 'Validation'
                }
            }
        }
    } elseif ($Mode -eq 'DeepAudit') {
        # Build one reservation dictionary per Master. This keeps full-farm
        # comparison O(M*R) instead of repeatedly scanning each reservation list.
        $deepReservationsByMaster = @{}
        $reservationKeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($snapshot in $comparableSnapshots) {
            $masterKey = ([string]$snapshot.MasterName).ToLowerInvariant()
            $reservationMap = @{}
            foreach ($reservation in @($snapshot.DeepReservations)) {
                $reservationKey = [string]$reservation.IPAddress
                if ([string]::IsNullOrWhiteSpace($reservationKey)) { continue }
                $reservationMap[$reservationKey] = $reservation
                [void]$reservationKeySet.Add($reservationKey)
            }
            $deepReservationsByMaster[$masterKey] = $reservationMap
        }
        $reservationKeys = @($reservationKeySet | Sort-Object)
        foreach ($reservationKey in $reservationKeys) {
            $entries = [System.Collections.Generic.List[object]]::new()
            foreach ($snapshot in $comparableSnapshots) {
                $masterKey = ([string]$snapshot.MasterName).ToLowerInvariant()
                $reservation = if ($deepReservationsByMaster[$masterKey].ContainsKey($reservationKey)) { $deepReservationsByMaster[$masterKey][$reservationKey] } else { $null }
                if ($null -eq $reservation) {
                    & $addFinding "Reservation $reservationKey is missing on $($snapshot.MasterName)." 'Validation'
                } else {
                    [void]$entries.Add([PSCustomObject]@{ Master = $snapshot; Reservation = $reservation })
                }
            }
            if ($entries.Count -eq 0) { continue }

            $names = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $namePairs = [System.Collections.Generic.List[string]]::new()
            $clientIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $clientIdPairs = [System.Collections.Generic.List[string]]::new()
            foreach ($entry in $entries) {
                [void]$names.Add([string]$entry.Reservation.Name)
                [void]$namePairs.Add("$($entry.Master.MasterName)='$($entry.Reservation.Name)'")
                [void]$clientIds.Add([string]$entry.Reservation.ClientIdKey)
                [void]$clientIdPairs.Add("$($entry.Master.MasterName)='$($entry.Reservation.ClientId)'")
            }
            if ($names.Count -gt 1) {
                & $addFinding "Reservation $reservationKey name mismatch: $($namePairs -join ' | ')" 'Validation'
            }
            if ($clientIds.Count -gt 1) {
                & $addFinding "Reservation $reservationKey client ID mismatch: $($clientIdPairs -join ' | ')" 'Validation'
            }

            foreach ($optionId in @(3,6,11,15,67)) {
                $pairs = [System.Collections.Generic.List[string]]::new()
                $values = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($entry in $entries) {
                    $cell = $entry.Reservation.EffectiveOptions[[string]$optionId]
                    [void]$pairs.Add("$($entry.Master.MasterName)='$(Get-DhcpOptionCellDisplayText -Cell $cell)'")
                    [void]$values.Add((Get-DhcpOptionCellSignature -Cell $cell))
                }
                if ($values.Count -gt 1) {
                    & $addFinding "Reservation $reservationKey effective default option $optionId ($($optionMap[$optionId])) mismatch: $($pairs -join ' | ')" 'Validation'
                }
            }
            foreach ($entry in $entries) {
                $cell66 = $entry.Reservation.EffectiveOptions['66']
                if ([string]$cell66.State -ne 'Value' -or [string]$cell66.Value -ne [string]$entry.Master.ExpectedIPv4) {
                    & $addFinding "Reservation $reservationKey effective default option 66 on $($entry.Master.MasterName) expected '$($entry.Master.ExpectedIPv4)' but found '$(Get-DhcpOptionCellDisplayText -Cell $cell66)'." 'Validation'
                }
            }
        }
    }

    $completionStatus = if ([bool]$collection.OverallTimedOut -or ($snapshots.Count -gt 0 -and @($snapshots | Where-Object { $_.CompletionStatus -eq 'TimedOut' }).Count -eq $snapshots.Count)) {
        'TimedOut'
    } elseif ($completeSnapshots.Count -eq 0) {
        'Unavailable'
    } elseif ($completeSnapshots.Count -lt $masters.Count -or $findingState.CoverageIssues -gt 0) {
        'Partial'
    } else {
        'Complete'
    }
    if ($findingState.CoverageIssues -gt 0) { $notComparable = $true }
    # Findings from incomplete snapshots remain visible, but they cannot replace
    # the primary coverage conclusion. Only complete authoritative coverage may
    # report either Issues detected or No differences found.
    $validationStatus = if ($completionStatus -ne 'Complete' -or $notComparable) {
        'Not comparable'
    } elseif ($findingState.ValidationIssues -gt 0) {
        'Issues detected'
    } else {
        'No differences found'
    }
    $outcome = switch ($completionStatus) {
        'Complete' { 'Completed' }
        'Partial' { 'Partial' }
        'TimedOut' { 'TimedOut' }
        default { 'Unavailable' }
    }

    $targetResults = @($snapshots | ForEach-Object { $_.TargetResults })
    $reservationCount = if ($Mode -eq 'DeepAudit') { @($snapshots | ForEach-Object { $_.DeepReservations }).Count } else { 0 }
    $remoteCallCount = [int](($snapshots | Measure-Object -Property RemoteCallCount -Sum).Sum)
    $callCountIsLowerBound = @($snapshots | Where-Object { $_.CompletionStatus -eq 'TimedOut' -or $null -eq $_.RemoteCallCount }).Count -gt 0

    return [PSCustomObject]@{
        Requested = $true
        Mode = $Mode
        Outcome = $outcome
        CompletionStatus = $completionStatus
        IsComplete = ($completionStatus -eq 'Complete')
        ValidationStatus = $validationStatus
        MastersDiscovered = $masters.Count
        MastersComplete = $completeSnapshots.Count
        MasterResults = $snapshots
        TargetResults = $targetResults
        TargetCount = @($targetSet.Targets).Count
        ResolvedTargetCount = [int]$targetSet.ResolvedCount
        UnresolvedTargetCount = [int]$targetSet.UnresolvedCount
        ReservationCount = $reservationCount
        Findings = @($findings.ToArray())
        FindingCount = [int]$findingState.Count
        FindingsOmitted = [Math]::Max(0, [int]$findingState.Count - $findings.Count)
        RemoteCallCount = $remoteCallCount
        CallCountIsLowerBound = $callCountIsLowerBound
        DurationMs = [int]$overallBudgetTimer.ElapsedMilliseconds
        ErrorMessage = ''
    }
}

function Format-PvsDhcpResultText {
    param([object]$DhcpResult)

    if ($null -eq $DhcpResult -or -not [bool]$DhcpResult.Requested) { return '' }
    $modeLabel = switch ([string]$DhcpResult.Mode) {
        'Baseline' { 'Whole Farm Baseline' }
        'Targeted' {
            switch ([string]$DhcpResult.RetrievalScope) {
                'SingleServer' { 'Single Server Targeted' }
                'DeviceCollection' { 'Device Collection Targeted' }
                default { 'Server List Targeted' }
            }
        }
        'DeepAudit' { 'Whole Farm Deep Audit' }
        default { [string]$DhcpResult.Mode }
    }

    $lines = [System.Collections.Generic.List[string]]::new()
    $optionMap = Get-DhcpValidationOptionMap
    [void]$lines.Add("Mode: $modeLabel")
    [void]$lines.Add("Status: $($DhcpResult.CompletionStatus)")
    [void]$lines.Add("Configuration snapshots: $($DhcpResult.MastersComplete) of $($DhcpResult.MastersDiscovered) complete")
    [void]$lines.Add("Validation: $($DhcpResult.ValidationStatus)")
    [void]$lines.Add('Option evaluation: Default standard DHCP options only; policy, vendor-class, and user-class effective values are outside this validation.')
    $failedBatchMasters = @($DhcpResult.MasterResults | Where-Object { [bool]$_.TargetBatchQueryFailed })
    if ($failedBatchMasters.Count -gt 0) {
        [void]$lines.Add("Target lookup batch failure: $($failedBatchMasters.Count) Master(s); affected targets are incomplete and no scope-wide fallback was attempted.")
    }
    $callCountLabel = if ([bool]$DhcpResult.CallCountIsLowerBound) { "$($DhcpResult.RemoteCallCount) or more (timed-out work is not fully observable)" } else { [string]$DhcpResult.RemoteCallCount }
    [void]$lines.Add("Recorded remote management calls: $callCountLabel")
    [void]$lines.Add("DHCP execution: $(Format-ExecutionTime -Elapsed ([TimeSpan]::FromMilliseconds([double]$DhcpResult.DurationMs)))")

    if ([string]$DhcpResult.Outcome -eq 'Declined') {
        [void]$lines.Add('The large targeted DHCP validation was declined and was not run.')
        return ($lines -join [Environment]::NewLine)
    }

    if ([string]$DhcpResult.Mode -eq 'Baseline') {
        [void]$lines.Add('Reservation audit: Not run. Reservation-level overrides were not evaluated.')
    } elseif ([string]$DhcpResult.Mode -eq 'Targeted') {
        [void]$lines.Add("Targets: $($DhcpResult.TargetCount) matched, $($DhcpResult.ResolvedTargetCount) resolved, $($DhcpResult.UnresolvedTargetCount) unresolved")
        [void]$lines.Add('Unrelated reservations: Never scanned by Targeted mode, including when a target batch query fails.')
    } elseif ([string]$DhcpResult.Mode -eq 'DeepAudit') {
        [void]$lines.Add("Reservation-Master records audited: $($DhcpResult.ReservationCount)")
    }

    [void]$lines.Add('')
    [void]$lines.Add('Master Results:')
    foreach ($master in @($DhcpResult.MasterResults | Sort-Object Ordinal)) {
        $masterCallText = if ($null -eq $master.RemoteCallCount) { 'Unknown' } else { [string]$master.RemoteCallCount }
        $batchFailureText = if ([bool]$master.TargetBatchQueryFailed) { ', TargetBatchQueryFailed=Yes' } else { '' }
        [void]$lines.Add(" - $($master.MasterName): ConfigStatus=$($master.CompletionStatus), Service=$($master.ServiceState), Scopes=$(@($master.Scopes).Count), Calls=$masterCallText, DurationMs=$($master.DurationMs)$batchFailureText")
        $masterErrors = @($master.Errors)
        foreach ($errorText in @($masterErrors | Select-Object -First 10)) {
            [void]$lines.Add("   - $errorText")
        }
        if ($masterErrors.Count -gt 10) {
            [void]$lines.Add("   - $($masterErrors.Count - 10) additional Master errors were omitted from this view.")
        }
    }

    if ([string]$DhcpResult.Mode -eq 'Targeted' -and [int]$DhcpResult.TargetCount -eq 1) {
        [void]$lines.Add('')
        [void]$lines.Add('Target Details:')
        foreach ($targetResult in @($DhcpResult.TargetResults | Sort-Object MasterName)) {
            $scopeText = if ($targetResult.ScopeId) { [string]$targetResult.ScopeId } else { 'N/A' }
            $reservationText = if ($targetResult.ReservationIPAddress) { "$($targetResult.ReservationIPAddress) [$($targetResult.ReservationName)]" } else { 'Not Found' }
            $clientIdText = if ($targetResult.ReservationClientId) { [string]$targetResult.ReservationClientId } else { 'N/A' }
            [void]$lines.Add(" - $($targetResult.MasterName): Status=$($targetResult.Status), Scope=$scopeText, Reservation=$reservationText, ClientId=$clientIdText")
            if ($targetResult.Status -eq 'Found') {
                foreach ($optionId in @(3,6,11,15,66,67)) {
                    $cell = $targetResult.EffectiveOptions[[string]$optionId]
                    $sourceText = if ($cell.Source) { [string]$cell.Source } else { 'None' }
                    [void]$lines.Add(("   - {0:D3} {1}: {2} ({3})" -f $optionId, $optionMap[$optionId], (Get-DhcpOptionCellDisplayText -Cell $cell), $sourceText))
                }
            }
        }
    }

    [void]$lines.Add('')
    [void]$lines.Add("Findings: $($DhcpResult.FindingCount)")
    if ([int]$DhcpResult.FindingCount -eq 0) {
        [void]$lines.Add(' - None')
    } else {
        foreach ($finding in @($DhcpResult.Findings)) { [void]$lines.Add(" - $finding") }
        if ([int]$DhcpResult.FindingsOmitted -gt 0) {
            [void]$lines.Add(" - $($DhcpResult.FindingsOmitted) additional findings were omitted from the GUI summary.")
        }
    }
    return ($lines -join [Environment]::NewLine)
}

# Formats pipe-delimited group detail into panel-friendly multi-line text.
function Format-GroupDetailForPanel {
    param(
        [string]$Detail
    )
    if (-not $Detail) { return '' }
    return ($Detail -replace ': ', ":`r`n - " -replace ' \| ', "`r`n - ")
}

# Formats count summary sections for output text panels.
function Format-CountSection {
    param(
        [string]$Heading,
        [string]$Description,
        [string]$Detail
    )

    $prefix = "${Heading}: "
    $items = $Detail
    if ($items.StartsWith($prefix)) {
        $items = $items.Substring($prefix.Length)
    }
    $items = $items -replace ' \| ', "`r`n - "

    return @(
        $Description
        " - $items"
    ) -join "`r`n"
}

<#
.SYNOPSIS
    Applies consistent DataGrid formatting for auto-generated columns.
.DESCRIPTION
    Adds friendly headers (PascalCase -> spaced words), centered text,
    and minimum widths so both PVS and DDC tables remain readable.
#>
function Set-DataGridDisplayFormat {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [System.Windows.Window]$Window
    )

    if ($null -eq $DataGrid -or $null -eq $Window) { return }

    $DataGrid.Add_AutoGeneratingColumn({
        param($sender, $e)

        $headerText = [regex]::Replace($e.PropertyName, '(?<!^)([A-Z])', ' $1')
        $headerText = $headerText -replace '\bI P\b', 'IP'
        $headerText = $headerText -replace '\bC S R\b', 'CSR'
        $headerText = $headerText -replace '\bX D C\b', 'XDC'
        $headerText = $headerText -replace '\bv Disk\b', 'vDisk'
        $e.Column.Header = $headerText
        $minWidth = 120
        $columnWidthType = [System.Windows.Controls.DataGridLengthUnitType]::SizeToCells

        switch ($e.PropertyName.ToLowerInvariant()) {
            'vdisk' {
                # vDisk names are typically long; make this column wider by default.
                $minWidth = 320
                $columnWidthType = [System.Windows.Controls.DataGridLengthUnitType]::SizeToHeader
            }
            'maintenancemode' {
                # Keep full header text visible without truncation.
                $minWidth = 190
                $columnWidthType = [System.Windows.Controls.DataGridLengthUnitType]::SizeToHeader
            }
        }

        $e.Column.MinWidth = $minWidth
        $e.Column.Width = New-Object System.Windows.Controls.DataGridLength(1, $columnWidthType)

        if ($e.Column -is [System.Windows.Controls.DataGridTextColumn]) {
            $e.Column.ElementStyle = $Window.FindResource('CenteredDataGridText')
            $e.Column.EditingElementStyle = $Window.FindResource('CenteredDataGridEditingText')
        }
    })
}

# -------------------------------------------------------------
# Region: PVS Mode
# -------------------------------------------------------------
# Builds the PVS GUI and wires all PVS-specific retrieval events. The UI is
# defined inline so the script remains a single deployable .ps1 file.
<#
.SYNOPSIS
    Displays the PVS Server Info window and owns its event-driven workflow.
.DESCRIPTION
    Builds the XAML, binds named controls, maintains independent scope inputs,
    retrieves PVS devices, shapes selected fields, and optionally runs all-Master
    DHCP validation. Output data and DHCP details intentionally use separate tabs.
#>
function Show-PVSWindow {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName 'System.Windows.Forms'

    # --- Phase 1: Define the complete PVS window.
    # SharedStyles is interpolated into this XAML before WPF parses it.
    $XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Citrix PVS - Server Info" Height="640" Width="940"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#ECF0F1" MinWidth="760" MinHeight="500">
  $($script:SharedStyles)
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <!-- Window header: product name, purpose, and the local execution host. -->
    <Border Grid.Row="0" Background="{StaticResource HeaderBgBrush}" Padding="20,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0">
          <TextBlock Text="PVS Server Info" FontSize="18" FontWeight="Bold"
                     Foreground="{StaticResource HeaderFgBrush}"/>
          <TextBlock Text="Retrieve server information from Citrix PVS Master" FontSize="11"
                     Foreground="#AEB6BF" Margin="0,3,0,0"/>
        </StackPanel>
        <TextBlock Grid.Column="1" Name="lblServer" FontSize="11" Foreground="#AEB6BF"
                   VerticalAlignment="Center"/>
      </Grid>
    </Border>

    <TabControl Name="tabMain" Grid.Row="1" Margin="12,10,12,10">
      <!-- Execution owns inputs and commands; it does not display retrieved rows. -->
      <TabItem Header="Execution">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="{StaticResource CardBgBrush}" CornerRadius="6"
                  BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1"
                  Margin="4,6,4,0" Padding="16,10">
            <StackPanel>
              <StackPanel>
                <TextBlock Text="1. Retrieval Scope" FontSize="13" FontWeight="SemiBold"
                           Foreground="#2C3E50"/>
                <WrapPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="0,8,0,0">
                  <RadioButton Name="rbWholeFarm" Content="Whole Farm" GroupName="PvsScope"
                               IsChecked="True" VerticalAlignment="Center" Margin="0,0,18,0"/>
                  <RadioButton Name="rbDeviceCollection" Content="Device Collection" GroupName="PvsScope"
                               VerticalAlignment="Center" Margin="0,0,18,0"/>
                  <RadioButton Name="rbServerList" Content="Server List" GroupName="PvsScope"
                               VerticalAlignment="Center" Margin="0,0,18,0"/>
                  <RadioButton Name="rbSingleServer" Content="Single Server" GroupName="PvsScope"
                               VerticalAlignment="Center"/>
                </WrapPanel>
                <StackPanel Name="pnlCollectionSelection" Visibility="Collapsed" Margin="0,8,0,0">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Device Collection:" FontSize="12"
                               VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <ComboBox Grid.Column="1" Name="cmbDeviceCollections" Height="30"
                              DisplayMemberPath="DisplayName" SelectedValuePath="ChoiceKey"
                              IsEditable="False" IsTextSearchEnabled="True" IsEnabled="False"
                              MaxDropDownHeight="320" VerticalContentAlignment="Center"/>
                    <Button Grid.Column="2" Name="btnRefreshCollections" Content="Refresh"
                            Style="{StaticResource SecondaryButton}" Width="82" Height="30"
                            Margin="8,0,0,0" IsEnabled="False"
                            ToolTip="Reload the Device Collection choices from PVS"/>
                  </Grid>
                  <TextBlock Name="txtCollectionLoadStatus" Text="Loading Device Collections..."
                             FontSize="10" Foreground="#7F8C8D" Margin="0,4,0,0"
                             TextWrapping="Wrap"/>
                </StackPanel>
                <DockPanel Name="pnlServerInput" Margin="0,8,0,0" LastChildFill="True" Visibility="Collapsed">
                  <TextBlock Name="lblServerNamePrompt" Text="Server Name:" FontSize="12"
                             VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <TextBox Name="txtServerNames" Height="30" Style="{StaticResource StyledTextBox}"
                           IsEnabled="False" AcceptsReturn="True" TextWrapping="NoWrap" MaxLength="262144"
                           VerticalScrollBarVisibility="Auto" VerticalAlignment="Center"/>
                </DockPanel>
                <TextBlock Name="txtScopeHint" Text="Whole farm selected; no server names are required."
                           FontSize="10" Foreground="#7F8C8D" Margin="0,4,0,0"/>
              </StackPanel>
              <DockPanel Margin="0,16,0,8">
                <TextBlock Text="2. Server Information" FontSize="13" FontWeight="SemiBold"
                           Foreground="#2C3E50" VerticalAlignment="Center"/>
                <Button Name="btnToggle" Content="Select All Information" Style="{StaticResource LinkButton}"
                        HorizontalAlignment="Right" VerticalAlignment="Center"/>
              </DockPanel>
              <WrapPanel Orientation="Horizontal">
                <CheckBox Name="chkDevices" Content="Device Name"  IsChecked="True"  Margin="0,0,18,6"/>
                <CheckBox Name="chkVDisk"   Content="vDisk"        IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkPingReachability" Content="Ping Reachability" IsChecked="False" Margin="0,0,18,6"
                          ToolTip="After successful IPv4 name resolution, attempts one ICMP echo per selected device. Reachable confirms only a ping reply; it does not confirm Citrix registration, Windows health, or actual power state. No ICMP Reply can also mean that a firewall blocks ping."/>
                <CheckBox Name="chkServerIP" Content="Server IP"   IsChecked="False" Margin="0,0,18,6"/>
              </WrapPanel>
              <TextBlock Text="Device Collection is always included in the output."
                         FontSize="10" Foreground="#7F8C8D" Margin="0,0,0,8"/>

              <TextBlock Text="Personality Information" FontSize="12" FontWeight="SemiBold"
                         Foreground="#34495E"/>
              <WrapPanel Orientation="Horizontal" Margin="0,6,0,0">
                <CheckBox Name="chkReboot" Content="Reboot Day" IsChecked="False" Margin="0,0,18,6"
                          ToolTip="Read from the Reboot device-personality value"/>
                <CheckBox Name="chkCSR" Content="CSR Server" IsChecked="False" Margin="0,0,18,6"
                          ToolTip="Read from the CSAServer device-personality value"/>
                <StackPanel Orientation="Horizontal" Margin="0,0,18,6">
                  <CheckBox Name="chkXDC" Content="XDC Server" IsChecked="False" Margin="0,0,5,0"/>
                  <Border Width="16" Height="16" CornerRadius="8" Background="#D4E6F1"
                          ToolTip="XDC Server is relevant only for DR nodes. It is read from the XDC_LIST device-personality value.">
                    <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#1A5276"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                  </Border>
                </StackPanel>
              </WrapPanel>
              <TextBlock Text="These values are read from PVS device personality."
                         FontSize="10" Foreground="#7F8C8D"/>

              <Border Margin="0,12,0,0" Padding="10,8" CornerRadius="4"
                      Background="#F8F9F9" BorderBrush="#D5DBDB" BorderThickness="1">
                <StackPanel>
                  <TextBlock Text="3. DHCP Validation (Optional)" FontSize="12" FontWeight="SemiBold"
                             Foreground="#34495E"/>
                  <StackPanel Orientation="Horizontal" Margin="0,7,0,0">
                    <CheckBox Name="chkDhcpCheck" Content="Run DHCP Validation across all PVS Masters"
                              IsChecked="False" Margin="0,0,5,0"/>
                    <Border Width="16" Height="16" CornerRadius="8" Background="#D4E6F1">
                      <Border.ToolTip>
                        <ToolTip>
                          <TextBlock Width="390" TextWrapping="Wrap"
                                     Text="The work follows the selected retrieval scope. Whole Farm performs a lightweight baseline of DHCP service, exact scopes, and default standard server/scope options on every PVS Master without reading reservations. Device Collection validates only devices in the selected collection, Server List validates only matched servers, and Single Server validates only that server, across every Master. DHCP policies, vendor classes, and user classes are outside this validation. Failed reads are shown as incomplete and are never treated as matching values. Windows DHCP PowerShell cmdlets are required. Deadlines and Cancel are cooperative; if a vendor RPC does not return, closing the window a second time offers an explicit Force Exit."/>
                        </ToolTip>
                      </Border.ToolTip>
                      <TextBlock Text="i" FontSize="10" FontWeight="Bold" Foreground="#1A5276"
                                 HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                  </StackPanel>
                  <CheckBox Name="chkDhcpDeepAudit" Content="Deep Full Farm Audit (all reservations)"
                            IsChecked="False" IsEnabled="False" Visibility="Collapsed" Margin="20,7,0,0"
                            ToolTip="Reads every DHCP reservation and its reservation-level options on every PVS Master. This can be very slow and must be run off-peak."/>
                  <TextBlock Name="txtDhcpModeHint" Text="Enable DHCP Validation to check the Whole Farm baseline."
                             FontSize="10" Foreground="#7F8C8D" Margin="0,5,0,0" TextWrapping="Wrap"/>
                  <TextBlock Text="DHCP is not selected by Select All Information. Only one DHCP validation should run per PVS farm."
                             FontSize="10" Foreground="#7F8C8D" Margin="0,3,0,0" TextWrapping="Wrap"/>
                </StackPanel>
              </Border>
            </StackPanel>
          </Border>

          <StackPanel Grid.Row="1" HorizontalAlignment="Center" Margin="4,12,4,0">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
              <Button Name="btnRetrieve" Content="Retrieve Data" Style="{StaticResource PrimaryButton}" Width="140" Height="34" Margin="0,0,10,0"/>
              <Button Name="btnCancel" Content="Cancel" Style="{StaticResource SecondaryButton}" Width="90" Height="34" Margin="0,0,10,0"
                      IsEnabled="False" ToolTip="Request cancellation after the current SDK call returns"/>
              <Button Name="btnExport" Content="Export to CSV" Style="{StaticResource SecondaryButton}" Width="140" Height="34" Margin="0,0,10,0"/>
              <Button Name="btnViewLog" Content="View Log" Style="{StaticResource SecondaryButton}" Width="110" Height="34" Margin="0,0,10,0"
                      ToolTip="Open today's execution log in Notepad"/>
              <Button Name="btnClose" Content="Close" Style="{StaticResource CloseButton}" Width="100" Height="34"/>
            </StackPanel>
            <TextBlock Name="lblLogPath" Text="" Width="650" TextAlignment="Center" TextTrimming="CharacterEllipsis"
                       FontSize="10" Foreground="#7F8C8D" Margin="0,6,0,0"/>
            <TextBlock Text="Run only one copy per PVS farm. Same-server duplicate launches are blocked automatically."
                       Width="700" TextAlignment="Center" FontSize="10" Foreground="#7F8C8D" Margin="0,3,0,0"/>
          </StackPanel>

          <Border Grid.Row="2" Background="{StaticResource CardBgBrush}" CornerRadius="6"
                  BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1"
                  Margin="4,12,4,0" Padding="10,6">
            <StackPanel>
              <DockPanel>
                <TextBlock Text="Filter:" FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="#7F8C8D"/>
                <TextBox Name="txtSearch" Style="{StaticResource StyledTextBox}" BorderThickness="0"
                         FontSize="12" VerticalAlignment="Center"/>
              </DockPanel>
              <TextBlock Text="Filters retrieved results only. Enter text to search all columns, or use Field:Value to search one column (examples: ServerName:PVS01, PingReachability:Reachable). Clear the box to show all rows."
                         FontSize="10" Foreground="#7F8C8D" Margin="0,4,0,0" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>

          <StackPanel Grid.Row="3" Margin="4,8,4,0">
            <TextBlock Name="lblProgressDetail" Text="" FontSize="11" Foreground="#2C3E50"
                       TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,4"/>
            <ProgressBar Name="progressBar" Style="{StaticResource AccentProgress}" Visibility="Hidden"/>
          </StackPanel>
          </Grid>
        </ScrollViewer>
      </TabItem>

      <!-- Output owns PVS/collection summaries and the filterable server table. -->
      <TabItem Header="Output">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="130" MinHeight="80"/>
            <RowDefinition Height="6"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="{StaticResource CardBgBrush}" CornerRadius="6"
                  BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1"
                  Margin="4,6,4,0" Padding="10,6">
            <TextBox Name="txtCollectionSummary" IsReadOnly="True" BorderThickness="0" Background="Transparent"
                     TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontSize="11"
                     Foreground="#2C3E50"/>
          </Border>

          <GridSplitter Grid.Row="1" Height="6" HorizontalAlignment="Stretch" VerticalAlignment="Center"
                        Background="#D5DBDB" ResizeDirection="Rows" ShowsPreview="True"/>

          <DataGrid Name="dataGrid" Grid.Row="2" Margin="4,10,4,8" AutoGenerateColumns="True"
                    IsReadOnly="True" AlternatingRowBackground="{StaticResource GridAltRowBrush}"
                    GridLinesVisibility="All" CanUserSortColumns="True"
                    Background="White" BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1"
                    HorizontalGridLinesBrush="#E5E8E8" VerticalGridLinesBrush="#E5E8E8" RowHeight="32"
                    ColumnHeaderHeight="34"
                    SelectionMode="Extended" SelectionUnit="FullRow"/>

          <Border Grid.Row="3" Background="{StaticResource StatusBgBrush}" Padding="16,7" Margin="4,0,4,4">
            <DockPanel>
              <Button Name="btnCopyTable" Content="Copy Table" Style="{StaticResource PrimaryButton}"
                      Width="100" Height="26" Padding="8,2" FontSize="10" DockPanel.Dock="Right" Margin="8,-2,0,-2"
                      ToolTip="Copy the complete currently displayed table, including headings"/>
              <Button Name="btnCopyServerList" Content="Copy Server List" Style="{StaticResource PrimaryButton}"
                      Width="125" Height="26" Padding="8,2" FontSize="10" DockPanel.Dock="Right" Margin="16,-2,0,-2"
                      ToolTip="Copy selected server names, or all currently displayed names when no rows are selected"/>
              <TextBlock Name="lblRecordCount" Text="" FontSize="11"
                         Foreground="{StaticResource StatusFgBrush}" DockPanel.Dock="Right"/>
              <TextBlock Name="lblExecutionTime" Text="" FontSize="11" Margin="0,0,18,0"
                         Foreground="{StaticResource StatusFgBrush}" DockPanel.Dock="Right"/>
              <TextBlock Name="statusText" Text="Ready" FontSize="11"
                         Foreground="{StaticResource StatusFgBrush}"/>
            </DockPanel>
          </Border>
        </Grid>
      </TabItem>

      <!-- Hidden until requested. This tab is the exclusive owner of DHCP output. -->
      <TabItem Name="tabDhcpDetails" Header="DHCP Details" Visibility="Collapsed">
        <Grid Margin="4,6,4,4">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <Border Grid.Row="0" Background="{StaticResource CardBgBrush}" CornerRadius="6"
                  BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1" Padding="12,8">
            <StackPanel>
              <TextBlock Text="DHCP Validation Details" FontSize="13" FontWeight="SemiBold"
                         Foreground="#2C3E50"/>
              <TextBlock Text="This tab contains all DHCP validation results and opens automatically when DHCP Validation is selected."
                         FontSize="10" Foreground="#7F8C8D" Margin="0,3,0,0"/>
            </StackPanel>
          </Border>
          <TextBox Name="txtDhcpFullDetails" Grid.Row="1" Margin="0,8,0,0" IsReadOnly="True"
                   Background="White" BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1"
                   Padding="12,10" FontSize="11" Foreground="#2C3E50" TextWrapping="Wrap"
                   VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"/>
        </Grid>
      </TabItem>
    </TabControl>
  </Grid>
</Window>
"@

    # --- Phase 2: Parse XAML and obtain the named controls used by event handlers.
    $window = [Windows.Markup.XamlReader]::Parse($XAML)

    # Keep these references local to the modal window. Event-handler scriptblocks
    # close over them for the lifetime of Show-PVSWindow.
    $tabMain       = $window.FindName('tabMain')
    $chkDevices    = $window.FindName('chkDevices')
    $chkVDisk      = $window.FindName('chkVDisk')
    $chkReboot     = $window.FindName('chkReboot')
    $chkCSR        = $window.FindName('chkCSR')
    $chkXDC        = $window.FindName('chkXDC')
    $chkPingReachability = $window.FindName('chkPingReachability')
    $chkServerIP   = $window.FindName('chkServerIP')
    $rbWholeFarm   = $window.FindName('rbWholeFarm')
    $rbDeviceCollection = $window.FindName('rbDeviceCollection')
    $rbSingleServer = $window.FindName('rbSingleServer')
    $rbServerList  = $window.FindName('rbServerList')
    $pnlCollectionSelection = $window.FindName('pnlCollectionSelection')
    $cmbDeviceCollections = $window.FindName('cmbDeviceCollections')
    $btnRefreshCollections = $window.FindName('btnRefreshCollections')
    $txtCollectionLoadStatus = $window.FindName('txtCollectionLoadStatus')
    $chkDhcpCheck  = $window.FindName('chkDhcpCheck')
    $chkDhcpDeepAudit = $window.FindName('chkDhcpDeepAudit')
    $txtDhcpModeHint = $window.FindName('txtDhcpModeHint')
    $pnlServerInput = $window.FindName('pnlServerInput')
    $lblServerNamePrompt = $window.FindName('lblServerNamePrompt')
    $txtServerNames = $window.FindName('txtServerNames')
    $txtScopeHint  = $window.FindName('txtScopeHint')
    $btnToggle     = $window.FindName('btnToggle')
    $btnRetrieve   = $window.FindName('btnRetrieve')
    $btnCancel     = $window.FindName('btnCancel')
    $btnExport     = $window.FindName('btnExport')
    $btnViewLog    = $window.FindName('btnViewLog')
    $btnClose      = $window.FindName('btnClose')
    $txtSearch     = $window.FindName('txtSearch')
    $dataGrid      = $window.FindName('dataGrid')
    $btnCopyTable  = $window.FindName('btnCopyTable')
    $btnCopyServerList = $window.FindName('btnCopyServerList')
    $progressBar   = $window.FindName('progressBar')
    $lblProgressDetail = $window.FindName('lblProgressDetail')
    $txtCollectionSummary = $window.FindName('txtCollectionSummary')
    $tabDhcpDetails = $window.FindName('tabDhcpDetails')
    $txtDhcpFullDetails = $window.FindName('txtDhcpFullDetails')
    $statusText    = $window.FindName('statusText')
    $lblRecordCount = $window.FindName('lblRecordCount')
    $lblExecutionTime = $window.FindName('lblExecutionTime')
    $lblLogPath    = $window.FindName('lblLogPath')
    $lblServer     = $window.FindName('lblServer')

    Set-DataGridDisplayFormat -DataGrid $dataGrid -Window $window

    $lblServer.Text = "$($env:COMPUTERNAME)"
    $lblLogPath.Text = "Daily log: $($script:LogPath)"
    $lblLogPath.ToolTip = "Daily log, appended by date. Contains retrieval scope and fields, counts, warnings, errors, export paths, and execution time.`n$($script:LogPath)"

    # --- Phase 3: Initialize per-window state.
    # This remains the unfiltered source. Searching only swaps DataGrid.ItemsSource;
    # clearing the filter restores all retrieved rows from this dataset.
    $script:pvsFullData = $null

    # --- Phase 4: Register selection, scope, filter, and retrieval handlers.
    $btnToggle.Add_Click({
        # DHCP validation is intentionally excluded because it can be expensive in large farms.
        $allChecks = @($chkDevices, $chkVDisk, $chkReboot, $chkCSR, $chkXDC, $chkPingReachability, $chkServerIP)
        $anyUnchecked = $allChecks | Where-Object { -not $_.IsChecked }
        $newState = [bool]$anyUnchecked
        foreach ($cb in $allChecks) { $cb.IsChecked = $newState }
        $btnToggle.Content = if ($newState) { 'Deselect All Information' } else { 'Select All Information' }
    })

    # One visible TextBox is reused for both named scopes. These are deliberately
    # separate buffers so a pasted Server List never becomes Single Server input.
    $pvsScopeInputState = @{
        CurrentScope = 'WholeFarm'
        SingleServer = ''
        ServerList = ''
    }

    # Device Collection uses a two-phase flow. Selecting its radio button starts
    # one background preview read and fills the inline ComboBox. Retrieve receives
    # only the opaque ChoiceKey, revalidates it against one fresh farm inventory,
    # and never opens a modal picker.
    $script:pvsRetrievalContext = $null
    $script:pvsCollectionLoadContext = $null
    $pvsCollectionState = @{
        Loaded = $false
        Inventory = @()
        InventorySnapshotId = ''
        Collections = @()
        SelectedChoiceKey = ''
        UnselectableDeviceCount = 0
        LoadedAt = $null
        LastError = ''
        Binding = $false
    }

    $isPvsCollectionSnapshotFresh = {
        return [bool]$pvsCollectionState.Loaded -and $null -ne $pvsCollectionState.LoadedAt -and
            ((Get-Date) - $pvsCollectionState.LoadedAt).TotalMinutes -le $script:NamedScopeSnapshotMaximumAgeMinutes
    }

    $updatePvsRetrieveEnabled = {
        $mainBusy = $null -ne $script:pvsRetrievalContext
        $listBusy = $null -ne $script:pvsCollectionLoadContext
        $collectionReady = (& $isPvsCollectionSnapshotFresh) -and $null -ne $cmbDeviceCollections.SelectedItem
        $btnRetrieve.IsEnabled = -not $mainBusy -and -not $listBusy -and
            (-not [bool]$rbDeviceCollection.IsChecked -or $collectionReady)
    }

    $bindPvsCollectionSelector = {
        if (-not [bool]$rbDeviceCollection.IsChecked) {
            $pnlCollectionSelection.Visibility = 'Collapsed'
            & $updatePvsRetrieveEnabled
            return
        }

        $pnlCollectionSelection.Visibility = 'Visible'
        $loading = $null -ne $script:pvsCollectionLoadContext
        $pvsCollectionState.Binding = $true
        try {
            $cmbDeviceCollections.ItemsSource = $null
            if ([bool]$pvsCollectionState.Loaded) {
                $cmbDeviceCollections.ItemsSource = @($pvsCollectionState.Collections)
                $restoreChoice = @($pvsCollectionState.Collections | Where-Object {
                    $_.ChoiceKey -ieq [string]$pvsCollectionState.SelectedChoiceKey
                } | Select-Object -First 1)
                if ($restoreChoice.Count -eq 1) {
                    $cmbDeviceCollections.SelectedItem = $restoreChoice[0]
                } else {
                    $cmbDeviceCollections.SelectedIndex = -1
                    $pvsCollectionState.SelectedChoiceKey = ''
                }
            } else {
                $cmbDeviceCollections.SelectedIndex = -1
            }
        } finally {
            $pvsCollectionState.Binding = $false
        }

        $choiceCount = @($pvsCollectionState.Collections).Count
        $snapshotFresh = & $isPvsCollectionSnapshotFresh
        $cmbDeviceCollections.IsEnabled = -not $loading -and $snapshotFresh -and $choiceCount -gt 0 -and $null -eq $script:pvsRetrievalContext
        $btnRefreshCollections.IsEnabled = -not $loading -and $null -eq $script:pvsRetrievalContext
        if ($loading) {
            $txtCollectionLoadStatus.Text = 'Loading Device Collections from the current PVS farm...'
            $txtCollectionLoadStatus.Foreground = '#7F8C8D'
        } elseif ([bool]$pvsCollectionState.Loaded -and -not $snapshotFresh) {
            $txtCollectionLoadStatus.Text = "This Device Collection list is older than $($script:NamedScopeSnapshotMaximumAgeMinutes) minutes. Click Refresh before retrieving."
            $txtCollectionLoadStatus.Foreground = '#922B21'
        } elseif ([bool]$pvsCollectionState.Loaded -and $choiceCount -gt 0) {
            $excludedText = if ([int]$pvsCollectionState.UnselectableDeviceCount -gt 0) {
                " $($pvsCollectionState.UnselectableDeviceCount) inventory record(s) without a safe collection identity were excluded."
            } else { '' }
            $loadedText = if ($null -ne $pvsCollectionState.LoadedAt) { "Loaded $($pvsCollectionState.LoadedAt.ToString('HH:mm:ss')). " } else { '' }
            $txtCollectionLoadStatus.Text = "${loadedText}Select one Device Collection. $choiceCount available from $(@($pvsCollectionState.Inventory).Count) devices.$excludedText"
            $txtCollectionLoadStatus.Foreground = '#7F8C8D'
        } elseif ([bool]$pvsCollectionState.Loaded) {
            $txtCollectionLoadStatus.Text = 'No selectable Device Collections were found. Click Refresh to try again.'
            $txtCollectionLoadStatus.Foreground = '#922B21'
        } elseif (-not [string]::IsNullOrWhiteSpace([string]$pvsCollectionState.LastError)) {
            $txtCollectionLoadStatus.Text = "Could not load Device Collections: $($pvsCollectionState.LastError) Click Refresh to retry."
            $txtCollectionLoadStatus.Foreground = '#922B21'
        }
        & $updatePvsRetrieveEnabled
    }

    $finishPvsCollectionLoad = {
        param($Completion)

        $btnCancel.IsEnabled = $false
        $btnCancel.Content = 'Cancel'
        if ([string]$Completion.Outcome -eq 'Completed' -and $null -ne $Completion.Result -and [bool]$Completion.Result.InventoryOnly) {
            $result = $Completion.Result
            $pvsCollectionState.Inventory = @($result.Inventory)
            $pvsCollectionState.InventorySnapshotId = [string]$result.InventorySnapshotId
            $pvsCollectionState.Collections = @($result.Collections)
            $pvsCollectionState.UnselectableDeviceCount = [int]$result.UnselectableDeviceCount
            $pvsCollectionState.LoadedAt = Get-Date
            $pvsCollectionState.SelectedChoiceKey = ''
            $pvsCollectionState.LastError = ''
            $pvsCollectionState.Loaded = $true
            Write-Log "PVS inline Device Collection list loaded | Collections=$(@($result.Collections).Count) | Inventory=$($result.InventoryCount) | Snapshot=$($result.InventorySnapshotId)"
        } elseif ([string]$Completion.Outcome -eq 'Cancelled') {
            $pvsCollectionState.LastError = 'Loading was cancelled.'
            Write-Log 'PVS inline Device Collection load cancelled' -Level WARN
        } else {
            $loadError = [string]$Completion.ErrorMessage
            if ([string]::IsNullOrWhiteSpace($loadError)) { $loadError = 'The background inventory load returned no result.' }
            $pvsCollectionState.LastError = $loadError
            Write-Log "PVS inline Device Collection load failed: $loadError | Detail: $([string]$Completion.ErrorDetail)" -Level ERROR
        }
        & $bindPvsCollectionSelector
        if ([string]$Completion.Outcome -eq 'Completed' -and [bool]$rbDeviceCollection.IsChecked -and @($pvsCollectionState.Collections).Count -gt 0) {
            [void]$cmbDeviceCollections.Focus()
            $cmbDeviceCollections.IsDropDownOpen = $true
        }
    }

    $pvsCollectionLoadTimer = New-Object System.Windows.Threading.DispatcherTimer
    $pvsCollectionLoadTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $pvsCollectionLoadTimer.Add_Tick({
        $context = $script:pvsCollectionLoadContext
        if ($null -eq $context) { return }

        foreach ($message in @(Receive-CitrixRetrievalMessages -Context $context)) {
            if ([string]$message.Kind -eq 'Log') {
                Write-RetrievalWorkerMessageLog -Message $message -Context $context
            } elseif ([string]$message.Kind -eq 'Progress' -and [bool]$rbDeviceCollection.IsChecked -and -not [bool]$context.CancellationRequested) {
                $txtCollectionLoadStatus.Text = [string]$message.Message
            }
        }

        if ($null -ne $context.Handle -and $context.Handle.IsCompleted -and $context.MessageQueue.IsEmpty) {
            $pvsCollectionLoadTimer.Stop()
            $completion = Complete-CitrixRetrievalWorker -Context $context
            if ($null -ne $script:pvsCollectionLoadContext -and $script:pvsCollectionLoadContext.RunId -eq $context.RunId) {
                $script:pvsCollectionLoadContext = $null
                & $finishPvsCollectionLoad $completion
            }
        }
    })

    $startPvsCollectionLoad = {
        if ($null -ne $script:pvsRetrievalContext -or $null -ne $script:pvsCollectionLoadContext) { return }

        $pvsCollectionState.Loaded = $false
        $pvsCollectionState.Inventory = @()
        $pvsCollectionState.InventorySnapshotId = ''
        $pvsCollectionState.Collections = @()
        $pvsCollectionState.SelectedChoiceKey = ''
        $pvsCollectionState.UnselectableDeviceCount = 0
        $pvsCollectionState.LoadedAt = $null
        $pvsCollectionState.LastError = ''
        $cmbDeviceCollections.ItemsSource = $null
        $cmbDeviceCollections.SelectedIndex = -1
        & $bindPvsCollectionSelector

        try {
            $script:pvsCollectionLoadContext = Start-CitrixRetrievalWorker -Role PVS -Request @{ ScopeInventoryOnly = $true }
            $btnCancel.Content = 'Cancel Load'
            $btnCancel.IsEnabled = $true
            Write-Log "PVS inline Device Collection load started | RunId=$($script:pvsCollectionLoadContext.RunId)"
            $pvsCollectionLoadTimer.Start()
        } catch {
            $script:pvsCollectionLoadContext = $null
            & $finishPvsCollectionLoad ([PSCustomObject]@{
                Outcome = 'Failed'
                Result = $null
                ErrorMessage = $_.Exception.Message
                ErrorDetail = Get-ErrorDiagnosticText -ErrorRecord $_ -Phase 'PVS collection-preview worker startup'
            })
        }
        & $bindPvsCollectionSelector
    }

    $cmbDeviceCollections.Add_SelectionChanged({
        if ([bool]$pvsCollectionState.Binding) { return }
        $pvsCollectionState.SelectedChoiceKey = if ($null -ne $cmbDeviceCollections.SelectedItem) {
            [string]$cmbDeviceCollections.SelectedItem.ChoiceKey
        } else { '' }
        & $updatePvsRetrieveEnabled
    })
    $btnRefreshCollections.Add_Click({ & $startPvsCollectionLoad })

    # DHCP mode is derived from the retrieval scope. The expensive Deep Audit is
    # offered only for Whole Farm and always starts unchecked.
    $updatePvsDhcpModeControls = {
        $dhcpEnabled = [bool]$chkDhcpCheck.IsChecked
        $wholeFarm = [bool]$rbWholeFarm.IsChecked
        $busy = $null -ne $script:pvsRetrievalContext

        if ($dhcpEnabled -and $wholeFarm) {
            $chkDhcpDeepAudit.Visibility = 'Visible'
            $chkDhcpDeepAudit.IsEnabled = -not $busy
            if ([bool]$chkDhcpDeepAudit.IsChecked) {
                $txtDhcpModeHint.Text = 'Deep Audit: reads every reservation and reservation-level option on every PVS Master. Run off-peak.'
                $txtDhcpModeHint.Foreground = '#922B21'
            } else {
                $txtDhcpModeHint.Text = 'Whole Farm Baseline: checks Master service, exact scopes, and server/scope options. Reservations are not read.'
                $txtDhcpModeHint.Foreground = '#7F8C8D'
            }
        } elseif ($dhcpEnabled -and [bool]$rbSingleServer.IsChecked) {
            $chkDhcpDeepAudit.IsChecked = $false
            $chkDhcpDeepAudit.IsEnabled = $false
            $chkDhcpDeepAudit.Visibility = 'Collapsed'
            $txtDhcpModeHint.Text = 'Single Server Targeted: point-queries only the matched server across every PVS Master; a failed batch is reported incomplete without scanning unrelated reservations.'
            $txtDhcpModeHint.Foreground = '#7F8C8D'
        } elseif ($dhcpEnabled -and [bool]$rbDeviceCollection.IsChecked) {
            $chkDhcpDeepAudit.IsChecked = $false
            $chkDhcpDeepAudit.IsEnabled = $false
            $chkDhcpDeepAudit.Visibility = 'Collapsed'
            $txtDhcpModeHint.Text = 'Device Collection Targeted: validates only devices in the collection selected above, across every PVS Master.'
            $txtDhcpModeHint.Foreground = '#7F8C8D'
        } elseif ($dhcpEnabled) {
            $chkDhcpDeepAudit.IsChecked = $false
            $chkDhcpDeepAudit.IsEnabled = $false
            $chkDhcpDeepAudit.Visibility = 'Collapsed'
            $txtDhcpModeHint.Text = 'Server List Targeted: point-queries matched list entries across every PVS Master; failed batches are incomplete and never broaden to scope scans.'
            $txtDhcpModeHint.Foreground = '#7F8C8D'
        } else {
            $chkDhcpDeepAudit.IsChecked = $false
            $chkDhcpDeepAudit.IsEnabled = $false
            $chkDhcpDeepAudit.Visibility = 'Collapsed'
            $scopeLabel = if ($wholeFarm) {
                'Whole Farm baseline'
            } elseif ([bool]$rbSingleServer.IsChecked) {
                'Single Server targeted validation'
            } elseif ([bool]$rbDeviceCollection.IsChecked) {
                'Device Collection targeted validation'
            } else {
                'Server List targeted validation'
            }
            $txtDhcpModeHint.Text = "Enable DHCP Validation to run $scopeLabel."
            $txtDhcpModeHint.Foreground = '#7F8C8D'
        }
    }

    $updatePvsScopeControls = {
        # Save the outgoing scope before loading the incoming scope. Whole Farm
        # clears only the visible TextBox and must not overwrite either buffer.
        if ($pvsScopeInputState.CurrentScope -eq 'SingleServer') {
            $pvsScopeInputState.SingleServer = [string]$txtServerNames.Text
        } elseif ($pvsScopeInputState.CurrentScope -eq 'ServerList') {
            $pvsScopeInputState.ServerList = [string]$txtServerNames.Text
        }

        $newScope = if ($rbWholeFarm.IsChecked) {
            'WholeFarm'
        } elseif ($rbDeviceCollection.IsChecked) {
            'DeviceCollection'
        } elseif ($rbSingleServer.IsChecked) {
            'SingleServer'
        } else {
            'ServerList'
        }

        $selectedServerScope = [bool]$rbSingleServer.IsChecked -or [bool]$rbServerList.IsChecked
        $txtServerNames.IsEnabled = $selectedServerScope

        if ($newScope -eq 'WholeFarm') {
            $pnlServerInput.Visibility = 'Collapsed'
            $txtServerNames.Text = ''
            $txtServerNames.Height = 30
            $txtScopeHint.Text = 'Whole farm selected; no server names are required.'
        } elseif ($newScope -eq 'DeviceCollection') {
            $pnlServerInput.Visibility = 'Collapsed'
            $txtServerNames.Text = ''
            $txtServerNames.Height = 30
            $txtScopeHint.Text = 'Select a Device Collection below. The current farm list loads automatically.'
        } elseif ($newScope -eq 'SingleServer') {
            $pnlServerInput.Visibility = 'Visible'
            $lblServerNamePrompt.Text = 'Server Name:'
            $txtServerNames.Height = 30
            $txtServerNames.Text = [string]$pvsScopeInputState.SingleServer
            $txtScopeHint.Text = 'Enter one short name, FQDN, or DOMAIN\SERVER value.'
            $txtServerNames.Focus()
        } else {
            $pnlServerInput.Visibility = 'Visible'
            $lblServerNamePrompt.Text = 'Server Names:'
            $txtServerNames.Height = 62
            $txtServerNames.Text = [string]$pvsScopeInputState.ServerList
            $txtScopeHint.Text = 'Paste server names separated by new lines, commas, or semicolons.'
            $txtServerNames.Focus()
        }

        $pvsScopeInputState.CurrentScope = $newScope
        & $bindPvsCollectionSelector
        if ($newScope -eq 'DeviceCollection' -and -not (& $isPvsCollectionSnapshotFresh) -and $null -eq $script:pvsCollectionLoadContext) {
            & $startPvsCollectionLoad
        }
        & $updatePvsDhcpModeControls
        & $updatePvsRetrieveEnabled
    }
    $chkDhcpCheck.Add_Checked($updatePvsDhcpModeControls)
    $chkDhcpCheck.Add_Unchecked($updatePvsDhcpModeControls)
    $chkDhcpDeepAudit.Add_Checked($updatePvsDhcpModeControls)
    $chkDhcpDeepAudit.Add_Unchecked($updatePvsDhcpModeControls)
    $rbWholeFarm.Add_Checked($updatePvsScopeControls)
    $rbDeviceCollection.Add_Checked($updatePvsScopeControls)
    $rbSingleServer.Add_Checked($updatePvsScopeControls)
    $rbServerList.Add_Checked($updatePvsScopeControls)
    & $updatePvsScopeControls

    # Filtering is local and non-destructive; it never reruns a Citrix query.
    # An unknown Field:Value field naturally falls back to all-column text search.
    $txtSearch.Add_TextChanged({
        if ($null -eq $script:pvsFullData) { return }
        $filterResult = Apply-GridFilter -Data $script:pvsFullData -Term $txtSearch.Text
        $dataGrid.ItemsSource = $filterResult.Rows
        $lblRecordCount.Text = $filterResult.Summary
    })

    # --- Phase 5: Main PVS retrieval pipeline.
    # Every Retrieve action reads one current inventory in the retrieval runspace.
    # Device Collection revalidates the previewed opaque key before filtering.
    # Per-device enrichment remains in the worker, and only this UI runspace
    # reads or changes WPF controls.
    $script:pvsActiveRequest = $null
    $script:pvsExecutionTimer = $null

    $pvsInputControls = @(
        $rbWholeFarm, $rbDeviceCollection, $rbServerList, $rbSingleServer, $txtServerNames, $btnToggle,
        $cmbDeviceCollections, $btnRefreshCollections,
        $chkDevices, $chkVDisk, $chkReboot, $chkCSR, $chkXDC,
        $chkPingReachability, $chkServerIP, $chkDhcpCheck, $chkDhcpDeepAudit, $txtSearch,
        $btnCopyServerList, $btnCopyTable
    )

    $setPvsBusyState = {
        param([bool]$Busy)

        foreach ($control in $pvsInputControls) {
            if ($null -ne $control) { $control.IsEnabled = -not $Busy }
        }
        if (-not $Busy) {
            $txtServerNames.IsEnabled = [bool]$rbSingleServer.IsChecked -or [bool]$rbServerList.IsChecked
            & $updatePvsDhcpModeControls
            & $bindPvsCollectionSelector
        } else {
            $cmbDeviceCollections.IsEnabled = $false
            $btnRefreshCollections.IsEnabled = $false
        }

        $btnRetrieve.Content = if ($Busy) { 'Retrieving...' } else { 'Retrieve Data' }
        $btnCancel.IsEnabled = $Busy
        $btnCancel.Content = 'Cancel'
        $btnExport.IsEnabled = -not $Busy
        $btnViewLog.IsEnabled = -not $Busy
        $btnClose.IsEnabled = -not $Busy
        $window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { [System.Windows.Input.Cursors]::Arrow }
        if ($Busy) { $btnRetrieve.IsEnabled = $false } else { & $updatePvsRetrieveEnabled }
    }

    $processPvsWorkerMessages = {
        param($Context)

        foreach ($message in @(Receive-CitrixRetrievalMessages -Context $Context)) {
            switch ([string]$message.Kind) {
                'Log' {
                    Write-RetrievalWorkerMessageLog -Message $message -Context $Context
                }
                'Progress' {
                    if ([bool]$Context.CancellationRequested) { continue }
                    $statusText.Text = [string]$message.Message
                    $lblProgressDetail.Text = [string]$message.Message
                    $progressBar.IsIndeterminate = [bool]$message.IsIndeterminate
                    if (-not [bool]$message.IsIndeterminate -and [int]$message.Total -gt 0) {
                        $progressBar.Value = ([double][int]$message.Current / [double][int]$message.Total) * 100
                    }
                }
                'WorkloadConfirmation' {
                    if ([bool]$Context.CancellationRequested) {
                        $Context.ApprovalState.Approved = $false
                        $Context.ApprovalGate.Set()
                        continue
                    }

                    $details = $message.Details
                    $isDhcpTargetConfirmation = [string]$details.ConfirmationType -eq 'DhcpTargeted'
                    if ($isDhcpTargetConfirmation) {
                        $previewLines = @(
                            "This targeted DHCP request contains $($details.TargetCount) matched servers across $($details.MasterCount) PVS Masters."
                            "It can require up to $($details.TargetMasterCells) reservation-level target checks."
                            ''
                            'Only matched target IPs are point-queried. A failed batch is reported incomplete and never broadens into a full-scope reservation scan.'
                            'Run only one DHCP validation per PVS farm, preferably outside peak hours for a large targeted selection.'
                            ''
                            'Continue with DHCP validation? Choosing No keeps the normal PVS output.'
                        )
                        $dialogTitle = 'Confirm Large Targeted DHCP Validation'
                        Write-Log "Targeted DHCP workload confirmation | Targets=$($details.TargetCount) | Masters=$($details.MasterCount) | TargetMasterCells=$($details.TargetMasterCells)"
                    } else {
                        $previewLines = @(
                            "This request matched $($details.DeviceCount) PVS devices."
                            ''
                            'Additional PVS SDK calls through the local PVS management layer:'
                            " - vDisk: $($details.VDiskCalls)"
                            " - Personality: $($details.PersonalityCalls)"
                            " - Total additional PVS SDK calls: $($details.AdditionalPvsSdkCalls)"
                            ''
                            'Other possible network operations:'
                            " - ICMP ping attempts: $($details.PingAttempts)"
                            " - DNS fallbacks: up to $($details.MaximumDnsFallbacks)"
                            ''
                            'The initial shared farm inventory read is already complete.'
                            'DHCP validation is separate and is not included in this estimate.'
                            'Run only one copy of this tool per PVS farm.'
                            ''
                            'Continue?'
                        )
                        $dialogTitle = 'Confirm Large PVS Workload'
                        Write-Log "PVS workload confirmation | Devices=$($details.DeviceCount) | AdditionalSdkCalls=$($details.AdditionalPvsSdkCalls) | PingAttempts=$($details.PingAttempts) | MaxDnsFallbacks=$($details.MaximumDnsFallbacks)"
                    }
                    $previewText = $previewLines -join [Environment]::NewLine

                    $timerWasRunning = $null -ne $script:pvsExecutionTimer -and $script:pvsExecutionTimer.IsRunning
                    if ($timerWasRunning) {
                        $script:pvsExecutionTimer.Stop()
                    }
                    $choice = [System.Windows.MessageBoxResult]::No
                    try {
                        $choice = [System.Windows.MessageBox]::Show(
                            $previewText,
                            $dialogTitle,
                            [System.Windows.MessageBoxButton]::YesNo,
                            [System.Windows.MessageBoxImage]::Warning,
                            [System.Windows.MessageBoxResult]::No)
                    } catch {
                        # Fail closed. Treat a dialog failure as No so the worker
                        # gate is always released and optional work cannot start.
                        $dialogFailureMessage = $_.Exception.Message
                        try { Write-Log "Workload confirmation dialog failed; defaulting to No: $dialogFailureMessage" -Level WARN } catch { }
                    } finally {
                        if ($timerWasRunning -and $null -ne $script:pvsExecutionTimer -and -not $script:pvsExecutionTimer.IsRunning) {
                            $script:pvsExecutionTimer.Start()
                        }
                    }

                    if ($choice -eq [System.Windows.MessageBoxResult]::Yes) {
                        $Context.ApprovalState.Approved = $true
                        $Context.ApprovalGate.Set()
                        Write-Log $(if ($isDhcpTargetConfirmation) { 'Targeted DHCP workload confirmation approved' } else { 'PVS workload confirmation approved' })
                    } elseif ($isDhcpTargetConfirmation) {
                        $Context.ApprovalState.Approved = $false
                        $Context.ApprovalGate.Set()
                        Write-Log 'Targeted DHCP workload confirmation declined; normal PVS output will still be published' -Level WARN
                        $statusText.Text = 'Continuing without the large targeted DHCP validation...'
                        $lblProgressDetail.Text = 'Continuing without the large targeted DHCP validation...'
                    } else {
                        $Context.ApprovalState.Approved = $false
                        $Context.ApprovalGate.Set()
                        Write-Log 'PVS workload confirmation declined; cancelling before optional processing' -Level WARN
                        $statusText.Text = 'Cancelling before optional PVS processing...'
                        $lblProgressDetail.Text = 'Cancelling before optional PVS processing...'
                        $btnCancel.IsEnabled = $false
                        $btnCancel.Content = 'Cancelling...'
                        Request-CitrixRetrievalCancellation -Context $Context
                    }
                }
            }
        }
    }

    $finishPvsRetrieval = {
        param($Completion)

        $btnCancel.IsEnabled = $false
        $btnCancel.Content = 'Cancel'
        $request = $script:pvsActiveRequest
        $publicationSnapshot = $null
        $publicationCommitted = $false
        try {
            if ([string]$Completion.Outcome -eq 'Cancelled') {
                $hasPreviousOutput = @($script:pvsFullData).Count -gt 0
                $statusText.Text = if ($hasPreviousOutput) { 'Retrieval cancelled; previous successful output retained' } else { 'Retrieval cancelled; no partial table was published' }
                $tabMain.SelectedIndex = if ($hasPreviousOutput) { 1 } else { 0 }
                Write-Log "PVS retrieval cancelled by user | PreviousOutputRetained=$hasPreviousOutput" -Level WARN
                return
            }

            if ([string]$Completion.Outcome -ne 'Completed' -or $null -eq $Completion.Result) {
                $errorText = [string]$Completion.ErrorMessage
                if ([string]::IsNullOrWhiteSpace($errorText)) { $errorText = 'The background retrieval failed without an error message.' }
                [System.Windows.MessageBox]::Show("Error: $errorText", 'Error', 'OK', 'Error') | Out-Null
                $errorDetail = [string]$Completion.ErrorDetail
                Write-Log "PVS background retrieval error: $errorText | Detail: $errorDetail" -Level ERROR
                $hasPreviousOutput = @($script:pvsFullData).Count -gt 0
                $statusText.Text = if ($hasPreviousOutput) { 'Retrieval failed; previous successful output retained' } else { 'Error during retrieval' }
                $tabMain.SelectedIndex = if ($hasPreviousOutput) { 1 } else { 0 }
                return
            }

            $result = $Completion.Result
            $devices = @($result.SourceItems)
            $deviceInfo = @($result.Data)
            $missingServers = @($result.MissingNames)
            $ambiguousServers = @($result.AmbiguousNames)
            $duplicateServers = @($result.DuplicateNames)
            $dhcpResult = $result.DhcpResult
            $totalDevices = $deviceInfo.Count
            $requestedServers = @($request.RequestedServers)
            $singleServerMode = [bool]$request.SingleServerMode
            $serverListMode = [bool]$request.ServerListMode
            $deviceCollectionMode = [bool]$request.DeviceCollectionMode
            $selectedServerMode = [bool]$request.SelectedServerMode
            $selectedCollectionChoice = $result.SelectedCollectionChoice
            $selectedCollectionLocation = if ($deviceCollectionMode -and $null -ne $selectedCollectionChoice) {
                [string]$selectedCollectionChoice.LocationName
            } else { '' }
            $targetServer = [string]$request.TargetServer
            if ($singleServerMode -and $devices.Count -eq 1) {
                $targetServer = [string]$devices[0].Name
            }

            Write-Log "PVS shared inventory returned $($result.AllInventoryCount) devices; selected $totalDevices"
            if ($missingServers.Count -gt 0) {
                Write-Log "PVS requested servers not found ($($missingServers.Count)): $(Format-ServerNamePreview -Names $missingServers)" -Level WARN
            }
            if ($ambiguousServers.Count -gt 0) {
                Write-Log "PVS requested servers are ambiguous ($($ambiguousServers.Count)): $(Format-ServerNamePreview -Names $ambiguousServers)" -Level WARN
            }
            if ($duplicateServers.Count -gt 0) {
                Write-Log "PVS duplicate server aliases ignored ($($duplicateServers.Count)): $(Format-ServerNamePreview -Names $duplicateServers)" -Level WARN
            }

            $selectionIssueText = Format-ServerSelectionIssues -MissingNames $missingServers -AmbiguousNames $ambiguousServers -DuplicateNames $duplicateServers
            if (-not [string]::IsNullOrWhiteSpace($selectionIssueText) -and $totalDevices -gt 0) {
                [System.Windows.MessageBox]::Show(
                    "Matched $totalDevices of $($requestedServers.Count) requested entries. Retrieval completed for the unambiguous matches.$([Environment]::NewLine)$([Environment]::NewLine)$selectionIssueText",
                    'Server List Needs Attention', 'OK', 'Warning') | Out-Null
            }

            if ($totalDevices -eq 0) {
                if ($selectedServerMode) {
                    $requestedPreview = Format-ServerNamePreview -Names $requestedServers
                    $selectionDetail = if ([string]::IsNullOrWhiteSpace($selectionIssueText)) { "Requested: $requestedPreview" } else { $selectionIssueText }
                    [System.Windows.MessageBox]::Show(
                        "No unambiguous requested server matches were available in the PVS device list.$([Environment]::NewLine)$([Environment]::NewLine)$selectionDetail",
                        'Information', 'OK', 'Information') | Out-Null
                } elseif ($deviceCollectionMode) {
                    [System.Windows.MessageBox]::Show(
                        'The selected PVS Device Collection contained no devices in the retrieved farm inventory.',
                        'Information', 'OK', 'Information') | Out-Null
                } else {
                    [System.Windows.MessageBox]::Show('No PVS devices found.', 'Information', 'OK', 'Information') | Out-Null
                }
                # Prepare every derived value before replacing the last-good UI.
                # If a formatter or binding later fails, the catch block restores
                # this snapshot rather than leaving mixed old/new presentation.
                $emptyDhcpRequested = $null -ne $dhcpResult -and [bool]$dhcpResult.Requested
                $emptyDhcpText = if ($emptyDhcpRequested) { Format-PvsDhcpResultText -DhcpResult $dhcpResult } else { '' }
                $emptyStatusText = if ($emptyDhcpRequested) {
                    "No PVS devices found | DHCP $($dhcpResult.CompletionStatus), $($dhcpResult.ValidationStatus)"
                } else { 'No devices found' }
                $publicationSnapshot = [PSCustomObject]@{
                    FullData = $script:pvsFullData
                    ItemsSource = $dataGrid.ItemsSource
                    SearchText = $txtSearch.Text
                    RecordCountText = $lblRecordCount.Text
                    CollectionSummaryText = $txtCollectionSummary.Text
                    DhcpDetailsText = $txtDhcpFullDetails.Text
                    DhcpVisibility = $tabDhcpDetails.Visibility
                    SelectedTab = $tabMain.SelectedItem
                    StatusText = $statusText.Text
                }
                $txtSearch.Text = ''
                $script:pvsFullData = @()
                $dataGrid.ItemsSource = @()
                $lblRecordCount.Text = '0 records'
                $txtCollectionSummary.Text = ''
                $txtDhcpFullDetails.Text = $emptyDhcpText
                $tabDhcpDetails.Visibility = if ($emptyDhcpRequested) { 'Visible' } else { 'Collapsed' }
                $statusText.Text = $emptyStatusText
                if ($emptyDhcpRequested) {
                    $tabMain.SelectedItem = $tabDhcpDetails
                    $txtDhcpFullDetails.ScrollToHome()
                    Write-Log "PVS DHCP result with no device rows | Mode=$($dhcpResult.Mode) | Completion=$($dhcpResult.CompletionStatus) | Validation=$($dhcpResult.ValidationStatus)"
                } else {
                    $tabMain.SelectedIndex = 0
                }
                $publicationCommitted = $true
                Write-Log 'No selected PVS devices were available' -Level WARN
                return
            }

            # Prepare summaries and optional DHCP text without touching the
            # current publication. The final block swaps all visible state once.
            $retrievalStatusText = if ($deviceCollectionMode) {
                "Retrieved $totalDevices devices from collection [$selectedCollectionLocation]"
            } elseif ($serverListMode) {
                "Retrieved $totalDevices of $($requestedServers.Count) list entries | Not found: $($missingServers.Count) | Ambiguous: $($ambiguousServers.Count) | Duplicates: $($duplicateServers.Count)"
            } else {
                "Retrieved $totalDevices devices"
            }
            if ([int]$result.WarningCount -gt 0) {
                $retrievalStatusText += " | Optional-field warnings: $($result.WarningCount)"
            }

            $collectionSummaryText = ''
            $groupedLogText = ''
            $collectionCompletionLogText = ''
            if ($singleServerMode) {
                $collectionSummaryText = "Single Server Mode:$([Environment]::NewLine) - Collection aggregation skipped for [$targetServer]"
            } elseif ($deviceCollectionMode) {
                $collectionSummaryText = @(
                    'Device Collection Mode:'
                    "Selected: $selectedCollectionLocation"
                    "Matched: $totalDevices devices"
                    "Farm inventory: $($result.AllInventoryCount) devices"
                ) -join [Environment]::NewLine
                $collectionCompletionLogText = "PVS Device Collection retrieval complete | Site=[$($selectedCollectionChoice.SiteName)] | Collection=[$($selectedCollectionChoice.CollectionName)] | CollectionId=[$($selectedCollectionChoice.CollectionId)] | Matched=$totalDevices | FarmInventory=$($result.AllInventoryCount)"
            } else {
                $pvsCollectionSummary = Get-GroupedCountSummary -Items $devices -PropertyName 'CollectionName' -Label 'Collections'
                $collectionText = Format-CountSection -Heading 'Collections' -Description 'No. of servers in each Device Collection:' -Detail $pvsCollectionSummary.Detail
                if ($serverListMode) {
                    $summaryIssueText = if ([string]::IsNullOrWhiteSpace($selectionIssueText)) { 'Selection issues: None' } else { $selectionIssueText }
                    $collectionSummaryText = "Server List Mode:$([Environment]::NewLine)Matched: $totalDevices of $($requestedServers.Count) list entries$([Environment]::NewLine)$summaryIssueText$([Environment]::NewLine)$([Environment]::NewLine)$collectionText"
                } else {
                    $collectionSummaryText = $collectionText
                }
                $groupedLogText = "PVS grouped counts | $($pvsCollectionSummary.Detail)"
            }

            $dhcpRequested = $null -ne $dhcpResult -and [bool]$dhcpResult.Requested
            $dhcpDetailsText = if ($dhcpRequested) { Format-PvsDhcpResultText -DhcpResult $dhcpResult } else { '' }
            if ($dhcpRequested) {
                $retrievalStatusText += " | DHCP $($dhcpResult.CompletionStatus), $($dhcpResult.ValidationStatus)"
            }

            $publicationSnapshot = [PSCustomObject]@{
                FullData = $script:pvsFullData
                ItemsSource = $dataGrid.ItemsSource
                SearchText = $txtSearch.Text
                RecordCountText = $lblRecordCount.Text
                CollectionSummaryText = $txtCollectionSummary.Text
                DhcpDetailsText = $txtDhcpFullDetails.Text
                DhcpVisibility = $tabDhcpDetails.Visibility
                SelectedTab = $tabMain.SelectedItem
                StatusText = $statusText.Text
            }
            $txtSearch.Text = ''
            $script:pvsFullData = $deviceInfo
            $dataGrid.ItemsSource = $deviceInfo
            $lblRecordCount.Text = "$totalDevices records"
            $txtCollectionSummary.Text = $collectionSummaryText
            $txtDhcpFullDetails.Text = $dhcpDetailsText
            $tabDhcpDetails.Visibility = if ($dhcpRequested) { 'Visible' } else { 'Collapsed' }
            $statusText.Text = $retrievalStatusText
            if ($dhcpRequested) {
                $tabMain.SelectedItem = $tabDhcpDetails
                $window.UpdateLayout()
                $txtDhcpFullDetails.ScrollToHome()
            } else {
                $tabMain.SelectedIndex = 1
            }
            $publicationCommitted = $true

            if (-not [string]::IsNullOrWhiteSpace($collectionCompletionLogText)) { Write-Log $collectionCompletionLogText }
            if (-not [string]::IsNullOrWhiteSpace($groupedLogText)) { Write-Log $groupedLogText }
            if ($dhcpRequested) {
                Write-Log "PVS DHCP result | Mode=$($dhcpResult.Mode) | Outcome=$($dhcpResult.Outcome) | Completion=$($dhcpResult.CompletionStatus) | Validation=$($dhcpResult.ValidationStatus) | Masters=$($dhcpResult.MastersComplete)/$($dhcpResult.MastersDiscovered) | Calls=$($dhcpResult.RemoteCallCount) | Findings=$($dhcpResult.FindingCount) | DurationMs=$($dhcpResult.DurationMs)"
            } else {
                Write-Log 'PVS DHCP validation skipped (checkbox not selected)'
            }
            Write-Log "PVS retrieval complete: $totalDevices devices"
            if ($singleServerMode) {
                Write-Log "PVS grouped counts skipped (single-server mode: [$targetServer])"
            } elseif ($deviceCollectionMode) {
                Write-Log 'PVS grouped counts replaced by the selected Device Collection summary'
            } elseif ($serverListMode) {
                Write-Log "PVS server-list retrieval complete: matched $totalDevices of $($requestedServers.Count); missing $($missingServers.Count); ambiguous $($ambiguousServers.Count); duplicates $($duplicateServers.Count)"
            }
        } catch {
            $publicationError = $_
            $rollbackSucceeded = $false
            if ($null -ne $publicationSnapshot -and -not $publicationCommitted) {
                try {
                    $script:pvsFullData = $publicationSnapshot.FullData
                    $dataGrid.ItemsSource = $publicationSnapshot.ItemsSource
                    $txtSearch.Text = $publicationSnapshot.SearchText
                    $lblRecordCount.Text = $publicationSnapshot.RecordCountText
                    $txtCollectionSummary.Text = $publicationSnapshot.CollectionSummaryText
                    $txtDhcpFullDetails.Text = $publicationSnapshot.DhcpDetailsText
                    $tabDhcpDetails.Visibility = $publicationSnapshot.DhcpVisibility
                    $tabMain.SelectedItem = $publicationSnapshot.SelectedTab
                    $rollbackSucceeded = $true
                    Write-Log 'PVS result publication rolled back to the previous UI state' -Level WARN
                } catch {
                    Write-Log "PVS result publication rollback failed | $(Get-ErrorDiagnosticText -ErrorRecord $_ -Phase 'PVS UI rollback')" -Level ERROR
                }
            }
            [System.Windows.MessageBox]::Show("Error: $($publicationError.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
            Write-Log "PVS result publication error | $(Get-ErrorDiagnosticText -ErrorRecord $publicationError -Phase 'PVS UI publication') | RollbackSucceeded=$rollbackSucceeded" -Level ERROR
            # Retrieval starts on Execution, so a finish-time snapshot can no
            # longer identify the previously visible result tab. If old rows or
            # DHCP details survived preparation/rollback, expose them explicitly.
            $retainedStateAvailable = $rollbackSucceeded -or $null -eq $publicationSnapshot
            $retainedRowsAvailable = $retainedStateAvailable -and @($script:pvsFullData).Count -gt 0
            $retainedDhcpAvailable = $retainedStateAvailable -and
                [string]$tabDhcpDetails.Visibility -eq 'Visible' -and
                -not [string]::IsNullOrWhiteSpace([string]$txtDhcpFullDetails.Text)
            if ($retainedRowsAvailable) {
                $tabMain.SelectedIndex = 1
            } elseif ($retainedDhcpAvailable) {
                $tabMain.SelectedItem = $tabDhcpDetails
            }
            $statusText.Text = if ($rollbackSucceeded -and ($retainedRowsAvailable -or $retainedDhcpAvailable)) {
                'Result publication failed; previous successful output restored'
            } elseif ($null -eq $publicationSnapshot -and ($retainedRowsAvailable -or $retainedDhcpAvailable)) {
                'Result preparation failed; previous successful output retained'
            } else {
                'Error while publishing retrieved data'
            }
        } finally {
            if ($null -ne $script:pvsExecutionTimer) {
                $script:pvsExecutionTimer.Stop()
                $executionTimeText = Format-ExecutionTime -Elapsed $script:pvsExecutionTimer.Elapsed
                $lblExecutionTime.Text = "Execution: $executionTimeText"
                Write-Log "PVS retrieval execution time: $executionTimeText"
            }
            $progressBar.IsIndeterminate = $false
            $progressBar.Visibility = 'Hidden'
            $lblProgressDetail.Text = ''
            & $setPvsBusyState $false
            $script:pvsActiveRequest = $null
            $script:pvsExecutionTimer = $null
        }
    }

    $pvsPollTimer = New-Object System.Windows.Threading.DispatcherTimer
    $pvsPollTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $pvsPollTimer.Add_Tick({
        $context = $script:pvsRetrievalContext
        if ($null -eq $context) { return }

        & $processPvsWorkerMessages $context
        if ($null -ne $context.Handle -and $context.Handle.IsCompleted -and $context.MessageQueue.IsEmpty) {
            $pvsPollTimer.Stop()
            $completion = Complete-CitrixRetrievalWorker -Context $context
            if ($null -ne $script:pvsRetrievalContext -and $script:pvsRetrievalContext.RunId -eq $context.RunId) {
                $script:pvsRetrievalContext = $null
                & $finishPvsRetrieval $completion
            }
        }
    })

    $btnRetrieve.Add_Click({
        if ($null -ne $script:pvsRetrievalContext -or $null -ne $script:pvsCollectionLoadContext) { return }

        $wantDeviceName = [bool]$chkDevices.IsChecked
        $wantVDisk = [bool]$chkVDisk.IsChecked
        $wantReboot = [bool]$chkReboot.IsChecked
        $wantCSR = [bool]$chkCSR.IsChecked
        $wantXDC = [bool]$chkXDC.IsChecked
        $wantPingReachability = [bool]$chkPingReachability.IsChecked
        $wantServerIP = [bool]$chkServerIP.IsChecked
        $wantDhcpCheck = [bool]$chkDhcpCheck.IsChecked
        $deepDhcpAudit = [bool]$chkDhcpDeepAudit.IsChecked
        $deviceCollectionMode = [bool]$rbDeviceCollection.IsChecked
        $singleServerMode = [bool]$rbSingleServer.IsChecked
        $serverListMode = [bool]$rbServerList.IsChecked
        $selectedServerMode = $singleServerMode -or $serverListMode
        $wholeFarmMode = [bool]$rbWholeFarm.IsChecked
        $activeScopeCount = [int]$wholeFarmMode + [int]$deviceCollectionMode + [int]$singleServerMode + [int]$serverListMode
        if (-not $wholeFarmMode) { $deepDhcpAudit = $false }
        if ($activeScopeCount -ne 1) {
            [System.Windows.MessageBox]::Show(
                'Please select exactly one retrieval scope.',
                'Validation', 'OK', 'Warning') | Out-Null
            $statusText.Text = 'Validation failed'
            return
        }
        try {
            $requestedServers = if ($selectedServerMode) { @(ConvertTo-ServerNameList -InputText $txtServerNames.Text) } else { @() }
        } catch {
            [System.Windows.MessageBox]::Show(
                $_.Exception.Message, 'Server Input Validation', 'OK', 'Warning') | Out-Null
            Write-Log "PVS server input rejected: $($_.Exception.Message)" -Level WARN
            $statusText.Text = 'Server input validation failed'
            return
        }
        $targetServer = if ($singleServerMode -and $requestedServers.Count -eq 1) { $requestedServers[0] } else { $null }

        if ($selectedServerMode -and $requestedServers.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                'Please enter at least one server name for the selected retrieval scope.',
                'Validation', 'OK', 'Warning') | Out-Null
            $statusText.Text = 'Validation failed'
            return
        }
        if ($singleServerMode -and $requestedServers.Count -ne 1) {
            [System.Windows.MessageBox]::Show(
                'Single Server mode accepts exactly one server name. Use Server List mode for multiple names.',
                'Validation', 'OK', 'Warning') | Out-Null
            $statusText.Text = 'Validation failed'
            return
        }
        if ($deviceCollectionMode -and [bool]$pvsCollectionState.Loaded -and -not (& $isPvsCollectionSnapshotFresh)) {
            Write-Log "PVS Device Collection preview expired after $($script:NamedScopeSnapshotMaximumAgeMinutes) minutes; refreshing before retrieval" -Level WARN
            & $startPvsCollectionLoad
            return
        }
        if ($deviceCollectionMode -and (
            -not [bool]$pvsCollectionState.Loaded -or
            [string]::IsNullOrWhiteSpace([string]$pvsCollectionState.InventorySnapshotId) -or
            $null -eq $cmbDeviceCollections.SelectedItem)) {
            [System.Windows.MessageBox]::Show(
                'Wait for the Device Collection list to finish loading, then select one collection before retrieving.',
                'Device Collection Required', 'OK', 'Warning') | Out-Null
            $statusText.Text = 'Select a Device Collection before retrieving'
            return
        }

        # Deep Audit is intentionally a separate, default-No decision. Declining
        # it falls back to the lightweight Whole Farm baseline instead of
        # cancelling the normal PVS retrieval.
        if ($wantDhcpCheck -and $deepDhcpAudit) {
            $deepWarning = @(
                'Deep Full Farm Audit reads every DHCP reservation and its reservation-level options on every PVS Master.'
                ''
                'In a large farm this can create thousands of remote DHCP reads and may take a long time. Run it off-peak and allow only one audit per PVS farm.'
                ''
                'Continue with Deep Audit? Choosing No runs the lightweight Whole Farm baseline instead.'
            ) -join [Environment]::NewLine
            $deepChoice = [System.Windows.MessageBox]::Show(
                $deepWarning,
                'Confirm Deep Full Farm DHCP Audit',
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning,
                [System.Windows.MessageBoxResult]::No)
            if ($deepChoice -ne [System.Windows.MessageBoxResult]::Yes) {
                $deepDhcpAudit = $false
                $chkDhcpDeepAudit.IsChecked = $false
                Write-Log 'Deep DHCP audit declined; continuing with Whole Farm baseline' -Level WARN
            } else {
                Write-Log 'Deep DHCP audit explicitly approved' -Level WARN
            }
        }

        $dhcpMode = if (-not $wantDhcpCheck) {
            'None'
        } elseif ($deepDhcpAudit -and $wholeFarmMode) {
            'DeepAudit'
        } elseif ($wholeFarmMode) {
            'Baseline'
        } else {
            'Targeted'
        }

        $request = @{
            WantDeviceName = $wantDeviceName
            WantVDisk = $wantVDisk
            WantReboot = $wantReboot
            WantCSR = $wantCSR
            WantXDC = $wantXDC
            WantPingReachability = $wantPingReachability
            WantServerIP = $wantServerIP
            WantDhcpCheck = $wantDhcpCheck
            Dhcp = @{
                Enabled = $wantDhcpCheck
                Mode = $dhcpMode
                RetrievalScope = if ($wholeFarmMode) {
                    'WholeFarm'
                } elseif ($deviceCollectionMode) {
                    'DeviceCollection'
                } elseif ($singleServerMode) {
                    'SingleServer'
                } else {
                    'ServerList'
                }
                MasterThrottle = 2
                MasterSoftTimeoutSeconds = if ($deepDhcpAudit) { 300 } else { 90 }
                OverallSoftTimeoutSeconds = 900
                MaxFindings = 250
                TargetConfirmationThreshold = 1000
                MaximumTargetCount = 1000
                MaximumTargetMasterCells = 5000
            }
            WholeFarmMode = $wholeFarmMode
            SingleServerMode = $singleServerMode
            ServerListMode = $serverListMode
            DeviceCollectionMode = $deviceCollectionMode
            SelectedServerMode = $selectedServerMode
            RequestedServers = @($requestedServers)
            TargetServer = $targetServer
            WorkloadConfirmationThreshold = $script:PvsWorkloadConfirmationThreshold
            SelectedCollectionChoiceKey = if ($deviceCollectionMode) { [string]$cmbDeviceCollections.SelectedItem.ChoiceKey } else { '' }
        }

        $selectedFields = @()
        if ($wantDeviceName) { $selectedFields += 'DeviceName' }
        $selectedFields += 'Collection'
        if ($wantVDisk) { $selectedFields += 'vDisk' }
        if ($wantPingReachability) { $selectedFields += 'PingReachability' }
        if ($wantReboot) { $selectedFields += 'RebootDay' }
        if ($wantCSR) { $selectedFields += 'CSRServer' }
        if ($wantXDC) { $selectedFields += 'XDCServer' }
        if ($wantServerIP) { $selectedFields += 'ServerIP' }
        if ($wantDhcpCheck) { $selectedFields += "DHCP:$dhcpMode" }
        Write-Log "PVS retrieval requested - fields: $($selectedFields -join ', ')"
        if ($singleServerMode) {
            Write-Log "PVS retrieval scope: single server [$targetServer]"
        } elseif ($deviceCollectionMode) {
            Write-Log "PVS retrieval scope: Device Collection [$($cmbDeviceCollections.SelectedItem.DisplayName)] | Preview=$($pvsCollectionState.InventorySnapshotId)"
        } elseif ($serverListMode) {
            Write-Log "PVS retrieval scope: server list ($($requestedServers.Count) requested) [$(Format-ServerNamePreview -Names $requestedServers)]"
        } else {
            Write-Log 'PVS retrieval scope: whole farm'
        }

        $previousOutputCount = @($script:pvsFullData).Count
        $statusText.Text = if ($previousOutputCount -gt 0) { 'Starting background PVS retrieval; previous output is retained until completion...' } else { 'Starting background PVS retrieval...' }
        $lblProgressDetail.Text = 'Starting background PVS retrieval...'
        $lblExecutionTime.Text = 'Execution: running...'
        $progressBar.Value = 0
        $progressBar.IsIndeterminate = $true
        $progressBar.Visibility = 'Visible'
        $tabMain.SelectedIndex = 0
        & $setPvsBusyState $true
        Invoke-UiRefresh -Window $window

        $script:pvsActiveRequest = $request
        $script:pvsExecutionTimer = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $script:pvsRetrievalContext = Start-CitrixRetrievalWorker -Role PVS -Request $request
            Write-Log "PVS background retrieval started | RunId=$($script:pvsRetrievalContext.RunId)"
            $pvsPollTimer.Start()
        } catch {
            $script:pvsRetrievalContext = $null
            & $finishPvsRetrieval ([PSCustomObject]@{
                Outcome = 'Failed'
                Result = $null
                ErrorMessage = $_.Exception.Message
                ErrorDetail = Get-ErrorDiagnosticText -ErrorRecord $_ -Phase 'PVS worker startup'
            })
        }
    })

    $btnCancel.Add_Click({
        $context = if ($null -ne $script:pvsRetrievalContext) { $script:pvsRetrievalContext } else { $script:pvsCollectionLoadContext }
        if ($null -eq $context -or [bool]$context.CancellationRequested) { return }
        $operationLabel = if ($null -ne $script:pvsRetrievalContext) { 'retrieval' } else { 'Device Collection list load' }
        Write-Log "PVS $operationLabel cancellation requested | RunId=$($context.RunId)" -Level WARN
        if ($null -ne $script:pvsRetrievalContext) {
            $statusText.Text = 'Cancelling after the current SDK call returns...'
            $lblProgressDetail.Text = 'Cancelling after the current SDK call returns...'
        } else {
            $txtCollectionLoadStatus.Text = 'Cancelling the Device Collection list load...'
        }
        $btnCancel.IsEnabled = $false
        $btnCancel.Content = 'Cancelling...'
        Request-CitrixRetrievalCancellation -Context $context
    })

    $window.Add_Closing({
        param($sender, $e)
        $context = if ($null -ne $script:pvsRetrievalContext) {
            $script:pvsRetrievalContext
        } else {
            $script:pvsCollectionLoadContext
        }
        $operationLabel = if ($null -ne $script:pvsRetrievalContext) { 'retrieval' } else { 'Device Collection list load' }
        if ($null -eq $context -or [bool]$context.Completed) { return }

        $e.Cancel = $true
        if ([bool]$context.CancellationRequested) {
            $forceChoice = [System.Windows.MessageBox]::Show(
                "Cancellation of the $operationLabel has already been requested, but the current Citrix SDK call has not returned.$([Environment]::NewLine)$([Environment]::NewLine)Force-exit the current PowerShell process now?$([Environment]::NewLine)$([Environment]::NewLine)This stops every script running in this PowerShell process and discards the current result.",
                'Force Exit PowerShell?',
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Error,
                [System.Windows.MessageBoxResult]::No)
            if ($forceChoice -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-Log "PVS force exit requested while background $operationLabel remained active | RunId=$($context.RunId)" -Level ERROR
                [Environment]::Exit(2)
            }
            return
        }

        $choice = [System.Windows.MessageBox]::Show(
            "The PVS $operationLabel is still running. Request cancellation and keep this window open until cleanup completes?",
            'PVS Background Work in Progress',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning,
            [System.Windows.MessageBoxResult]::Yes)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) {
            if ($null -ne $script:pvsRetrievalContext) {
                $statusText.Text = 'Cancelling after the current SDK call returns...'
                $lblProgressDetail.Text = 'Cancelling after the current SDK call returns...'
                $btnCancel.IsEnabled = $false
                $btnCancel.Content = 'Cancelling...'
            } else {
                $txtCollectionLoadStatus.Text = 'Cancelling the Device Collection list load...'
                $btnCancel.IsEnabled = $false
                $btnCancel.Content = 'Cancelling...'
            }
            Request-CitrixRetrievalCancellation -Context $context
        }
    })
    # --- Phase 7: Register output actions and open the modal window.
    # Export/copy actions consume the currently displayed (possibly filtered)
    # rows. Copy Server List uses selected displayed rows, or all when none selected.
    $btnExport.Add_Click({ Export-DataGridToCSV -DataGrid $dataGrid -DefaultFileName 'PVS_Devices.csv' })
    $btnCopyServerList.Add_Click({
        Copy-ServerListFromGrid -DataGrid $dataGrid -PropertyName 'ServerName' -FieldDisplayName 'Device Name' -StatusText $statusText -Context 'PVS'
    })
    $btnCopyTable.Add_Click({
        Copy-DataGridTable -DataGrid $dataGrid -StatusText $statusText -Context 'PVS'
    })
    $btnViewLog.Add_Click({ Open-CurrentLog })
    $btnClose.Add_Click({ $window.Close() })

    $window.ShowDialog() | Out-Null
}

# -------------------------------------------------------------
# Region: DDC Mode
# -------------------------------------------------------------
# Builds the XDC GUI and wires Broker-SDK retrieval events. "DDC" remains in the
# code for backward familiarity, while user-facing text uses "XDC".
<#
.SYNOPSIS
    Displays the XDC Server Info window and owns its event-driven workflow.
.DESCRIPTION
    Builds the XAML, binds named controls, maintains independent scope inputs,
    retrieves Broker inventory with paging, shapes selected fields, and creates
    catalog, Delivery Group, and health summaries.
#>
function Show-DDCWindow {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName 'System.Windows.Forms'

    # --- Phase 1: Define the complete XDC window.
    # SharedStyles is interpolated into this XAML before WPF parses it.
    $XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Citrix XDC - Server Info" Height="640" Width="940"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#ECF0F1" MinWidth="760" MinHeight="500">
  $($script:SharedStyles)
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <!-- Window header: product name, purpose, and the local execution host. -->
    <Border Grid.Row="0" Background="{StaticResource HeaderBgBrush}" Padding="20,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0">
          <TextBlock Text="XDC Server Info" FontSize="18" FontWeight="Bold"
                     Foreground="{StaticResource HeaderFgBrush}"/>
          <TextBlock Text="Retrieve server info from Citrix XDC" FontSize="11"
                     Foreground="#AEB6BF" Margin="0,3,0,0"/>
        </StackPanel>
        <TextBlock Grid.Column="1" Name="lblServer" FontSize="11" Foreground="#AEB6BF"
                   VerticalAlignment="Center"/>
      </Grid>
    </Border>

    <TabControl Name="tabMain" Grid.Row="1" Margin="12,10,12,10">
      <!-- Execution owns retrieval scope, selected fields, and commands. -->
      <TabItem Header="Execution">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="{StaticResource CardBgBrush}" CornerRadius="6"
                  BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1"
                  Margin="4,6,4,0" Padding="16,10">
            <StackPanel>
              <StackPanel>
                <TextBlock Text="1. Retrieval Scope" FontSize="13" FontWeight="SemiBold"
                           Foreground="#2C3E50"/>
                <WrapPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="0,8,0,0">
                  <RadioButton Name="rbWholeFarm" Content="Whole Farm" GroupName="DdcScope"
                               IsChecked="True" VerticalAlignment="Center" Margin="0,0,18,0"/>
                  <RadioButton Name="rbMachineCatalog" Content="Machine Catalog (MC)" GroupName="DdcScope"
                               VerticalAlignment="Center" Margin="0,0,18,0"/>
                  <RadioButton Name="rbDeliveryGroup" Content="Delivery Group (DG)" GroupName="DdcScope"
                               VerticalAlignment="Center" Margin="0,0,18,0"/>
                  <RadioButton Name="rbServerList" Content="Server List" GroupName="DdcScope"
                               VerticalAlignment="Center" Margin="0,0,18,0"/>
                  <RadioButton Name="rbSingleServer" Content="Single Server" GroupName="DdcScope"
                               VerticalAlignment="Center"/>
                </WrapPanel>
                <StackPanel Name="pnlDdcScopeSelection" Visibility="Collapsed" Margin="0,8,0,0">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="Auto"/>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Name="lblDdcScopePrompt" Text="Machine Catalog:" FontSize="12"
                               VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <ComboBox Grid.Column="1" Name="cmbDdcScopes" Height="30"
                              DisplayMemberPath="DisplayName" SelectedValuePath="ChoiceKey"
                              IsEditable="False" IsTextSearchEnabled="True" IsEnabled="False"
                              MaxDropDownHeight="320" VerticalContentAlignment="Center"/>
                    <Button Grid.Column="2" Name="btnRefreshDdcScopes" Content="Refresh"
                            Style="{StaticResource SecondaryButton}" Width="82" Height="30"
                            Margin="8,0,0,0" IsEnabled="False"
                            ToolTip="Reload Machine Catalog and Delivery Group choices from Citrix XDC"/>
                  </Grid>
                  <TextBlock Name="txtDdcScopeLoadStatus" Text="Loading Citrix XDC scopes..."
                             FontSize="10" Foreground="#7F8C8D" Margin="0,4,0,0"
                             TextWrapping="Wrap"/>
                </StackPanel>
                <DockPanel Name="pnlServerInput" Margin="0,8,0,0" LastChildFill="True" Visibility="Collapsed">
                  <TextBlock Name="lblServerNamePrompt" Text="Server Name:" FontSize="12"
                             VerticalAlignment="Center" Margin="0,0,8,0"/>
                  <TextBox Name="txtServerNames" Height="30" Style="{StaticResource StyledTextBox}"
                           IsEnabled="False" AcceptsReturn="True" TextWrapping="NoWrap" MaxLength="262144"
                           VerticalScrollBarVisibility="Auto" VerticalAlignment="Center"/>
                </DockPanel>
                <TextBlock Name="txtScopeHint" Text="Whole farm selected; no server names are required."
                           FontSize="10" Foreground="#7F8C8D" Margin="0,4,0,0"/>
              </StackPanel>
              <DockPanel Margin="0,16,0,8">
                <TextBlock Text="2. Data Fields" FontSize="13" FontWeight="SemiBold"
                           Foreground="#2C3E50" VerticalAlignment="Center"/>
                <Button Name="btnToggle" Content="Select All" Style="{StaticResource LinkButton}"
                        HorizontalAlignment="Right" VerticalAlignment="Center"/>
              </DockPanel>
              <WrapPanel Orientation="Horizontal">
                <CheckBox Name="chkMachineName"        Content="Machine Name"        IsChecked="True"  Margin="0,0,18,6"/>
                <CheckBox Name="chkMachineCatalog"     Content="Machine Catalog"     IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkDeliveryGroup"      Content="Delivery Group"      IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkRegistrationState"  Content="Registration State"  IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkInMaintenanceMode"  Content="Maintenance Mode"    IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkLogonMode"          Content="Logon Mode"          IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkServerIP"           Content="Server IP"           IsChecked="False" Margin="0,0,18,6"/>
              </WrapPanel>
            </StackPanel>
          </Border>

          <StackPanel Grid.Row="1" HorizontalAlignment="Center" Margin="4,12,4,0">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
              <Button Name="btnRetrieve" Content="Retrieve Data" Style="{StaticResource PrimaryButton}" Width="140" Height="34" Margin="0,0,10,0"/>
              <Button Name="btnCancel" Content="Cancel" Style="{StaticResource SecondaryButton}" Width="90" Height="34" Margin="0,0,10,0"
                      IsEnabled="False" ToolTip="Request cancellation after the current SDK call returns"/>
              <Button Name="btnExport" Content="Export to CSV" Style="{StaticResource SecondaryButton}" Width="140" Height="34" Margin="0,0,10,0"/>
              <Button Name="btnViewLog" Content="View Log" Style="{StaticResource SecondaryButton}" Width="110" Height="34" Margin="0,0,10,0"
                      ToolTip="Open today's execution log in Notepad"/>
              <Button Name="btnClose" Content="Close" Style="{StaticResource CloseButton}" Width="100" Height="34"/>
            </StackPanel>
            <TextBlock Name="lblLogPath" Text="" Width="650" TextAlignment="Center" TextTrimming="CharacterEllipsis"
                       FontSize="10" Foreground="#7F8C8D" Margin="0,6,0,0"/>
            <TextBlock Text="Run only one copy per XDC site. Same-server duplicate launches are blocked automatically."
                       Width="700" TextAlignment="Center" FontSize="10" Foreground="#7F8C8D" Margin="0,3,0,0"/>
          </StackPanel>

          <Border Grid.Row="2" Background="{StaticResource CardBgBrush}" CornerRadius="6"
                  BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1"
                  Margin="4,12,4,0" Padding="10,6">
            <StackPanel>
              <DockPanel>
                <TextBlock Text="Filter:" FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0" Foreground="#7F8C8D"/>
                <TextBox Name="txtSearch" Style="{StaticResource StyledTextBox}" BorderThickness="0"
                         FontSize="12" VerticalAlignment="Center"/>
              </DockPanel>
              <TextBlock Text="Filters retrieved results only. Enter text to search all columns, or use Field:Value to search one column (examples: MachineName:SERVER01, RegistrationState:Registered). Clear the box to show all rows."
                         FontSize="10" Foreground="#7F8C8D" Margin="0,4,0,0" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>

          <StackPanel Grid.Row="3" Margin="4,8,4,0">
            <TextBlock Name="lblProgressDetail" Text="" FontSize="11" Foreground="#2C3E50"
                       TextAlignment="Center" TextWrapping="Wrap" Margin="0,0,0,4"/>
            <ProgressBar Name="progressBar" Style="{StaticResource AccentProgress}" Visibility="Hidden"/>
          </StackPanel>
        </Grid>
      </TabItem>

      <!-- Output owns MC/DG summaries and the filterable Broker-machine table. -->
      <TabItem Header="Output">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="130" MinHeight="80"/>
            <RowDefinition Height="6"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Border Grid.Row="0" Background="{StaticResource CardBgBrush}" CornerRadius="6"
                  BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1"
                  Margin="4,6,4,0" Padding="10,6">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBox Name="txtCatalogSummary" Grid.Column="0" IsReadOnly="True" BorderThickness="0" Background="Transparent"
                       TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontSize="11"
                       Foreground="#2C3E50"/>
              <TextBox Name="txtDeliverySummary" Grid.Column="2" IsReadOnly="True" BorderThickness="0" Background="Transparent"
                       TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontSize="11"
                       Foreground="#2C3E50"/>
            </Grid>
          </Border>

          <GridSplitter Grid.Row="1" Height="6" HorizontalAlignment="Stretch" VerticalAlignment="Center"
                        Background="#D5DBDB" ResizeDirection="Rows" ShowsPreview="True"/>

          <DataGrid Name="dataGrid" Grid.Row="2" Margin="4,10,4,8" AutoGenerateColumns="True"
                    IsReadOnly="True" AlternatingRowBackground="{StaticResource GridAltRowBrush}"
                    GridLinesVisibility="All" CanUserSortColumns="True"
                    Background="White" BorderBrush="{StaticResource CardBorderBrush}" BorderThickness="1"
                    HorizontalGridLinesBrush="#E5E8E8" VerticalGridLinesBrush="#E5E8E8" RowHeight="32"
                    ColumnHeaderHeight="34"
                    SelectionMode="Extended" SelectionUnit="FullRow"/>

          <Border Grid.Row="3" Background="{StaticResource StatusBgBrush}" Padding="16,7" Margin="4,0,4,4">
            <DockPanel>
              <Button Name="btnCopyTable" Content="Copy Table" Style="{StaticResource PrimaryButton}"
                      Width="100" Height="26" Padding="8,2" FontSize="10" DockPanel.Dock="Right" Margin="8,-2,0,-2"
                      ToolTip="Copy the complete currently displayed table, including headings"/>
              <Button Name="btnCopyServerList" Content="Copy Server List" Style="{StaticResource PrimaryButton}"
                      Width="125" Height="26" Padding="8,2" FontSize="10" DockPanel.Dock="Right" Margin="16,-2,0,-2"
                      ToolTip="Copy selected server names, or all currently displayed names when no rows are selected"/>
              <TextBlock Name="lblRecordCount" Text="" FontSize="11"
                         Foreground="{StaticResource StatusFgBrush}" DockPanel.Dock="Right"/>
              <TextBlock Name="lblExecutionTime" Text="" FontSize="11" Margin="0,0,18,0"
                         Foreground="{StaticResource StatusFgBrush}" DockPanel.Dock="Right"/>
              <TextBlock Name="statusText" Text="Ready" FontSize="11"
                         Foreground="{StaticResource StatusFgBrush}"/>
            </DockPanel>
          </Border>
        </Grid>
      </TabItem>
    </TabControl>
  </Grid>
</Window>
"@

    # --- Phase 2: Parse XAML and obtain controls used by event handlers.
    $window = [Windows.Markup.XamlReader]::Parse($XAML)

    # Event-handler scriptblocks close over these local control references for
    # the lifetime of the modal window.
    $tabMain              = $window.FindName('tabMain')
    $chkMachineName       = $window.FindName('chkMachineName')
    $chkMachineCatalog    = $window.FindName('chkMachineCatalog')
    $chkDeliveryGroup     = $window.FindName('chkDeliveryGroup')
    $chkRegistrationState = $window.FindName('chkRegistrationState')
    $chkInMaintenanceMode = $window.FindName('chkInMaintenanceMode')
    $chkLogonMode         = $window.FindName('chkLogonMode')
    $chkServerIP          = $window.FindName('chkServerIP')
    $rbWholeFarm          = $window.FindName('rbWholeFarm')
    $rbMachineCatalog     = $window.FindName('rbMachineCatalog')
    $rbDeliveryGroup      = $window.FindName('rbDeliveryGroup')
    $rbSingleServer       = $window.FindName('rbSingleServer')
    $rbServerList         = $window.FindName('rbServerList')
    $pnlDdcScopeSelection = $window.FindName('pnlDdcScopeSelection')
    $lblDdcScopePrompt    = $window.FindName('lblDdcScopePrompt')
    $cmbDdcScopes         = $window.FindName('cmbDdcScopes')
    $btnRefreshDdcScopes  = $window.FindName('btnRefreshDdcScopes')
    $txtDdcScopeLoadStatus = $window.FindName('txtDdcScopeLoadStatus')
    $pnlServerInput       = $window.FindName('pnlServerInput')
    $lblServerNamePrompt  = $window.FindName('lblServerNamePrompt')
    $txtServerNames       = $window.FindName('txtServerNames')
    $txtScopeHint         = $window.FindName('txtScopeHint')
    $btnToggle     = $window.FindName('btnToggle')
    $btnRetrieve   = $window.FindName('btnRetrieve')
    $btnCancel     = $window.FindName('btnCancel')
    $btnExport     = $window.FindName('btnExport')
    $btnViewLog    = $window.FindName('btnViewLog')
    $btnClose      = $window.FindName('btnClose')
    $txtSearch     = $window.FindName('txtSearch')
    $dataGrid      = $window.FindName('dataGrid')
    $btnCopyTable  = $window.FindName('btnCopyTable')
    $btnCopyServerList = $window.FindName('btnCopyServerList')
    $progressBar   = $window.FindName('progressBar')
    $lblProgressDetail = $window.FindName('lblProgressDetail')
    $txtCatalogSummary = $window.FindName('txtCatalogSummary')
    $txtDeliverySummary = $window.FindName('txtDeliverySummary')
    $statusText    = $window.FindName('statusText')
    $lblRecordCount = $window.FindName('lblRecordCount')
    $lblExecutionTime = $window.FindName('lblExecutionTime')
    $lblLogPath    = $window.FindName('lblLogPath')
    $lblServer     = $window.FindName('lblServer')

    Set-DataGridDisplayFormat -DataGrid $dataGrid -Window $window

    $lblServer.Text = "$($env:COMPUTERNAME)"
    $lblLogPath.Text = "Daily log: $($script:LogPath)"
    $lblLogPath.ToolTip = "Daily log, appended by date. Contains retrieval scope and fields, counts, warnings, errors, export paths, and execution time.`n$($script:LogPath)"

    # --- Phase 3: Initialize per-window state.
    # This is the unfiltered result; filtering changes only DataGrid.ItemsSource.
    $script:ddcFullData = $null

    # --- Phase 4: Register selection, scope, filter, and retrieval handlers.
    $btnToggle.Add_Click({
        $allChecks = @($chkMachineName, $chkMachineCatalog, $chkDeliveryGroup,
                       $chkRegistrationState, $chkInMaintenanceMode, $chkLogonMode, $chkServerIP)
        $anyUnchecked = $allChecks | Where-Object { -not $_.IsChecked }
        $newState = [bool]$anyUnchecked
        foreach ($cb in $allChecks) { $cb.IsChecked = $newState }
        $btnToggle.Content = if ($newState) { 'Deselect All' } else { 'Select All' }
    })

    # One visible TextBox is reused for both named scopes. Separate buffers stop a
    # pasted Server List from appearing when the user switches to Single Server.
    $ddcScopeInputState = @{
        CurrentScope = 'WholeFarm'
        SingleServer = ''
        ServerList = ''
    }

    # MC and DG share one Broker inventory preview. Both choice lists are derived
    # from it, so switching between the radio buttons is immediate. Retrieve never
    # opens a modal picker; it revalidates the selected key against one fresh site
    # inventory so registration and other live values are current.
    $script:ddcRetrievalContext = $null
    $script:ddcScopeLoadContext = $null
    $ddcNamedScopeState = @{
        Loaded = $false
        Inventory = @()
        InventorySnapshotId = ''
        MachineCatalogChoices = @()
        DeliveryGroupChoices = @()
        SelectedMachineCatalogKey = ''
        SelectedDeliveryGroupKey = ''
        MachineCatalogUnsafeMetadataMachineCount = 0
        DeliveryGroupUnassignedMachineCount = 0
        DeliveryGroupUnsafeMetadataMachineCount = 0
        LoadedAt = $null
        LastError = ''
        Binding = $false
    }

    $isDdcNamedScopeSnapshotFresh = {
        return [bool]$ddcNamedScopeState.Loaded -and $null -ne $ddcNamedScopeState.LoadedAt -and
            ((Get-Date) - $ddcNamedScopeState.LoadedAt).TotalMinutes -le $script:NamedScopeSnapshotMaximumAgeMinutes
    }

    $updateDdcRetrieveEnabled = {
        $mainBusy = $null -ne $script:ddcRetrievalContext
        $listBusy = $null -ne $script:ddcScopeLoadContext
        $namedScope = [bool]$rbMachineCatalog.IsChecked -or [bool]$rbDeliveryGroup.IsChecked
        $namedScopeReady = (& $isDdcNamedScopeSnapshotFresh) -and $null -ne $cmbDdcScopes.SelectedItem
        $btnRetrieve.IsEnabled = -not $mainBusy -and -not $listBusy -and (-not $namedScope -or $namedScopeReady)
    }

    $bindDdcScopeSelector = {
        $machineCatalogSelected = [bool]$rbMachineCatalog.IsChecked
        $deliveryGroupSelected = [bool]$rbDeliveryGroup.IsChecked
        if (-not $machineCatalogSelected -and -not $deliveryGroupSelected) {
            $pnlDdcScopeSelection.Visibility = 'Collapsed'
            & $updateDdcRetrieveEnabled
            return
        }

        $pnlDdcScopeSelection.Visibility = 'Visible'
        $scopeLabel = if ($machineCatalogSelected) { 'Machine Catalog' } else { 'Delivery Group' }
        $lblDdcScopePrompt.Text = "${scopeLabel}:"
        $choices = if ($machineCatalogSelected) {
            @($ddcNamedScopeState.MachineCatalogChoices)
        } else {
            @($ddcNamedScopeState.DeliveryGroupChoices)
        }
        $savedKey = if ($machineCatalogSelected) {
            [string]$ddcNamedScopeState.SelectedMachineCatalogKey
        } else {
            [string]$ddcNamedScopeState.SelectedDeliveryGroupKey
        }
        $loading = $null -ne $script:ddcScopeLoadContext

        $ddcNamedScopeState.Binding = $true
        try {
            $cmbDdcScopes.ItemsSource = $null
            if ([bool]$ddcNamedScopeState.Loaded) {
                $cmbDdcScopes.ItemsSource = @($choices)
                $restoreChoice = @($choices | Where-Object { $_.ChoiceKey -ieq $savedKey } | Select-Object -First 1)
                if ($restoreChoice.Count -eq 1) {
                    $cmbDdcScopes.SelectedItem = $restoreChoice[0]
                } else {
                    $cmbDdcScopes.SelectedIndex = -1
                    if ($machineCatalogSelected) {
                        $ddcNamedScopeState.SelectedMachineCatalogKey = ''
                    } else {
                        $ddcNamedScopeState.SelectedDeliveryGroupKey = ''
                    }
                }
            } else {
                $cmbDdcScopes.SelectedIndex = -1
            }
        } finally {
            $ddcNamedScopeState.Binding = $false
        }

        $snapshotFresh = & $isDdcNamedScopeSnapshotFresh
        $cmbDdcScopes.IsEnabled = -not $loading -and $snapshotFresh -and $choices.Count -gt 0 -and $null -eq $script:ddcRetrievalContext
        $btnRefreshDdcScopes.IsEnabled = -not $loading -and $null -eq $script:ddcRetrievalContext
        if ($loading) {
            $txtDdcScopeLoadStatus.Text = "Loading Machine Catalogs and Delivery Groups from Citrix XDC..."
            $txtDdcScopeLoadStatus.Foreground = '#7F8C8D'
        } elseif ([bool]$ddcNamedScopeState.Loaded -and -not $snapshotFresh) {
            $txtDdcScopeLoadStatus.Text = "This Citrix XDC scope list is older than $($script:NamedScopeSnapshotMaximumAgeMinutes) minutes. Click Refresh before retrieving."
            $txtDdcScopeLoadStatus.Foreground = '#922B21'
        } elseif ([bool]$ddcNamedScopeState.Loaded -and $choices.Count -gt 0) {
            $detailText = if ($machineCatalogSelected -and [int]$ddcNamedScopeState.MachineCatalogUnsafeMetadataMachineCount -gt 0) {
                " $($ddcNamedScopeState.MachineCatalogUnsafeMetadataMachineCount) machine record(s) with unsafe catalog identity metadata were excluded."
            } elseif ($deliveryGroupSelected) {
                $parts = @()
                if ([int]$ddcNamedScopeState.DeliveryGroupUnassignedMachineCount -gt 0) {
                    $parts += "$($ddcNamedScopeState.DeliveryGroupUnassignedMachineCount) unassigned machine record(s) do not appear."
                }
                if ([int]$ddcNamedScopeState.DeliveryGroupUnsafeMetadataMachineCount -gt 0) {
                    $parts += "$($ddcNamedScopeState.DeliveryGroupUnsafeMetadataMachineCount) machine record(s) with unsafe group identity metadata were excluded."
                }
                if ($parts.Count -gt 0) { ' ' + ($parts -join ' ') } else { '' }
            } else { '' }
            $loadedText = if ($null -ne $ddcNamedScopeState.LoadedAt) { "Loaded $($ddcNamedScopeState.LoadedAt.ToString('HH:mm:ss')). " } else { '' }
            $txtDdcScopeLoadStatus.Text = "${loadedText}Select one $scopeLabel. $($choices.Count) available from $(@($ddcNamedScopeState.Inventory).Count) site machines.$detailText"
            $txtDdcScopeLoadStatus.Foreground = '#7F8C8D'
        } elseif ([bool]$ddcNamedScopeState.Loaded) {
            $txtDdcScopeLoadStatus.Text = "No selectable $scopeLabel choices were found. Click Refresh to try again."
            $txtDdcScopeLoadStatus.Foreground = '#922B21'
        } elseif (-not [string]::IsNullOrWhiteSpace([string]$ddcNamedScopeState.LastError)) {
            $txtDdcScopeLoadStatus.Text = "Could not load Citrix XDC scopes: $($ddcNamedScopeState.LastError) Click Refresh to retry."
            $txtDdcScopeLoadStatus.Foreground = '#922B21'
        }
        & $updateDdcRetrieveEnabled
    }

    $finishDdcScopeLoad = {
        param($Completion)

        $btnCancel.IsEnabled = $false
        $btnCancel.Content = 'Cancel'
        if ([string]$Completion.Outcome -eq 'Completed' -and $null -ne $Completion.Result -and [bool]$Completion.Result.InventoryOnly) {
            $result = $Completion.Result
            $ddcNamedScopeState.Inventory = @($result.Inventory)
            $ddcNamedScopeState.InventorySnapshotId = [string]$result.InventorySnapshotId
            $ddcNamedScopeState.MachineCatalogChoices = @($result.MachineCatalogChoices)
            $ddcNamedScopeState.DeliveryGroupChoices = @($result.DeliveryGroupChoices)
            $ddcNamedScopeState.MachineCatalogUnsafeMetadataMachineCount = [int]$result.MachineCatalogUnsafeMetadataMachineCount
            $ddcNamedScopeState.DeliveryGroupUnassignedMachineCount = [int]$result.DeliveryGroupUnassignedMachineCount
            $ddcNamedScopeState.DeliveryGroupUnsafeMetadataMachineCount = [int]$result.DeliveryGroupUnsafeMetadataMachineCount
            $ddcNamedScopeState.LoadedAt = Get-Date
            $ddcNamedScopeState.SelectedMachineCatalogKey = ''
            $ddcNamedScopeState.SelectedDeliveryGroupKey = ''
            $ddcNamedScopeState.LastError = ''
            $ddcNamedScopeState.Loaded = $true
            Write-Log "XDC inline scope lists loaded | MachineCatalogs=$(@($result.MachineCatalogChoices).Count) | DeliveryGroups=$(@($result.DeliveryGroupChoices).Count) | Inventory=$($result.InventoryCount) | Snapshot=$($result.InventorySnapshotId)"
        } elseif ([string]$Completion.Outcome -eq 'Cancelled') {
            $ddcNamedScopeState.LastError = 'Loading was cancelled.'
            Write-Log 'XDC inline scope-list load cancelled' -Level WARN
        } else {
            $loadError = [string]$Completion.ErrorMessage
            if ([string]::IsNullOrWhiteSpace($loadError)) { $loadError = 'The background inventory load returned no result.' }
            $ddcNamedScopeState.LastError = $loadError
            Write-Log "XDC inline scope-list load failed: $loadError | Detail: $([string]$Completion.ErrorDetail)" -Level ERROR
        }
        & $bindDdcScopeSelector
        if ([string]$Completion.Outcome -eq 'Completed' -and
            ([bool]$rbMachineCatalog.IsChecked -or [bool]$rbDeliveryGroup.IsChecked) -and
            $cmbDdcScopes.Items.Count -gt 0) {
            [void]$cmbDdcScopes.Focus()
            $cmbDdcScopes.IsDropDownOpen = $true
        }
    }

    $ddcScopeLoadTimer = New-Object System.Windows.Threading.DispatcherTimer
    $ddcScopeLoadTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $ddcScopeLoadTimer.Add_Tick({
        $context = $script:ddcScopeLoadContext
        if ($null -eq $context) { return }

        foreach ($message in @(Receive-CitrixRetrievalMessages -Context $context)) {
            if ([string]$message.Kind -eq 'Log') {
                Write-RetrievalWorkerMessageLog -Message $message -Context $context
            } elseif ([string]$message.Kind -eq 'Progress' -and
                ([bool]$rbMachineCatalog.IsChecked -or [bool]$rbDeliveryGroup.IsChecked) -and
                -not [bool]$context.CancellationRequested) {
                $txtDdcScopeLoadStatus.Text = [string]$message.Message
            }
        }

        if ($null -ne $context.Handle -and $context.Handle.IsCompleted -and $context.MessageQueue.IsEmpty) {
            $ddcScopeLoadTimer.Stop()
            $completion = Complete-CitrixRetrievalWorker -Context $context
            if ($null -ne $script:ddcScopeLoadContext -and $script:ddcScopeLoadContext.RunId -eq $context.RunId) {
                $script:ddcScopeLoadContext = $null
                & $finishDdcScopeLoad $completion
            }
        }
    })

    $startDdcScopeLoad = {
        if ($null -ne $script:ddcRetrievalContext -or $null -ne $script:ddcScopeLoadContext) { return }

        $ddcNamedScopeState.Loaded = $false
        $ddcNamedScopeState.Inventory = @()
        $ddcNamedScopeState.InventorySnapshotId = ''
        $ddcNamedScopeState.MachineCatalogChoices = @()
        $ddcNamedScopeState.DeliveryGroupChoices = @()
        $ddcNamedScopeState.SelectedMachineCatalogKey = ''
        $ddcNamedScopeState.SelectedDeliveryGroupKey = ''
        $ddcNamedScopeState.MachineCatalogUnsafeMetadataMachineCount = 0
        $ddcNamedScopeState.DeliveryGroupUnassignedMachineCount = 0
        $ddcNamedScopeState.DeliveryGroupUnsafeMetadataMachineCount = 0
        $ddcNamedScopeState.LoadedAt = $null
        $ddcNamedScopeState.LastError = ''
        $cmbDdcScopes.ItemsSource = $null
        $cmbDdcScopes.SelectedIndex = -1
        & $bindDdcScopeSelector

        try {
            $script:ddcScopeLoadContext = Start-CitrixRetrievalWorker -Role DDC -Request @{ ScopeInventoryOnly = $true }
            $btnCancel.Content = 'Cancel Load'
            $btnCancel.IsEnabled = $true
            Write-Log "XDC inline scope-list load started | RunId=$($script:ddcScopeLoadContext.RunId)"
            $ddcScopeLoadTimer.Start()
        } catch {
            $script:ddcScopeLoadContext = $null
            & $finishDdcScopeLoad ([PSCustomObject]@{
                Outcome = 'Failed'
                Result = $null
                ErrorMessage = $_.Exception.Message
                ErrorDetail = Get-ErrorDiagnosticText -ErrorRecord $_ -Phase 'XDC scope-preview worker startup'
            })
        }
        & $bindDdcScopeSelector
    }

    $cmbDdcScopes.Add_SelectionChanged({
        if ([bool]$ddcNamedScopeState.Binding) { return }
        $choiceKey = if ($null -ne $cmbDdcScopes.SelectedItem) { [string]$cmbDdcScopes.SelectedItem.ChoiceKey } else { '' }
        if ([bool]$rbMachineCatalog.IsChecked) {
            $ddcNamedScopeState.SelectedMachineCatalogKey = $choiceKey
        } elseif ([bool]$rbDeliveryGroup.IsChecked) {
            $ddcNamedScopeState.SelectedDeliveryGroupKey = $choiceKey
        }
        & $updateDdcRetrieveEnabled
    })
    $btnRefreshDdcScopes.Add_Click({ & $startDdcScopeLoad })

    $updateDdcScopeControls = {
        # Save the outgoing scope before loading the incoming one. Whole Farm
        # clears only the visible control and preserves both saved values.
        if ($ddcScopeInputState.CurrentScope -eq 'SingleServer') {
            $ddcScopeInputState.SingleServer = [string]$txtServerNames.Text
        } elseif ($ddcScopeInputState.CurrentScope -eq 'ServerList') {
            $ddcScopeInputState.ServerList = [string]$txtServerNames.Text
        }

        $newScope = if ($rbWholeFarm.IsChecked) {
            'WholeFarm'
        } elseif ($rbMachineCatalog.IsChecked) {
            'MachineCatalog'
        } elseif ($rbDeliveryGroup.IsChecked) {
            'DeliveryGroup'
        } elseif ($rbSingleServer.IsChecked) {
            'SingleServer'
        } else {
            'ServerList'
        }

        $selectedServerScope = [bool]$rbSingleServer.IsChecked -or [bool]$rbServerList.IsChecked
        $txtServerNames.IsEnabled = $selectedServerScope

        if ($newScope -eq 'WholeFarm') {
            $pnlServerInput.Visibility = 'Collapsed'
            $txtServerNames.Text = ''
            $txtServerNames.Height = 30
            $txtScopeHint.Text = 'Whole farm selected; no server names are required.'
        } elseif ($newScope -eq 'MachineCatalog') {
            $pnlServerInput.Visibility = 'Collapsed'
            $txtServerNames.Text = ''
            $txtServerNames.Height = 30
            $txtScopeHint.Text = 'Select a Machine Catalog below. The current Citrix XDC list loads automatically.'
        } elseif ($newScope -eq 'DeliveryGroup') {
            $pnlServerInput.Visibility = 'Collapsed'
            $txtServerNames.Text = ''
            $txtServerNames.Height = 30
            $txtScopeHint.Text = 'Select a Delivery Group below. The current Citrix XDC list loads automatically.'
        } elseif ($newScope -eq 'SingleServer') {
            $pnlServerInput.Visibility = 'Visible'
            $lblServerNamePrompt.Text = 'Server Name:'
            $txtServerNames.Height = 30
            $txtServerNames.Text = [string]$ddcScopeInputState.SingleServer
            $txtScopeHint.Text = 'Enter one short name, FQDN, or DOMAIN\SERVER value.'
            $txtServerNames.Focus()
        } else {
            $pnlServerInput.Visibility = 'Visible'
            $lblServerNamePrompt.Text = 'Server Names:'
            $txtServerNames.Height = 62
            $txtServerNames.Text = [string]$ddcScopeInputState.ServerList
            $txtScopeHint.Text = 'Paste server names separated by new lines, commas, or semicolons.'
            $txtServerNames.Focus()
        }

        $ddcScopeInputState.CurrentScope = $newScope
        & $bindDdcScopeSelector
        if (($newScope -eq 'MachineCatalog' -or $newScope -eq 'DeliveryGroup') -and
            -not (& $isDdcNamedScopeSnapshotFresh) -and $null -eq $script:ddcScopeLoadContext) {
            & $startDdcScopeLoad
        }
        & $updateDdcRetrieveEnabled
    }
    $rbWholeFarm.Add_Checked($updateDdcScopeControls)
    $rbMachineCatalog.Add_Checked($updateDdcScopeControls)
    $rbDeliveryGroup.Add_Checked($updateDdcScopeControls)
    $rbSingleServer.Add_Checked($updateDdcScopeControls)
    $rbServerList.Add_Checked($updateDdcScopeControls)
    & $updateDdcScopeControls

    # Filtering is local and non-destructive; it never reruns Get-BrokerMachine.
    # Unknown Field:Value fields fall back to ordinary all-column text matching.
    $txtSearch.Add_TextChanged({
        if ($null -eq $script:ddcFullData) { return }
        $filterResult = Apply-GridFilter -Data $script:ddcFullData -Term $txtSearch.Text
        $dataGrid.ItemsSource = $filterResult.Rows
        $lblRecordCount.Text = $filterResult.Summary
    })

    # --- Phase 5: Main XDC retrieval pipeline.
    # Every Retrieve action pages one current Broker inventory in the worker.
    # MC/DG revalidate the previewed opaque key before local filtering; DNS
    # fallback and row shaping also stay in that worker, which never accesses WPF.
    $script:ddcActiveRequest = $null
    $script:ddcExecutionTimer = $null

    $ddcInputControls = @(
        $rbWholeFarm, $rbMachineCatalog, $rbDeliveryGroup, $rbServerList, $rbSingleServer, $txtServerNames, $btnToggle,
        $cmbDdcScopes, $btnRefreshDdcScopes,
        $chkMachineName, $chkMachineCatalog, $chkDeliveryGroup,
        $chkRegistrationState, $chkInMaintenanceMode, $chkLogonMode, $chkServerIP,
        $txtSearch, $btnCopyServerList, $btnCopyTable
    )

    $setDdcBusyState = {
        param([bool]$Busy)

        foreach ($control in $ddcInputControls) {
            if ($null -ne $control) { $control.IsEnabled = -not $Busy }
        }
        if (-not $Busy) {
            $txtServerNames.IsEnabled = [bool]$rbSingleServer.IsChecked -or [bool]$rbServerList.IsChecked
            & $bindDdcScopeSelector
        } else {
            $cmbDdcScopes.IsEnabled = $false
            $btnRefreshDdcScopes.IsEnabled = $false
        }

        $btnRetrieve.Content = if ($Busy) { 'Retrieving...' } else { 'Retrieve Data' }
        $btnCancel.IsEnabled = $Busy
        $btnCancel.Content = 'Cancel'
        $btnExport.IsEnabled = -not $Busy
        $btnViewLog.IsEnabled = -not $Busy
        $btnClose.IsEnabled = -not $Busy
        $window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { [System.Windows.Input.Cursors]::Arrow }
        if ($Busy) { $btnRetrieve.IsEnabled = $false } else { & $updateDdcRetrieveEnabled }
    }

    $processDdcWorkerMessages = {
        param($Context)

        foreach ($message in @(Receive-CitrixRetrievalMessages -Context $Context)) {
            switch ([string]$message.Kind) {
                'Log' {
                    Write-RetrievalWorkerMessageLog -Message $message -Context $Context
                }
                'Progress' {
                    if ([bool]$Context.CancellationRequested) { continue }
                    $statusText.Text = [string]$message.Message
                    $lblProgressDetail.Text = [string]$message.Message
                    $progressBar.IsIndeterminate = [bool]$message.IsIndeterminate
                    if (-not [bool]$message.IsIndeterminate -and [int]$message.Total -gt 0) {
                        $progressBar.Value = ([double][int]$message.Current / [double][int]$message.Total) * 100
                    }
                }
            }
        }
    }

    $finishDdcRetrieval = {
        param($Completion)

        $btnCancel.IsEnabled = $false
        $btnCancel.Content = 'Cancel'
        $request = $script:ddcActiveRequest
        $publicationSnapshot = $null
        $publicationCommitted = $false
        try {
            if ([string]$Completion.Outcome -eq 'Cancelled') {
                $hasPreviousOutput = @($script:ddcFullData).Count -gt 0
                $statusText.Text = if ($hasPreviousOutput) { 'Retrieval cancelled; previous successful output retained' } else { 'Retrieval cancelled; no partial table was published' }
                $tabMain.SelectedIndex = if ($hasPreviousOutput) { 1 } else { 0 }
                Write-Log "DDC retrieval cancelled by user | PreviousOutputRetained=$hasPreviousOutput" -Level WARN
                return
            }

            if ([string]$Completion.Outcome -ne 'Completed' -or $null -eq $Completion.Result) {
                $errorText = [string]$Completion.ErrorMessage
                if ([string]::IsNullOrWhiteSpace($errorText)) { $errorText = 'The background retrieval failed without an error message.' }
                [System.Windows.MessageBox]::Show("Error: $errorText", 'Error', 'OK', 'Error') | Out-Null
                $errorDetail = [string]$Completion.ErrorDetail
                Write-Log "DDC background retrieval error: $errorText | Detail: $errorDetail" -Level ERROR
                $hasPreviousOutput = @($script:ddcFullData).Count -gt 0
                $statusText.Text = if ($hasPreviousOutput) { 'Retrieval failed; previous successful output retained' } else { 'Error during retrieval' }
                $tabMain.SelectedIndex = if ($hasPreviousOutput) { 1 } else { 0 }
                return
            }

            $result = $Completion.Result
            $machines = @($result.SourceItems)
            $machineInfo = @($result.Data)
            $missingServers = @($result.MissingNames)
            $ambiguousServers = @($result.AmbiguousNames)
            $duplicateServers = @($result.DuplicateNames)
            $totalMachines = $machineInfo.Count
            $requestedServers = @($request.RequestedServers)
            $singleServerMode = [bool]$request.SingleServerMode
            $serverListMode = [bool]$request.ServerListMode
            $selectedServerMode = [bool]$request.SelectedServerMode
            $machineCatalogMode = [bool]$request.MachineCatalogMode
            $deliveryGroupMode = [bool]$request.DeliveryGroupMode
            $namedScopeMode = $machineCatalogMode -or $deliveryGroupMode
            $selectedScopeChoice = $result.SelectedScopeChoice
            $unselectableScopeMachineCount = [int]$result.UnselectableScopeMachineCount
            $unassignedScopeMachineCount = [int]$result.UnassignedScopeMachineCount
            $unsafeScopeMetadataMachineCount = [int]$result.UnsafeScopeMetadataMachineCount
            $targetServer = [string]$request.TargetServer

            Write-Log "DDC shared inventory returned $($result.AllInventoryCount) machines; selected $totalMachines"
            if ($missingServers.Count -gt 0) {
                Write-Log "DDC requested servers not found ($($missingServers.Count)): $(Format-ServerNamePreview -Names $missingServers)" -Level WARN
            }
            if ($ambiguousServers.Count -gt 0) {
                Write-Log "DDC requested servers are ambiguous ($($ambiguousServers.Count)): $(Format-ServerNamePreview -Names $ambiguousServers)" -Level WARN
            }
            if ($duplicateServers.Count -gt 0) {
                Write-Log "DDC duplicate server aliases ignored ($($duplicateServers.Count)): $(Format-ServerNamePreview -Names $duplicateServers)" -Level WARN
            }

            $selectionIssueText = Format-ServerSelectionIssues -MissingNames $missingServers -AmbiguousNames $ambiguousServers -DuplicateNames $duplicateServers
            if (-not [string]::IsNullOrWhiteSpace($selectionIssueText) -and $totalMachines -gt 0) {
                [System.Windows.MessageBox]::Show(
                    "Matched $totalMachines of $($requestedServers.Count) requested entries. Retrieval completed for the unambiguous matches.$([Environment]::NewLine)$([Environment]::NewLine)$selectionIssueText",
                    'Server List Needs Attention', 'OK', 'Warning') | Out-Null
            }

            if ($totalMachines -eq 0) {
                if ($selectedServerMode) {
                    $requestedPreview = Format-ServerNamePreview -Names $requestedServers
                    $selectionDetail = if ([string]::IsNullOrWhiteSpace($selectionIssueText)) { "Requested: $requestedPreview" } else { $selectionIssueText }
                    [System.Windows.MessageBox]::Show(
                        "No unambiguous requested server matches were available in the Broker machine list.$([Environment]::NewLine)$([Environment]::NewLine)$selectionDetail",
                        'Information', 'OK', 'Information') | Out-Null
                } elseif ($namedScopeMode) {
                    $scopeLabel = if ($machineCatalogMode) { 'Machine Catalog' } else { 'Delivery Group' }
                    $scopeName = if ($null -ne $selectedScopeChoice) { [string]$selectedScopeChoice.ScopeName } else { 'selected scope' }
                    [System.Windows.MessageBox]::Show(
                        "No Broker machines matched $scopeLabel [$scopeName] in the retrieved site snapshot.",
                        'Information', 'OK', 'Information') | Out-Null
                } else {
                    [System.Windows.MessageBox]::Show('No Broker Machines found.', 'Information', 'OK', 'Information') | Out-Null
                }
                # A completed empty query is authoritative and replaces an older
                # table. Snapshot first so any WPF assignment failure can restore
                # the previous internally consistent publication.
                $publicationSnapshot = [PSCustomObject]@{
                    FullData = $script:ddcFullData
                    ItemsSource = $dataGrid.ItemsSource
                    SearchText = $txtSearch.Text
                    RecordCountText = $lblRecordCount.Text
                    CatalogSummaryText = $txtCatalogSummary.Text
                    DeliverySummaryText = $txtDeliverySummary.Text
                    SelectedTab = $tabMain.SelectedItem
                    StatusText = $statusText.Text
                }
                $txtSearch.Text = ''
                $script:ddcFullData = @()
                $dataGrid.ItemsSource = @()
                $lblRecordCount.Text = '0 records'
                $txtCatalogSummary.Text = ''
                $txtDeliverySummary.Text = ''
                $statusText.Text = 'No machines found'
                $tabMain.SelectedIndex = 0
                $publicationCommitted = $true
                Write-Log 'No selected Broker machines were available' -Level WARN
                return
            }

            # Build all summaries before swapping the table so a formatter failure
            # cannot mix a new dataset with old status and grouping text.
            $ddcSummary = Get-DdcHealthSummary -Machines $machines
            $retrievalStatus = if ($serverListMode) {
                "Retrieved $totalMachines of $($requestedServers.Count) list entries | Not found: $($missingServers.Count) | Ambiguous: $($ambiguousServers.Count) | Duplicates: $($duplicateServers.Count)"
            } elseif ($machineCatalogMode) {
                "Retrieved $totalMachines machines from Machine Catalog [$($selectedScopeChoice.ScopeName)] | Site inventory: $($result.AllInventoryCount)"
            } elseif ($deliveryGroupMode) {
                "Retrieved $totalMachines machines from Delivery Group [$($selectedScopeChoice.ScopeName)] | Site inventory: $($result.AllInventoryCount)"
            } else {
                "Retrieved $totalMachines machines"
            }
            $retrievalStatusText = "$retrievalStatus | Maintenance: $($ddcSummary.MaintenanceCount) | Unregistered: $($ddcSummary.UnregisteredCount) | Logon Disabled: $($ddcSummary.LogonDisabledCount) | Logon Unavailable: $($ddcSummary.LogonUnavailableCount)"

            $catalogSummaryText = ''
            $deliverySummaryText = ''
            $catalogGroupedLogText = ''
            $deliveryGroupedLogText = ''
            if ($singleServerMode) {
                $catalogSummaryText = "Single Server Mode:$([Environment]::NewLine) - MC aggregation skipped for [$targetServer]"
                $deliverySummaryText = "Single Server Mode:$([Environment]::NewLine) - DG aggregation skipped for [$targetServer]"
            } else {
                $catalogSummary = Get-GroupedCountSummary -Items $machines -PropertyName 'CatalogName' -Label 'Catalogs'
                $deliveryGroupSummary = Get-GroupedCountSummary -Items $machines -PropertyName 'DesktopGroupName' -Label 'DeliveryGroups'
                $catalogText = Format-CountSection -Heading 'Catalogs' -Description 'No. of Machines in each MC:' -Detail $catalogSummary.Detail
                $deliveryGroupText = Format-CountSection -Heading 'DeliveryGroups' -Description 'No. of Machines in each DG:' -Detail $deliveryGroupSummary.Detail
                $selectedScopeIdentity = if ($null -ne $selectedScopeChoice) {
                    "$($selectedScopeChoice.IdentifierType): $($selectedScopeChoice.IdentifierValue)"
                } else {
                    'Unavailable'
                }
                if ($serverListMode) {
                    $summaryIssueText = if ([string]::IsNullOrWhiteSpace($selectionIssueText)) { 'Selection issues: None' } else { $selectionIssueText }
                    $catalogSummaryText = "Server List Mode:$([Environment]::NewLine)Matched: $totalMachines of $($requestedServers.Count) list entries$([Environment]::NewLine)$summaryIssueText$([Environment]::NewLine)$([Environment]::NewLine)$catalogText"
                    $deliverySummaryText = $deliveryGroupText
                } elseif ($machineCatalogMode) {
                    $catalogSummaryText = "Machine Catalog Scope:$([Environment]::NewLine)Selected: $($selectedScopeChoice.ScopeName)$([Environment]::NewLine)Identity: $selectedScopeIdentity$([Environment]::NewLine)Matched: $totalMachines of $($result.AllInventoryCount) site machines$([Environment]::NewLine)Unsafe metadata excluded: $unsafeScopeMetadataMachineCount"
                    $deliverySummaryText = $deliveryGroupText
                } elseif ($deliveryGroupMode) {
                    $catalogSummaryText = $catalogText
                    $deliverySummaryText = "Delivery Group Scope:$([Environment]::NewLine)Selected: $($selectedScopeChoice.ScopeName)$([Environment]::NewLine)Identity: $selectedScopeIdentity$([Environment]::NewLine)Matched: $totalMachines of $($result.AllInventoryCount) site machines$([Environment]::NewLine)Unassigned: $unassignedScopeMachineCount$([Environment]::NewLine)Unsafe metadata excluded: $unsafeScopeMetadataMachineCount"
                } else {
                    $catalogSummaryText = $catalogText
                    $deliverySummaryText = $deliveryGroupText
                }
                $catalogGroupedLogText = "DDC grouped counts | $($catalogSummary.Detail)"
                $deliveryGroupedLogText = "DDC grouped counts | $($deliveryGroupSummary.Detail)"
            }

            $publicationSnapshot = [PSCustomObject]@{
                FullData = $script:ddcFullData
                ItemsSource = $dataGrid.ItemsSource
                SearchText = $txtSearch.Text
                RecordCountText = $lblRecordCount.Text
                CatalogSummaryText = $txtCatalogSummary.Text
                DeliverySummaryText = $txtDeliverySummary.Text
                SelectedTab = $tabMain.SelectedItem
                StatusText = $statusText.Text
            }
            $txtSearch.Text = ''
            $script:ddcFullData = $machineInfo
            $dataGrid.ItemsSource = $machineInfo
            $lblRecordCount.Text = "$totalMachines records"
            $txtCatalogSummary.Text = $catalogSummaryText
            $txtDeliverySummary.Text = $deliverySummaryText
            $statusText.Text = $retrievalStatusText
            $tabMain.SelectedIndex = 1
            $publicationCommitted = $true

            if (-not [string]::IsNullOrWhiteSpace($catalogGroupedLogText)) { Write-Log $catalogGroupedLogText }
            if (-not [string]::IsNullOrWhiteSpace($deliveryGroupedLogText)) { Write-Log $deliveryGroupedLogText }
            Write-Log "DDC retrieval complete: $totalMachines machines"
            Write-Log "DDC summary counts | Maintenance=$($ddcSummary.MaintenanceCount) | Unregistered=$($ddcSummary.UnregisteredCount) | LogonDisabled=$($ddcSummary.LogonDisabledCount) | LogonUnavailable=$($ddcSummary.LogonUnavailableCount)"
            if ($singleServerMode) {
                Write-Log "DDC grouped counts skipped (single-server mode: [$targetServer])"
            } elseif ($serverListMode) {
                Write-Log "DDC server-list retrieval complete: matched $totalMachines of $($requestedServers.Count); missing $($missingServers.Count); ambiguous $($ambiguousServers.Count); duplicates $($duplicateServers.Count)"
            } elseif ($namedScopeMode) {
                $scopeLabel = if ($machineCatalogMode) { 'Machine Catalog' } else { 'Delivery Group' }
                Write-Log "DDC $scopeLabel retrieval complete | Name=[$($selectedScopeChoice.ScopeName)] | Identity=$($selectedScopeChoice.IdentifierType):$($selectedScopeChoice.IdentifierValue) | Matched=$totalMachines | SiteInventory=$($result.AllInventoryCount) | Unassigned=$unassignedScopeMachineCount | UnsafeMetadata=$unsafeScopeMetadataMachineCount | ListExcluded=$unselectableScopeMachineCount"
            }
        } catch {
            $publicationError = $_
            $rollbackSucceeded = $false
            if ($null -ne $publicationSnapshot -and -not $publicationCommitted) {
                try {
                    $script:ddcFullData = $publicationSnapshot.FullData
                    $dataGrid.ItemsSource = $publicationSnapshot.ItemsSource
                    $txtSearch.Text = $publicationSnapshot.SearchText
                    $lblRecordCount.Text = $publicationSnapshot.RecordCountText
                    $txtCatalogSummary.Text = $publicationSnapshot.CatalogSummaryText
                    $txtDeliverySummary.Text = $publicationSnapshot.DeliverySummaryText
                    $tabMain.SelectedItem = $publicationSnapshot.SelectedTab
                    $rollbackSucceeded = $true
                    Write-Log 'DDC result publication rolled back to the previous UI state' -Level WARN
                } catch {
                    Write-Log "DDC result publication rollback failed | $(Get-ErrorDiagnosticText -ErrorRecord $_ -Phase 'XDC UI rollback')" -Level ERROR
                }
            }
            [System.Windows.MessageBox]::Show("Error: $($publicationError.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
            Write-Log "DDC result publication error | $(Get-ErrorDiagnosticText -ErrorRecord $publicationError -Phase 'XDC UI publication') | RollbackSucceeded=$rollbackSucceeded" -Level ERROR
            $retainedStateAvailable = $rollbackSucceeded -or $null -eq $publicationSnapshot
            $retainedRowsAvailable = $retainedStateAvailable -and @($script:ddcFullData).Count -gt 0
            if ($retainedRowsAvailable) { $tabMain.SelectedIndex = 1 }
            $statusText.Text = if ($rollbackSucceeded -and $retainedRowsAvailable) {
                'Result publication failed; previous successful output restored'
            } elseif ($null -eq $publicationSnapshot -and $retainedRowsAvailable) {
                'Result preparation failed; previous successful output retained'
            } else {
                'Error while publishing retrieved data'
            }
        } finally {
            if ($null -ne $script:ddcExecutionTimer) {
                $script:ddcExecutionTimer.Stop()
                $executionTimeText = Format-ExecutionTime -Elapsed $script:ddcExecutionTimer.Elapsed
                $lblExecutionTime.Text = "Execution: $executionTimeText"
                Write-Log "DDC retrieval execution time: $executionTimeText"
            }
            $progressBar.IsIndeterminate = $false
            $progressBar.Visibility = 'Hidden'
            $lblProgressDetail.Text = ''
            & $setDdcBusyState $false
            $script:ddcActiveRequest = $null
            $script:ddcExecutionTimer = $null
        }
    }

    $ddcPollTimer = New-Object System.Windows.Threading.DispatcherTimer
    $ddcPollTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $ddcPollTimer.Add_Tick({
        $context = $script:ddcRetrievalContext
        if ($null -eq $context) { return }

        & $processDdcWorkerMessages $context
        if ($null -ne $context.Handle -and $context.Handle.IsCompleted -and $context.MessageQueue.IsEmpty) {
            $ddcPollTimer.Stop()
            $completion = Complete-CitrixRetrievalWorker -Context $context
            if ($null -ne $script:ddcRetrievalContext -and $script:ddcRetrievalContext.RunId -eq $context.RunId) {
                $script:ddcRetrievalContext = $null
                & $finishDdcRetrieval $completion
            }
        }
    })

    $btnRetrieve.Add_Click({
        if ($null -ne $script:ddcRetrievalContext -or $null -ne $script:ddcScopeLoadContext) { return }

        $wantName = [bool]$chkMachineName.IsChecked
        $wantCatalog = [bool]$chkMachineCatalog.IsChecked
        $wantDeliveryGroup = [bool]$chkDeliveryGroup.IsChecked
        $wantRegistrationState = [bool]$chkRegistrationState.IsChecked
        $wantMaintenanceMode = [bool]$chkInMaintenanceMode.IsChecked
        $wantLogonMode = [bool]$chkLogonMode.IsChecked
        $wantServerIP = [bool]$chkServerIP.IsChecked
        $wholeFarmMode = [bool]$rbWholeFarm.IsChecked
        $machineCatalogMode = [bool]$rbMachineCatalog.IsChecked
        $deliveryGroupMode = [bool]$rbDeliveryGroup.IsChecked
        $singleServerMode = [bool]$rbSingleServer.IsChecked
        $serverListMode = [bool]$rbServerList.IsChecked
        $activeScopeCount = [int]$wholeFarmMode + [int]$machineCatalogMode + [int]$deliveryGroupMode + [int]$singleServerMode + [int]$serverListMode
        $selectedServerMode = $singleServerMode -or $serverListMode

        if ($activeScopeCount -ne 1) {
            [System.Windows.MessageBox]::Show(
                'Please select exactly one retrieval scope.',
                'Validation', 'OK', 'Warning') | Out-Null
            $statusText.Text = 'Validation failed'
            return
        }
        try {
            $requestedServers = if ($selectedServerMode) { @(ConvertTo-ServerNameList -InputText $txtServerNames.Text) } else { @() }
        } catch {
            [System.Windows.MessageBox]::Show(
                $_.Exception.Message, 'Server Input Validation', 'OK', 'Warning') | Out-Null
            Write-Log "XDC server input rejected: $($_.Exception.Message)" -Level WARN
            $statusText.Text = 'Server input validation failed'
            return
        }
        $targetServer = if ($singleServerMode -and $requestedServers.Count -eq 1) { $requestedServers[0] } else { $null }
        if ($selectedServerMode -and $requestedServers.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                'Please enter at least one server name for the selected retrieval scope.',
                'Validation', 'OK', 'Warning') | Out-Null
            $statusText.Text = 'Validation failed'
            return
        }
        if ($singleServerMode -and $requestedServers.Count -ne 1) {
            [System.Windows.MessageBox]::Show(
                'Single Server mode accepts exactly one server name. Use Server List mode for multiple names.',
                'Validation', 'OK', 'Warning') | Out-Null
            $statusText.Text = 'Validation failed'
            return
        }
        $namedScopeMode = $machineCatalogMode -or $deliveryGroupMode
        if ($namedScopeMode -and [bool]$ddcNamedScopeState.Loaded -and -not (& $isDdcNamedScopeSnapshotFresh)) {
            Write-Log "XDC named-scope preview expired after $($script:NamedScopeSnapshotMaximumAgeMinutes) minutes; refreshing before retrieval" -Level WARN
            & $startDdcScopeLoad
            return
        }
        if ($namedScopeMode -and (
            -not [bool]$ddcNamedScopeState.Loaded -or
            [string]::IsNullOrWhiteSpace([string]$ddcNamedScopeState.InventorySnapshotId) -or
            $null -eq $cmbDdcScopes.SelectedItem)) {
            $scopeLabel = if ($machineCatalogMode) { 'Machine Catalog' } else { 'Delivery Group' }
            [System.Windows.MessageBox]::Show(
                "Wait for the Citrix XDC scope list to finish loading, then select one $scopeLabel before retrieving.",
                "$scopeLabel Required", 'OK', 'Warning') | Out-Null
            $statusText.Text = "Select a $scopeLabel before retrieving"
            return
        }

        $request = @{
            WantName = $wantName
            WantCatalog = $wantCatalog
            WantDeliveryGroup = $wantDeliveryGroup
            WantRegistrationState = $wantRegistrationState
            WantMaintenanceMode = $wantMaintenanceMode
            WantLogonMode = $wantLogonMode
            WantServerIP = $wantServerIP
            WholeFarmMode = $wholeFarmMode
            MachineCatalogMode = $machineCatalogMode
            DeliveryGroupMode = $deliveryGroupMode
            SingleServerMode = $singleServerMode
            ServerListMode = $serverListMode
            SelectedServerMode = $selectedServerMode
            RequestedServers = @($requestedServers)
            TargetServer = $targetServer
            SelectedDdcScopeChoiceKey = if ($namedScopeMode) { [string]$cmbDdcScopes.SelectedItem.ChoiceKey } else { '' }
        }

        $selectedFields = @()
        if ($wantName) { $selectedFields += 'MachineName' }
        if ($wantCatalog) { $selectedFields += 'CatalogName' }
        if ($wantDeliveryGroup) { $selectedFields += 'DeliveryGroup' }
        if ($wantRegistrationState) { $selectedFields += 'RegistrationState' }
        if ($wantMaintenanceMode) { $selectedFields += 'MaintenanceMode' }
        if ($wantLogonMode) { $selectedFields += 'LogonMode' }
        if ($wantServerIP) { $selectedFields += 'ServerIP' }
        Write-Log "DDC retrieval requested - fields: $($selectedFields -join ', ')"
        if ($singleServerMode) {
            Write-Log "DDC retrieval scope: single server [$targetServer]"
        } elseif ($serverListMode) {
            Write-Log "DDC retrieval scope: server list ($($requestedServers.Count) requested) [$(Format-ServerNamePreview -Names $requestedServers)]"
        } elseif ($machineCatalogMode) {
            Write-Log "DDC retrieval scope: Machine Catalog [$($cmbDdcScopes.SelectedItem.DisplayName)] | Preview=$($ddcNamedScopeState.InventorySnapshotId)"
        } elseif ($deliveryGroupMode) {
            Write-Log "DDC retrieval scope: Delivery Group [$($cmbDdcScopes.SelectedItem.DisplayName)] | Preview=$($ddcNamedScopeState.InventorySnapshotId)"
        } else {
            Write-Log 'DDC retrieval scope: whole farm'
        }

        $previousOutputCount = @($script:ddcFullData).Count
        $statusText.Text = if ($previousOutputCount -gt 0) { 'Starting background Broker retrieval; previous output is retained until completion...' } else { 'Starting background Broker retrieval...' }
        $lblProgressDetail.Text = 'Starting background Broker retrieval...'
        $lblExecutionTime.Text = 'Execution: running...'
        $progressBar.Value = 0
        $progressBar.IsIndeterminate = $true
        $progressBar.Visibility = 'Visible'
        $tabMain.SelectedIndex = 0
        & $setDdcBusyState $true
        Invoke-UiRefresh -Window $window

        $script:ddcActiveRequest = $request
        $script:ddcExecutionTimer = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $script:ddcRetrievalContext = Start-CitrixRetrievalWorker -Role DDC -Request $request
            Write-Log "DDC background retrieval started | RunId=$($script:ddcRetrievalContext.RunId)"
            $ddcPollTimer.Start()
        } catch {
            $script:ddcRetrievalContext = $null
            & $finishDdcRetrieval ([PSCustomObject]@{
                Outcome = 'Failed'
                Result = $null
                ErrorMessage = $_.Exception.Message
                ErrorDetail = Get-ErrorDiagnosticText -ErrorRecord $_ -Phase 'XDC worker startup'
            })
        }
    })

    $btnCancel.Add_Click({
        $context = if ($null -ne $script:ddcRetrievalContext) { $script:ddcRetrievalContext } else { $script:ddcScopeLoadContext }
        if ($null -eq $context -or [bool]$context.CancellationRequested) { return }
        $operationLabel = if ($null -ne $script:ddcRetrievalContext) { 'retrieval' } else { 'MC/DG list load' }
        Write-Log "DDC $operationLabel cancellation requested | RunId=$($context.RunId)" -Level WARN
        if ($null -ne $script:ddcRetrievalContext) {
            $statusText.Text = 'Cancelling after the current SDK call returns...'
            $lblProgressDetail.Text = 'Cancelling after the current SDK call returns...'
        } else {
            $txtDdcScopeLoadStatus.Text = 'Cancelling the MC/DG list load...'
        }
        $btnCancel.IsEnabled = $false
        $btnCancel.Content = 'Cancelling...'
        Request-CitrixRetrievalCancellation -Context $context
    })

    $window.Add_Closing({
        param($sender, $e)
        $context = if ($null -ne $script:ddcRetrievalContext) {
            $script:ddcRetrievalContext
        } else {
            $script:ddcScopeLoadContext
        }
        $operationLabel = if ($null -ne $script:ddcRetrievalContext) { 'retrieval' } else { 'MC/DG list load' }
        if ($null -eq $context -or [bool]$context.Completed) { return }

        $e.Cancel = $true
        if ([bool]$context.CancellationRequested) {
            $forceChoice = [System.Windows.MessageBox]::Show(
                "Cancellation of the $operationLabel has already been requested, but the current Citrix SDK call has not returned.$([Environment]::NewLine)$([Environment]::NewLine)Force-exit the current PowerShell process now?$([Environment]::NewLine)$([Environment]::NewLine)This stops every script running in this PowerShell process and discards the current result.",
                'Force Exit PowerShell?',
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Error,
                [System.Windows.MessageBoxResult]::No)
            if ($forceChoice -eq [System.Windows.MessageBoxResult]::Yes) {
                Write-Log "DDC force exit requested while background $operationLabel remained active | RunId=$($context.RunId)" -Level ERROR
                [Environment]::Exit(2)
            }
            return
        }

        $choice = [System.Windows.MessageBox]::Show(
            "The XDC $operationLabel is still running. Request cancellation and keep this window open until cleanup completes?",
            'XDC Background Work in Progress',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning,
            [System.Windows.MessageBoxResult]::Yes)
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) {
            if ($null -ne $script:ddcRetrievalContext) {
                $statusText.Text = 'Cancelling after the current SDK call returns...'
                $lblProgressDetail.Text = 'Cancelling after the current SDK call returns...'
                $btnCancel.IsEnabled = $false
                $btnCancel.Content = 'Cancelling...'
            } else {
                $txtDdcScopeLoadStatus.Text = 'Cancelling the MC/DG list load...'
                $btnCancel.IsEnabled = $false
                $btnCancel.Content = 'Cancelling...'
            }
            Request-CitrixRetrievalCancellation -Context $context
        }
    })
    # --- Phase 7: Register output actions and open the modal window.
    # Export/copy actions consume the currently displayed (possibly filtered)
    # rows. Copy Server List uses selected displayed rows, or all when none selected.
    $btnExport.Add_Click({ Export-DataGridToCSV -DataGrid $dataGrid -DefaultFileName 'Broker_MachineData.csv' })
    $btnCopyServerList.Add_Click({
        Copy-ServerListFromGrid -DataGrid $dataGrid -PropertyName 'MachineName' -FieldDisplayName 'Machine Name' -StatusText $statusText -Context 'DDC'
    })
    $btnCopyTable.Add_Click({
        Copy-DataGridTable -DataGrid $dataGrid -StatusText $statusText -Context 'DDC'
    })
    $btnViewLog.Add_Click({ Open-CurrentLog })
    $btnClose.Add_Click({ $window.Close() })

    $window.ShowDialog() | Out-Null
}

# -------------------------------------------------------------
# Region: Startup and Role Prerequisites
# -------------------------------------------------------------
# These checks return structured details instead of exiting directly. The entry
# point decides whether a finding is blocking and is responsible for user output.
<#
.SYNOPSIS
    Validates platform prerequisites that are common to both Citrix roles.
.DESCRIPTION
    Checks the supported Windows Server generation, PowerShell 5.1, and WPF
    availability. Issues block launch; warnings and environment details are
    returned for the daily log.
.OUTPUTS
    PSCustomObject with Passed, Issues, Warnings, and Details.
#>
function Test-StartupPrereqs {
    $issues = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $details = [System.Collections.Generic.List[string]]::new()

    $caption = 'Unknown OS'
    $osVersion = [Environment]::OSVersion.Version
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        if ($os.Caption) { $caption = [string]$os.Caption }
    } catch {
        $warnings.Add("Unable to read OS caption: $($_.Exception.Message)") | Out-Null
    }
    $details.Add("OS: $caption ($($osVersion.ToString()))") | Out-Null

    $supportedOs = $false
    if ($caption -match 'Windows Server 2012' -or
        $caption -match 'Windows Server 2016' -or
        $caption -match 'Windows Server 2019' -or
        $caption -match 'Windows Server 2022') {
        $supportedOs = $true
    } elseif (($osVersion.Major -eq 6 -and $osVersion.Minor -ge 2) -or $osVersion.Major -ge 10) {
        $supportedOs = $true
    }
    if (-not $supportedOs) {
        $issues.Add("Unsupported OS detected: $caption ($($osVersion.ToString())). Supported: Windows Server 2012/2016/2019/2022.") | Out-Null
    }

    $psv = $PSVersionTable.PSVersion
    # Do not name this local variable $psEdition: PowerShell variable names are
    # case-insensitive, so that spelling collides with the constant automatic
    # variable $PSEdition and prevents both the PVS and XDC GUIs from starting.
    $detectedPsEdition = if ($PSVersionTable.PSObject.Properties.Match('PSEdition').Count -gt 0) { [string]$PSVersionTable.PSEdition } else { 'Desktop' }
    $details.Add("PowerShell: $psv ($detectedPsEdition)") | Out-Null
    if ($psv -lt [version]'5.1') {
        $issues.Add("PowerShell 5.1 or later is required. Current version: $psv") | Out-Null
    }
    if ($detectedPsEdition -ne 'Desktop') {
        $issues.Add("Windows PowerShell 5.1 (Desktop edition) is required for the installed Citrix snap-ins. Current edition: $detectedPsEdition $psv") | Out-Null
    }

    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        $details.Add('WPF: Available') | Out-Null
    } catch {
        $issues.Add("WPF assembly load failed (PresentationFramework): $($_.Exception.Message)") | Out-Null
    }

    return [PSCustomObject]@{
        Passed = ($issues.Count -eq 0)
        Issues = @($issues)
        Warnings = @($warnings)
        Details = @($details)
    }
}

# Confirms that a fresh STA runspace can load the same Citrix command used by
# the background worker. A command available only in the UI runspace would pass
# the older precheck but fail as soon as Retrieve Data was clicked.
function Test-CitrixWorkerRunspaceCommand {
    param(
        [string]$CommandName
    )

    $runspace = $null
    $pipeline = $null
    try {
        $runspace = [RunspaceFactory]::CreateRunspace()
        $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
        $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $runspace.Open()

        $pipeline = [PowerShell]::Create()
        $pipeline.Runspace = $runspace
        $initializerText = Get-RetrievalWorkerFunctionText -FunctionNames @((Get-CitrixWorkerDependencyManifest).Probe)
        $probeScript = @(
            'param($RequiredCommand)'
            $initializerText
            'Initialize-CitrixWorkerCommand -CommandName $RequiredCommand'
            'if ($null -eq $script:TrustedCitrixWorkerCommandInfo) { throw ''Trusted Citrix command metadata was not produced.'' }'
            '$command = $script:TrustedCitrixWorkerCommandInfo'
            '[PSCustomObject]@{ Kind = ''CitrixWorkerProbeResult''; CommandName = [string]$command.CommandName; CommandType = [string]$command.CommandType; Source = [string]$command.Source; ModulePath = [string]$command.ModulePath; SnapInName = [string]$command.SnapInName }'
        ) -join "`r`n"
        $probeResult = @($pipeline.AddScript($probeScript).AddArgument($CommandName).Invoke())
        $successResult = @(
            $probeResult | Where-Object {
                $_.PSObject.Properties.Match('Kind').Count -gt 0 -and
                $_.Kind -eq 'CitrixWorkerProbeResult' -and
                $_.CommandName -ieq $CommandName
            }
        ) | Select-Object -First 1

        # Do not use PowerShell.HadErrors here. Windows PowerShell can keep that
        # flag set after an expected lookup miss handled with SilentlyContinue,
        # even when the error stream is empty and the command-ready marker was
        # returned successfully.
        if ($null -eq $successResult) {
            $probeErrors = @($pipeline.Streams.Error)
            $message = if ($probeErrors.Count -gt 0) {
                [string]$probeErrors[-1]
            } elseif ($null -ne $pipeline.InvocationStateInfo.Reason) {
                [string]$pipeline.InvocationStateInfo.Reason.Message
            } else {
                "The worker runspace probe completed without confirming command [$CommandName]."
            }
            return [PSCustomObject]@{ Passed = $false; ErrorMessage = $message }
        }
        return [PSCustomObject]@{ Passed = $true; ErrorMessage = '' }
    } catch {
        return [PSCustomObject]@{ Passed = $false; ErrorMessage = $_.Exception.Message }
    } finally {
        if ($null -ne $pipeline) { $pipeline.Dispose() }
        if ($null -ne $runspace) {
            try { $runspace.Close() } catch { }
            $runspace.Dispose()
        }
    }
}

<#
.SYNOPSIS
    Validates SDK prerequisites for the detected Citrix role.
.DESCRIPTION
    Attempts to load Citrix snap-ins, then verifies the core PVS or Broker
    command. Missing DHCP cmdlets are a warning because normal PVS retrieval can
    still run; missing core Citrix cmdlets are blocking issues.
.PARAMETER Role
    PVS for Get-PVSDevice validation, or DDC for Get-BrokerMachine validation.
.OUTPUTS
    PSCustomObject with Passed, Issues, Warnings, and Details.
#>
function Test-RolePrereqs {
    param(
        [ValidateSet('PVS','DDC')]
        [string]$Role
    )

    $issues = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $details = [System.Collections.Generic.List[string]]::new()

    $requiredCommandName = if ($Role -eq 'PVS') { 'Get-PVSDevice' } else { 'Get-BrokerMachine' }
    $roleLoadError = ''
    try {
        Initialize-CitrixWorkerCommand -CommandName $requiredCommandName
    } catch {
        # Continue so the role-specific message below remains concise while the
        # loader detail is retained for troubleshooting.
        $roleLoadError = $_.Exception.Message
    }

    if ($Role -eq 'PVS') {
        if (-not $roleLoadError) {
            $details.Add("Citrix PVS cmdlets: Available from trusted source [$($script:TrustedCitrixWorkerCommandInfo.Source)]") | Out-Null
            $workerCheck = Test-CitrixWorkerRunspaceCommand -CommandName 'Get-PVSDevice'
            if ($workerCheck.Passed) {
                $details.Add('Citrix PVS background runspace: Available') | Out-Null
            } else {
                $issues.Add("Citrix PVS SDK cannot load in the background runspace: $($workerCheck.ErrorMessage)") | Out-Null
            }
        } else {
            $message = 'Citrix PVS cmdlets not found (Get-PVSDevice missing). Install/load Citrix PVS PowerShell SDK.'
            if ($roleLoadError) { $message += " Details: $roleLoadError" }
            $issues.Add($message) | Out-Null
        }

        try {
            Initialize-TrustedDhcpServerModule
            $details.Add('DHCP cmdlets: Available from trusted Windows DHCPServer module') | Out-Null
        } catch {
            $warnings.Add("Trusted DHCP cmdlets unavailable. Normal PVS retrieval still works; DHCP validation will report Unavailable. Details: $($_.Exception.Message)") | Out-Null
        }
    }

    if ($Role -eq 'DDC') {
        if (-not $roleLoadError) {
            $details.Add("Citrix Broker cmdlets: Available from trusted source [$($script:TrustedCitrixWorkerCommandInfo.Source)]") | Out-Null
            $workerCheck = Test-CitrixWorkerRunspaceCommand -CommandName 'Get-BrokerMachine'
            if ($workerCheck.Passed) {
                $details.Add('Citrix Broker background runspace: Available') | Out-Null
            } else {
                $issues.Add("Citrix Broker SDK cannot load in the background runspace: $($workerCheck.ErrorMessage)") | Out-Null
            }
        } else {
            $message = 'Citrix Broker cmdlets not found (Get-BrokerMachine missing). Install/load Citrix Broker PowerShell SDK.'
            if ($roleLoadError) { $message += " Details: $roleLoadError" }
            $issues.Add($message) | Out-Null
        }
    }

    return [PSCustomObject]@{
        Passed = ($issues.Count -eq 0)
        Issues = @($issues)
        Warnings = @($warnings)
        Details = @($details)
    }
}

# Acquires one machine-wide mutex for this tool. This prevents duplicate copies
# on the same Windows server, including separate RDP sessions. It deliberately
# does not claim to coordinate copies launched on other PVS Masters or DDCs.
function Enter-CitrixDataPullSingleInstance {
    $mutexName = 'Global\CitrixDataPullScript'
    $mutex = $null
    try {
        $mutex = [System.Threading.Mutex]::new($false, $mutexName)
        $acquired = $false
        try {
            $acquired = $mutex.WaitOne(0, $false)
        } catch [System.Threading.AbandonedMutexException] {
            # The prior process ended without cleanup; ownership transfers here.
            $acquired = $true
        }

        if (-not $acquired) {
            $mutex.Dispose()
            return [PSCustomObject]@{
                Acquired = $false
                Mutex = $null
                FailureKind = 'Contended'
                ErrorMessage = 'Another copy is already running on this Windows server.'
            }
        }

        return [PSCustomObject]@{
            Acquired = $true
            Mutex = $mutex
            FailureKind = ''
            ErrorMessage = ''
        }
    } catch {
        if ($null -ne $mutex) { $mutex.Dispose() }
        return [PSCustomObject]@{
            Acquired = $false
            Mutex = $null
            FailureKind = 'VerificationError'
            ErrorMessage = "Unable to verify the same-server single-instance lock: $($_.Exception.Message)"
        }
    }
}

function Exit-CitrixDataPullSingleInstance {
    param(
        $LockResult
    )

    if ($null -eq $LockResult -or -not [bool]$LockResult.Acquired -or $null -eq $LockResult.Mutex) { return }
    try {
        $LockResult.Mutex.ReleaseMutex()
    } catch {
        Write-Log "Unable to release the same-server single-instance lock cleanly: $($_.Exception.Message)" -Level WARN
    } finally {
        try { $LockResult.Mutex.Dispose() } catch {
            Write-Log "Unable to dispose the same-server single-instance lock cleanly: $($_.Exception.Message)" -Level WARN
        }
    }
}

# -------------------------------------------------------------
# Region: Entry Point
# -------------------------------------------------------------
# The entry point is intentionally last so every helper exists before role
# detection begins. All exit paths record their outcome in the daily log.
# Safety boundary: an unhandled terminating error must end the script. `break`
# preserves a nonzero process status and still allows an enclosing finally block
# to release the machine-wide mutex; `continue` would resume partial state.
trap {
    $fatalMessage = if ($null -ne $_.Exception) { $_.Exception.Message } else { [string]$_ }
    Write-Log "UNHANDLED EXCEPTION | $(Get-ErrorDiagnosticText -ErrorRecord $_ -Phase 'Top-level entry point')" -Level ERROR
    try {
        [System.Windows.MessageBox]::Show(
            "An unexpected error occurred:`n$fatalMessage`n`nSee log for details:`n$($script:LogPath)",
            'Fatal Error', 'OK', 'Error') | Out-Null
    } catch {
        Write-Warning "Fatal error: $fatalMessage. Log: $($script:LogPath)"
    }
    break
}

Write-Log '====== Script started ======'
$scriptPathForAudit = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) { [IO.Path]::GetFullPath($PSCommandPath) } else { '<interactive-or-unknown>' }
$scriptSha256 = if ($scriptPathForAudit -ne '<interactive-or-unknown>') { Get-FileSha256Hex -LiteralPath $scriptPathForAudit } else { 'Unavailable' }
Write-Log "Release identity | ToolRelease=$($script:ToolRelease) | ScriptPath=[$scriptPathForAudit] | SHA256=$scriptSha256"
try {
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -FilePath $scriptPathForAudit -ErrorAction Stop
    $signerSubject = if ($null -ne $signature.SignerCertificate) { [string]$signature.SignerCertificate.Subject } else { 'None' }
    $signatureLevel = if ([string]$signature.Status -eq 'Valid') { 'INFO' } else { 'WARN' }
    Write-Log "Authenticode | Status=$($signature.Status) | Signer=[$signerSubject] | StatusMessage=[$($signature.StatusMessage)]" -Level $signatureLevel
} catch {
    Write-Log "Authenticode status could not be read: $($_.Exception.Message)" -Level WARN
}
try {
    $windowsIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $windowsPrincipal = [Security.Principal.WindowsPrincipal]::new($windowsIdentity)
    $identityName = [string]$windowsIdentity.Name
    $isElevated = $windowsPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {
    $identityName = [string]$env:USERNAME
    $isElevated = 'Unknown'
}
Write-Log "Host context | Host=$($env:COMPUTERNAME) | Identity=[$identityName] | Elevated=$isElevated | PSVersion=$($PSVersionTable.PSVersion) | PSEdition=$($PSVersionTable.PSEdition)"

$startupCheck = Test-StartupPrereqs
if ($null -eq $startupCheck) {
    $startupCheck = [PSCustomObject]@{
        Passed = $false
        Issues = @('Startup prerequisite validation did not return a result. Review the preceding log entries for the underlying error.')
        Warnings = @()
        Details = @()
    }
}
$startupDetails = @($startupCheck.Details | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
$startupWarnings = @($startupCheck.Warnings | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
$startupIssues = @($startupCheck.Issues | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
if (-not [bool]$startupCheck.Passed -and $startupIssues.Count -eq 0) {
    $startupIssues = @('Startup prerequisite validation failed without issue details. Review the preceding log entries for the underlying error.')
}
foreach ($line in $startupDetails) { Write-Log "Precheck | $line" }
foreach ($line in $startupWarnings) { Write-Log "Precheck warning | $line" -Level WARN }
if (-not $startupCheck.Passed) {
    foreach ($line in $startupIssues) { Write-Log "Precheck failed | $line" -Level ERROR }
    $msg = @(
        "Startup pre-check failed:"
        ($startupIssues | ForEach-Object { " - $_" })
        ""
        "See log for details:"
        " - $script:LogPath"
    ) -join "`r`n"
    Write-Host $msg -ForegroundColor Red
    [System.Windows.MessageBox]::Show($msg, 'Prerequisite Check Failed', 'OK', 'Error') | Out-Null
    Write-Log '====== Script ended ======'
    exit 10
}

$instanceLock = Enter-CitrixDataPullSingleInstance
if (-not [bool]$instanceLock.Acquired) {
    if ([string]$instanceLock.FailureKind -eq 'Contended') {
        $instanceTitle = 'Citrix Data Pull Script Already Running'
        $instanceMessage = @(
            $instanceLock.ErrorMessage
            ''
            'Close the existing copy before starting another one.'
            'This protection applies only to this Windows server; it cannot detect copies running on other PVS Masters or DDCs.'
        ) -join "`r`n"
    } else {
        $instanceTitle = 'Single-Instance Check Failed'
        $instanceMessage = @(
            $instanceLock.ErrorMessage
            ''
            'No Citrix query was started. Review the daily log and the permissions for creating a machine-wide mutex.'
        ) -join "`r`n"
    }
    Write-Log "Single-instance check blocked launch: $($instanceLock.ErrorMessage)" -Level WARN
    Write-Host $instanceMessage -ForegroundColor Yellow
    [System.Windows.MessageBox]::Show($instanceMessage, $instanceTitle, 'OK', 'Warning') | Out-Null
    Write-Log '====== Script ended ======'
    exit 11
}

Write-Log 'Same-server single-instance lock acquired. Cross-controller concurrency remains an operational control.'

# Runtime flow:
# 1) Prefer PVS mode when StreamService is active.
# 2) Otherwise try Delivery Controller mode.
# 3) Exit with guidance if neither role is detected.
try {
    if (Test-PVSServer) {
        Write-Log 'Detected PVS Server'
        $roleCheck = Test-RolePrereqs -Role 'PVS'
        foreach ($line in $roleCheck.Details) { Write-Log "Role precheck (PVS) | $line" }
        foreach ($line in $roleCheck.Warnings) { Write-Log "Role precheck (PVS) warning | $line" -Level WARN }
        if (-not $roleCheck.Passed) {
            foreach ($line in $roleCheck.Issues) { Write-Log "Role precheck (PVS) failed | $line" -Level ERROR }
            $msg = @(
                "PVS prerequisite check failed:"
                ($roleCheck.Issues | ForEach-Object { " - $_" })
                ""
                "See log for details:"
                " - $script:LogPath"
            ) -join "`r`n"
            Write-Host $msg -ForegroundColor Red
            [System.Windows.MessageBox]::Show($msg, 'Prerequisite Check Failed', 'OK', 'Error') | Out-Null
            exit 10
        }
        Show-PVSWindow
    } elseif (Test-DeliveryController) {
        Write-Log 'Detected Delivery Controller'
        $roleCheck = Test-RolePrereqs -Role 'DDC'
        foreach ($line in $roleCheck.Details) { Write-Log "Role precheck (DDC) | $line" }
        foreach ($line in $roleCheck.Warnings) { Write-Log "Role precheck (DDC) warning | $line" -Level WARN }
        if (-not $roleCheck.Passed) {
            foreach ($line in $roleCheck.Issues) { Write-Log "Role precheck (DDC) failed | $line" -Level ERROR }
            $msg = @(
                "DDC prerequisite check failed:"
                ($roleCheck.Issues | ForEach-Object { " - $_" })
                ""
                "See log for details:"
                " - $script:LogPath"
            ) -join "`r`n"
            Write-Host $msg -ForegroundColor Red
            [System.Windows.MessageBox]::Show($msg, 'Prerequisite Check Failed', 'OK', 'Error') | Out-Null
            exit 10
        }
        Show-DDCWindow
    } else {
        $roleDetails = @($script:RoleDetectionDetails | ForEach-Object { " - $_" })
        $roleMessage = @(
            'No running Citrix PVS or XDC role was detected on this server.'
            ''
            $roleDetails
            ''
            'Run the script on a PVS Master with StreamService running or an XDC with CitrixBrokerService running.'
            "Log file: $script:LogPath"
        ) -join "`r`n"
        Write-Host $roleMessage -ForegroundColor Yellow
        Write-Log 'Neither PVS nor XDC role detected on this server; no Citrix query was started' -Level ERROR
        [System.Windows.MessageBox]::Show($roleMessage, 'Citrix Role Not Detected', 'OK', 'Error') | Out-Null
        exit 12
    }
} finally {
    Exit-CitrixDataPullSingleInstance -LockResult $instanceLock
    Write-Log 'Same-server single-instance lock cleanup completed'
    Write-Log '====== Script ended ======'
}

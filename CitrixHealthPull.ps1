<#
.SYNOPSIS
    Citrix PVS Device Personality & Delivery Controller Machine Data Retrieval Tool.

.DESCRIPTION
    Auto-detects whether the host is a Citrix PVS Server or Delivery Controller and
    presents a WPF GUI for retrieving and exporting device/machine information.

    Features:
    - Modern styled WPF interface with Citrix-inspired colour scheme
    - Parallel pings via RunspacePool for fast power-state checks (PVS mode)
    - Uses Generic List for O(1) collection building
    - CSV export
    - Select All / Deselect All, search/filter, record count, status bar
    - Structured file logging to %TEMP%

.NOTES
    Requires: Citrix PVS PowerShell SDK (PVS Server) or Citrix Broker PowerShell SDK (DDC)
    Platform: Windows Server 2012/2016/2019/2022 (pre-check validated at startup)
    Author:   Sachin
    Release:  2.0.0-prod
    ReleaseDate: 2026-02-18
    ChangeLog:
    - Added split Execution/Output tabs in PVS and DDC modes for better readability.
    - Added optional DHCP Check toggle; disabled by default.
    - Added reservation-level DHCP comparison across two PVS servers.
    - Option 66 validation is per-server (must match each PVS server IP), not cross-server equality.
    - Added single-server DHCP detail view with reservation/scope-effective options.
    - Added startup and role prerequisite checks (OS/PowerShell/WPF/Citrix cmdlets).
    - Added single-server mode summary suppression for PVS collection and DDC MC/DG aggregates.
#>

#Requires -Version 5.1

# ─────────────────────────────────────────────────────────────
# Region: Logging
# ─────────────────────────────────────────────────────────────
$script:LogPath = Join-Path $env:TEMP "PVS_XDC_Tool_$(Get-Date -Format 'yyyyMMdd').log"

<#
.SYNOPSIS
    Writes a timestamped entry to the daily log file.
#>
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:LogPath -Value $entry -ErrorAction SilentlyContinue
}


# ─────────────────────────────────────────────────────────────
# Region: Server Detection
# ─────────────────────────────────────────────────────────────
<#
.SYNOPSIS
    Detects if the local machine is a Citrix PVS server.
#>
function Test-PVSServer {
    try {
        $svc = Get-Service -Name 'StreamService' -ErrorAction SilentlyContinue
        if ($null -ne $svc -and $svc.Status -eq 'Running') {
            Write-Log "StreamService found and running — PVS server detected"
            return $true
        }
        Write-Log "StreamService not found or not running" -Level WARN
        return $false
    } catch {
        Write-Log "PVS detection failed: $_" -Level ERROR
        return $false
    }
}

<#
.SYNOPSIS
    Detects if the local machine is a Citrix Delivery Controller.
#>
function Test-DeliveryController {
    try {
        $svc = Get-Service -Name 'CitrixBrokerService' -ErrorAction SilentlyContinue
        if ($null -ne $svc -and $svc.Status -eq 'Running') {
            Write-Log "CitrixBrokerService found and running — DDC detected"
            return $true
        }
        Write-Log "CitrixBrokerService not found or not running" -Level WARN
        return $false
    } catch {
        Write-Log "DDC detection failed: $_" -Level ERROR
        return $false
    }
}

# ─────────────────────────────────────────────────────────────
# Region: Parallel Ping Helper (RunspacePool)
# ─────────────────────────────────────────────────────────────
<#
.SYNOPSIS
    Pings multiple devices in parallel and returns On/Off status.
.DESCRIPTION
    Uses a runspace pool to reduce wait time for large device sets.
    Returned hashtable keys are device names and values are 'On' or 'Off'.
#>
function Get-BulkPingResults {
    param(
        [string[]]$ComputerNames,
        [int]$ThrottleLimit = 30,
        [int]$TimeoutMs     = 1000
    )

    $results = @{}
    if (-not $ComputerNames -or $ComputerNames.Count -eq 0) { return $results }

    $pool = [RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit)
    $pool.Open()

    $scriptBlock = {
        param($Name, $Timeout)
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            $reply = $ping.Send($Name, $Timeout)
            if ($reply.Status -eq 'Success') { return 'On' } else { return 'Off' }
        } catch { return 'Off' }
    }

    $jobs = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($name in $ComputerNames) {
        $ps = [PowerShell]::Create().AddScript($scriptBlock).AddArgument($name).AddArgument($TimeoutMs)
        $ps.RunspacePool = $pool
        $jobs.Add([PSCustomObject]@{
            Pipe   = $ps
            Handle = $ps.BeginInvoke()
            Name   = $name
        })
    }

    foreach ($job in $jobs) {
        try {
            $results[$job.Name] = $job.Pipe.EndInvoke($job.Handle) | Select-Object -First 1
        } catch {
            $results[$job.Name] = 'Off'
        } finally {
            $job.Pipe.Dispose()
        }
    }

    $pool.Close()
    $pool.Dispose()
    return $results
}

# ─────────────────────────────────────────────────────────────
# Region: Shared WPF Styles (injected into every Window)
# ─────────────────────────────────────────────────────────────
# Colour palette — Citrix-inspired dark teal / blue-grey
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

# ─────────────────────────────────────────────────────────────
# Region: Shared UI Helpers
# ─────────────────────────────────────────────────────────────
function Export-DataGridToCSV {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [string]$DefaultFileName = 'Export.csv'
    )

    if ($null -eq $DataGrid.ItemsSource -or @($DataGrid.ItemsSource).Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'No data available to export.', 'Warning', 'OK', 'Warning')
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter   = 'CSV Files (*.csv)|*.csv'
    $dialog.Title    = 'Save CSV File'
    $dialog.FileName = $DefaultFileName

    if ($dialog.ShowDialog() -eq 'OK') {
        try {
            $DataGrid.ItemsSource | Export-Csv -Path $dialog.FileName -NoTypeInformation -Force
            [System.Windows.MessageBox]::Show(
                "Exported $(@($DataGrid.ItemsSource).Count) records to:`n$($dialog.FileName)",
                'Export Complete', 'OK', 'Information')
            Write-Log "Exported data to $($dialog.FileName)"
        } catch {
            [System.Windows.MessageBox]::Show(
                "Export failed: $_", 'Error', 'OK', 'Error')
            Write-Log "Export failed: $_" -Level ERROR
        }
    }
}

function Update-StatusBar {
    param(
        [System.Windows.Controls.TextBlock]$StatusText,
        [string]$Message
    )
    $StatusText.Dispatcher.Invoke([action]{
        $StatusText.Text = $Message
    })
}

function Invoke-UiRefresh {
    param(
        [System.Windows.Window]$Window
    )
    if ($null -eq $Window) { return }
    $Window.UpdateLayout()
    $Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Render)
}

<#
.SYNOPSIS
    Applies search filtering to a dataset with optional field-specific syntax.
.DESCRIPTION
    Supports:
    - Plain text: searches all visible fields
    - Field filter: "FieldName:Value" (example: PowerState:On)
#>
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

    $trimmed = $Term.Trim()
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
                $value -and $value.ToString() -like "*$fieldValue*"
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
            if ($prop.Value -and $prop.Value.ToString() -like "*$trimmed*") { $match = $true; break }
        }
        $match
    }

    return [PSCustomObject]@{
        Rows = @($globalFiltered)
        Summary = "$(@($globalFiltered).Count) / $($allRows.Count) records"
    }
}

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
        }
    }

    $maintenanceCount = @($machineList | Where-Object { $_.InMaintenanceMode -eq $true }).Count
    $unregisteredCount = @($machineList | Where-Object { [string]$_.RegistrationState -eq 'Unregistered' }).Count

    $logonDisabledCount = 0
    foreach ($machine in $machineList) {
        $isDisabled = $false
        if ($machine.PSObject.Properties.Match('WindowsConnectionSetting').Count -gt 0 -and $null -ne $machine.WindowsConnectionSetting) {
            $isDisabled = ([string]$machine.WindowsConnectionSetting -ne 'LogonEnabled')
        } else {
            $isDisabled = [bool]$machine.InMaintenanceMode
        }
        if ($isDisabled) { $logonDisabledCount++ }
    }

    return [PSCustomObject]@{
        MaintenanceCount = $maintenanceCount
        UnregisteredCount = $unregisteredCount
        LogonDisabledCount = $logonDisabledCount
    }
}

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

function ConvertTo-DhcpOptionValueText {
    param(
        [object]$Value
    )

    if ($null -eq $Value) { return 'N/A' }
    $values = @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($values.Count -eq 0) { return 'N/A' }
    return (($values | Sort-Object) -join ', ')
}

function Resolve-PvsServerPairForDhcp {
    $servers = [System.Collections.Generic.List[string]]::new()

    $addServer = {
        param([string]$Name)
        if ([string]::IsNullOrWhiteSpace($Name)) { return }
        $trimmed = $Name.Trim()
        if ($trimmed.Contains('.')) { $trimmed = $trimmed.Split('.')[0] }
        if (-not [string]::IsNullOrWhiteSpace($trimmed) -and -not $servers.Contains($trimmed)) {
            [void]$servers.Add($trimmed)
        }
    }

    & $addServer $env:COMPUTERNAME

    if (Get-Command Get-PvsServer -ErrorAction SilentlyContinue) {
        try {
            $pvsServers = @(Get-PvsServer -ErrorAction Stop)
            foreach ($srv in $pvsServers) {
                foreach ($propName in @('ServerName','Name','HostName','DNSName')) {
                    if ($srv.PSObject.Properties.Match($propName).Count -gt 0) {
                        & $addServer ([string]$srv.$propName)
                    }
                }
            }
        } catch {
            Write-Log "Unable to enumerate PVS servers for DHCP comparison: $($_.Exception.Message)" -Level WARN
        }
    }

    $primary = if ($servers.Count -ge 1) { $servers[0] } else { $null }
    $secondary = if ($servers.Count -ge 2) { $servers[1] } else { $null }

    return [PSCustomObject]@{
        Primary = $primary
        Secondary = $secondary
    }
}

function Get-DhcpSnapshot {
    param(
        [string]$ComputerName,
        [hashtable]$OptionMap
    )

    $snapshot = [PSCustomObject]@{
        ComputerName = $ComputerName
        ServiceState = 'Unavailable'
        ScopeCount = 'N/A'
        ActiveScopeCount = 'N/A'
        ReservationCount = 0
        Reservations = @{}
        Error = $null
    }

    try {
        $service = Get-Service -ComputerName $ComputerName -Name 'DHCPServer' -ErrorAction SilentlyContinue
        if ($null -eq $service) {
            $snapshot.ServiceState = 'Not Installed'
            return $snapshot
        }

        $snapshot.ServiceState = [string]$service.Status

        if (-not (Get-Command Get-DhcpServerv4Scope -ErrorAction SilentlyContinue) -or
            -not (Get-Command Get-DhcpServerv4OptionValue -ErrorAction SilentlyContinue) -or
            -not (Get-Command Get-DhcpServerv4Reservation -ErrorAction SilentlyContinue)) {
            $snapshot.Error = 'DHCP cmdlets unavailable'
            return $snapshot
        }

        $scopes = @(Get-DhcpServerv4Scope -ComputerName $ComputerName -ErrorAction Stop)
        $snapshot.ScopeCount = $scopes.Count
        $snapshot.ActiveScopeCount = @($scopes | Where-Object { [string]$_.State -eq 'Active' }).Count

        $serverOptionCache = @{}
        foreach ($optionId in ($OptionMap.Keys | Sort-Object)) {
            $valueText = 'N/A'
            try {
                $opt = Get-DhcpServerv4OptionValue -ComputerName $ComputerName -OptionId $optionId -ErrorAction Stop
                if ($opt -and $opt.Value) { $valueText = ConvertTo-DhcpOptionValueText -Value $opt.Value }
            } catch {
                # Keep as N/A when option is not set at server level.
            }
            $serverOptionCache[$optionId] = $valueText
        }

        $scopeOptionCache = @{}
        foreach ($scope in $scopes) {
            $scopeId = [string]$scope.ScopeId
            $scopeOptionCache[$scopeId] = @{}
            foreach ($optionId in ($OptionMap.Keys | Sort-Object)) {
                $valueText = 'N/A'
                try {
                    $opt = Get-DhcpServerv4OptionValue -ComputerName $ComputerName -ScopeId $scope.ScopeId -OptionId $optionId -ErrorAction Stop
                    if ($opt -and $opt.Value) { $valueText = ConvertTo-DhcpOptionValueText -Value $opt.Value }
                } catch {
                    # Keep as N/A when option is not set at this scope.
                }
                $scopeOptionCache[$scopeId][$optionId] = $valueText
            }

            $reservations = @()
            try {
                $reservations = @(Get-DhcpServerv4Reservation -ComputerName $ComputerName -ScopeId $scope.ScopeId -ErrorAction Stop)
            } catch {
                $reservations = @()
            }

            foreach ($reservation in $reservations) {
                $reservationIp = [string]$reservation.IPAddress
                if ([string]::IsNullOrWhiteSpace($reservationIp)) { continue }

                $optionValues = @{}
                foreach ($optionId in ($OptionMap.Keys | Sort-Object)) {
                    $scopeValue = if ($scopeOptionCache[$scopeId].ContainsKey($optionId)) {
                        [string]$scopeOptionCache[$scopeId][$optionId]
                    } else {
                        'N/A'
                    }
                    $optionValues[$optionId] = if ($scopeValue -ne 'N/A') { $scopeValue } elseif ($serverOptionCache.ContainsKey($optionId)) { [string]$serverOptionCache[$optionId] } else { 'N/A' }
                }

                $snapshot.Reservations[$reservationIp] = [PSCustomObject]@{
                    IPAddress = $reservationIp
                    Name = [string]$reservation.Name
                    ScopeId = $scopeId
                    OptionValues = $optionValues
                }
            }
        }

        $snapshot.ReservationCount = $snapshot.Reservations.Count
    } catch {
        $snapshot.Error = $_.Exception.Message
    }

    return $snapshot
}

function Get-PvsDhcpSummary {
    param(
        [string]$TargetServer
    )

    $optionMap = @{
        3  = 'Router'
        6  = 'DNS Servers'
        11 = 'Resource Location Servers'
        15 = 'DNS Domain Name'
        66 = 'Boot Server Host Name'
        67 = 'Bootfile Name'
    }

    function Convert-IPv4ToUInt32 {
        param([string]$IpAddress)
        try {
            $ipObj = [System.Net.IPAddress]::Parse($IpAddress)
            if ($ipObj.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { return $null }
            $bytes = $ipObj.GetAddressBytes()
            [array]::Reverse($bytes)
            return [BitConverter]::ToUInt32($bytes, 0)
        } catch {
            return $null
        }
    }

    function Resolve-IPv4Address {
        param([string]$Name)
        if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
        try {
            $resolved = [System.Net.Dns]::GetHostAddresses($Name) |
                Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
                Select-Object -First 1
            if ($resolved) { return $resolved.IPAddressToString }
        } catch { }
        return $null
    }

    function Get-PvsDhcpSingleServerDetail {
        param(
            [string]$ServerName,
            [string[]]$DhcpServers,
            [hashtable]$DhcpOptionMap
        )

        $targetIp = Resolve-IPv4Address -Name $ServerName
        $lines = [System.Collections.Generic.List[string]]::new()
        [void]$lines.Add("Target Server: $ServerName")
        [void]$lines.Add("Resolved IP: $(if ($targetIp) { $targetIp } else { 'N/A' })")

        foreach ($dhcpServer in @($DhcpServers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
            [void]$lines.Add("")
            [void]$lines.Add("${dhcpServer}:")

            try {
                $svc = Get-Service -ComputerName $dhcpServer -Name 'DHCPServer' -ErrorAction SilentlyContinue
                if ($null -eq $svc) {
                    [void]$lines.Add(" - DHCP service not installed")
                    continue
                }

                if (-not (Get-Command Get-DhcpServerv4Scope -ErrorAction SilentlyContinue)) {
                    [void]$lines.Add(" - DHCP scope cmdlets unavailable")
                    continue
                }

                $scopes = @(Get-DhcpServerv4Scope -ComputerName $dhcpServer -ErrorAction Stop)

                $matchedScope = $null
                if ($targetIp) {
                    $targetIpNum = Convert-IPv4ToUInt32 -IpAddress $targetIp
                    foreach ($scope in $scopes) {
                        $startNum = Convert-IPv4ToUInt32 -IpAddress ([string]$scope.StartRange)
                        $endNum = Convert-IPv4ToUInt32 -IpAddress ([string]$scope.EndRange)
                        if ($null -ne $targetIpNum -and $null -ne $startNum -and $null -ne $endNum -and
                            $targetIpNum -ge $startNum -and $targetIpNum -le $endNum) {
                            $matchedScope = $scope
                            break
                        }
                    }
                }

                if ($matchedScope) {
                    [void]$lines.Add(" - Matched Scope: $([string]$matchedScope.ScopeId)")
                } else {
                    [void]$lines.Add(" - Matched Scope: Not Found")
                }

                $matchedReservation = $null
                $reservationText = 'Not Found'
                if ($matchedScope -and (Get-Command Get-DhcpServerv4Reservation -ErrorAction SilentlyContinue)) {
                    try {
                        $reservations = @(Get-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $matchedScope.ScopeId -ErrorAction Stop)
                        if ($targetIp) {
                            $matchedReservation = $reservations | Where-Object { [string]$_.IPAddress -eq $targetIp } | Select-Object -First 1
                        }
                        if (-not $matchedReservation) {
                            $matchedReservation = $reservations | Where-Object { [string]$_.Name -ieq $ServerName } | Select-Object -First 1
                        }
                        if ($matchedReservation) {
                            $reservationText = "Found: IP=$([string]$matchedReservation.IPAddress), Name=$([string]$matchedReservation.Name)"
                        }
                    } catch {
                        $reservationText = "Read failed: $($_.Exception.Message)"
                    }
                }
                [void]$lines.Add(" - Reservation: $reservationText")

                if (Get-Command Get-DhcpServerv4OptionValue -ErrorAction SilentlyContinue) {
                    if ($matchedReservation -and $matchedScope) {
                        $serverOptionCache = @{}
                        foreach ($optId in ($DhcpOptionMap.Keys | Sort-Object)) {
                            $sv = 'N/A'
                            try {
                                $svOpt = Get-DhcpServerv4OptionValue -ComputerName $dhcpServer -OptionId $optId -ErrorAction Stop
                                if ($svOpt -and $svOpt.Value) { $sv = ConvertTo-DhcpOptionValueText -Value $svOpt.Value }
                            } catch { }
                            $serverOptionCache[$optId] = $sv
                        }

                        [void]$lines.Add(" - Reservation Options:")
                        foreach ($optId in ($DhcpOptionMap.Keys | Sort-Object)) {
                            $optLabel = "{0:D3} {1}" -f $optId, $DhcpOptionMap[$optId]
                            $sc = 'N/A'
                            try {
                                $scOpt = Get-DhcpServerv4OptionValue -ComputerName $dhcpServer -ScopeId $matchedScope.ScopeId -OptionId $optId -ErrorAction Stop
                                if ($scOpt -and $scOpt.Value) { $sc = ConvertTo-DhcpOptionValueText -Value $scOpt.Value }
                            } catch { }
                            $effectiveValue = if ($sc -ne 'N/A') { $sc } elseif ($serverOptionCache.ContainsKey($optId)) { [string]$serverOptionCache[$optId] } else { 'N/A' }
                            [void]$lines.Add("   - ${optLabel}: $effectiveValue")
                        }
                    } else {
                        [void]$lines.Add(" - Reservation Options: N/A (reservation not found)")
                    }
                } else {
                    [void]$lines.Add(" - DHCP option cmdlets unavailable")
                }
            } catch {
                [void]$lines.Add(" - DHCP read failed: $($_.Exception.Message)")
            }
        }

        return ($lines -join "`r`n")
    }

    $pair = Resolve-PvsServerPairForDhcp
    if ([string]::IsNullOrWhiteSpace($pair.Primary)) {
        return [PSCustomObject]@{
            ServiceState = 'Unavailable'
            ScopeCount = 'N/A'
            ActiveScopeCount = 'N/A'
            Detail = 'DHCP check unavailable: no PVS server resolved.'
            ConsistencyDetail = 'DHCP comparison skipped.'
            SingleServerDetail = ''
            HasMismatch = $false
        }
    }

    $primarySnapshot = Get-DhcpSnapshot -ComputerName $pair.Primary -OptionMap $optionMap
    $detail = "$($pair.Primary): Service=$($primarySnapshot.ServiceState), Scopes=$($primarySnapshot.ScopeCount), Active=$($primarySnapshot.ActiveScopeCount), Reservations=$($primarySnapshot.ReservationCount)"

    # If we only have one PVS server, still return local DHCP status.
    if ([string]::IsNullOrWhiteSpace($pair.Secondary)) {
        if ($primarySnapshot.Error) {
            $detail = "$detail, Error=$($primarySnapshot.Error)"
        }
        return [PSCustomObject]@{
            ServiceState = $primarySnapshot.ServiceState
            ScopeCount = $primarySnapshot.ScopeCount
            ActiveScopeCount = $primarySnapshot.ActiveScopeCount
            Detail = $detail
            ConsistencyDetail = 'DHCP comparison skipped: only one PVS server found.'
            SingleServerDetail = if (-not [string]::IsNullOrWhiteSpace($TargetServer)) {
                Get-PvsDhcpSingleServerDetail -ServerName $TargetServer -DhcpServers @($pair.Primary) -DhcpOptionMap $optionMap
            } else {
                ''
            }
            HasMismatch = $false
        }
    }

    $secondarySnapshot = Get-DhcpSnapshot -ComputerName $pair.Secondary -OptionMap $optionMap
    $detail = "$detail || $($pair.Secondary): Service=$($secondarySnapshot.ServiceState), Scopes=$($secondarySnapshot.ScopeCount), Active=$($secondarySnapshot.ActiveScopeCount), Reservations=$($secondarySnapshot.ReservationCount)"
    if ($primarySnapshot.Error) { $detail += " | $($pair.Primary) Error=$($primarySnapshot.Error)" }
    if ($secondarySnapshot.Error) { $detail += " | $($pair.Secondary) Error=$($secondarySnapshot.Error)" }

    $mismatchLines = [System.Collections.Generic.List[string]]::new()
    $hasMismatch = $false

    $allReservationKeys = @(
        @($primarySnapshot.Reservations.Keys) +
        @($secondarySnapshot.Reservations.Keys)
    ) | Select-Object -Unique | Sort-Object

    $primaryIp = Resolve-IPv4Address -Name $pair.Primary
    $secondaryIp = Resolve-IPv4Address -Name $pair.Secondary

    foreach ($reservationKey in $allReservationKeys) {
        $leftReservation = if ($primarySnapshot.Reservations.ContainsKey($reservationKey)) { $primarySnapshot.Reservations[$reservationKey] } else { $null }
        $rightReservation = if ($secondarySnapshot.Reservations.ContainsKey($reservationKey)) { $secondarySnapshot.Reservations[$reservationKey] } else { $null }

        if ($null -eq $leftReservation) {
            $hasMismatch = $true
            [void]$mismatchLines.Add("Reservation $reservationKey missing on $($pair.Primary)")
            continue
        }
        if ($null -eq $rightReservation) {
            $hasMismatch = $true
            [void]$mismatchLines.Add("Reservation $reservationKey missing on $($pair.Secondary)")
            continue
        }

        if ([string]$leftReservation.Name -ne [string]$rightReservation.Name) {
            $hasMismatch = $true
            [void]$mismatchLines.Add("Reservation $reservationKey name mismatch: $($pair.Primary)='$([string]$leftReservation.Name)' | $($pair.Secondary)='$([string]$rightReservation.Name)'")
        }

        foreach ($optionId in ($optionMap.Keys | Sort-Object)) {
            $optionName = if ($optionMap.ContainsKey($optionId)) { $optionMap[$optionId] } else { "Option $optionId" }
            $leftVal = if ($leftReservation.OptionValues.ContainsKey($optionId)) { [string]$leftReservation.OptionValues[$optionId] } else { 'N/A' }
            $rightVal = if ($rightReservation.OptionValues.ContainsKey($optionId)) { [string]$rightReservation.OptionValues[$optionId] } else { 'N/A' }

            if ($optionId -eq 66) {
                $expectedLeft = if ([string]::IsNullOrWhiteSpace($primaryIp)) { 'N/A' } else { $primaryIp }
                $expectedRight = if ([string]::IsNullOrWhiteSpace($secondaryIp)) { 'N/A' } else { $secondaryIp }

                if ($leftVal -ne $expectedLeft) {
                    $hasMismatch = $true
                    [void]$mismatchLines.Add("Reservation $reservationKey option 66 ($optionName): $($pair.Primary) expected '$expectedLeft' but found '$leftVal'")
                }
                if ($rightVal -ne $expectedRight) {
                    $hasMismatch = $true
                    [void]$mismatchLines.Add("Reservation $reservationKey option 66 ($optionName): $($pair.Secondary) expected '$expectedRight' but found '$rightVal'")
                }
                continue
            }

            if ($leftVal -ne $rightVal) {
                $hasMismatch = $true
                [void]$mismatchLines.Add("Reservation $reservationKey option $optionId ($optionName): $($pair.Primary)='$leftVal' | $($pair.Secondary)='$rightVal'")
            }
        }
    }

    $consistencyDetail = if ($mismatchLines.Count -gt 0) {
        "Mismatch detected`r`n - " + ($mismatchLines -join "`r`n - ")
    } elseif ($allReservationKeys.Count -eq 0) {
        "DHCP comparison unavailable: no reservations read from either server."
    } else {
        "All reservation-level DHCP values match between $($pair.Primary) and $($pair.Secondary) (options 3, 6, 11, 15, 67 cross-server; option 66 validated against each server IP)."
    }

    $singleServerDetail = ''
    if (-not [string]::IsNullOrWhiteSpace($TargetServer)) {
        $singleServerDetail = Get-PvsDhcpSingleServerDetail -ServerName $TargetServer -DhcpServers @($pair.Primary, $pair.Secondary) -DhcpOptionMap $optionMap
    }

    return [PSCustomObject]@{
        ServiceState = "$($primarySnapshot.ServiceState)/$($secondarySnapshot.ServiceState)"
        ScopeCount = "$($primarySnapshot.ScopeCount)/$($secondarySnapshot.ScopeCount)"
        ActiveScopeCount = "$($primarySnapshot.ActiveScopeCount)/$($secondarySnapshot.ActiveScopeCount)"
        Detail = $detail
        ConsistencyDetail = $consistencyDetail
        SingleServerDetail = $singleServerDetail
        HasMismatch = $hasMismatch
    }
}

function Format-GroupDetailForPanel {
    param(
        [string]$Detail
    )
    if (-not $Detail) { return '' }
    return ($Detail -replace ': ', ":`r`n - " -replace ' \| ', "`r`n - ")
}

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

# ─────────────────────────────────────────────────────────────
# Region: PVS Mode
# ─────────────────────────────────────────────────────────────
<#
.SYNOPSIS
    Displays the PVS window and handles device retrieval/export.
#>
function Show-PVSWindow {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName 'System.Windows.Forms'

    $XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Citrix PVS — Device Personality Retrieval" Height="640" Width="940"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#ECF0F1" MinWidth="760" MinHeight="500">
  $($script:SharedStyles)
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="{StaticResource HeaderBgBrush}" Padding="20,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0">
          <TextBlock Text="PVS Device Personality" FontSize="18" FontWeight="Bold"
                     Foreground="{StaticResource HeaderFgBrush}"/>
          <TextBlock Text="Retrieve device data from Citrix Provisioning Services" FontSize="11"
                     Foreground="#AEB6BF" Margin="0,3,0,0"/>
        </StackPanel>
        <TextBlock Grid.Column="1" Name="lblServer" FontSize="11" Foreground="#AEB6BF"
                   VerticalAlignment="Center"/>
      </Grid>
    </Border>

    <TabControl Name="tabMain" Grid.Row="1" Margin="12,10,12,10">
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
              <DockPanel Margin="0,0,0,8">
                <TextBlock Text="Data Fields" FontSize="13" FontWeight="SemiBold"
                           Foreground="#2C3E50" VerticalAlignment="Center"/>
                <Button Name="btnToggle" Content="Select All" Style="{StaticResource LinkButton}"
                        HorizontalAlignment="Right" VerticalAlignment="Center"/>
              </DockPanel>
              <WrapPanel Orientation="Horizontal">
                <CheckBox Name="chkDevices" Content="Device Name"  IsChecked="True"  Margin="0,0,18,6"/>
                <CheckBox Name="chkVDisk"   Content="vDisk"        IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkReboot"  Content="Reboot Day"   IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkCSR"     Content="CSR Server"   IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkXDC"     Content="XDC Server"   IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkPower"   Content="Power State"  IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkServerIP" Content="Server IP"   IsChecked="False" Margin="0,0,18,6"/>
                <CheckBox Name="chkDhcpCheck" Content="DHCP Check" IsChecked="False" Margin="0,0,18,6"/>
              </WrapPanel>
              <WrapPanel Orientation="Horizontal" Margin="0,4,0,0" VerticalAlignment="Center">
                <CheckBox Name="chkSingleServer" Content="Single Server Only" IsChecked="False"
                          VerticalAlignment="Center" Margin="0,0,12,0"/>
                <TextBlock Text="Server Name:" FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <TextBox Name="txtServerName" Width="260" Style="{StaticResource StyledTextBox}"
                         IsEnabled="False" VerticalAlignment="Center" Margin="0,0,12,0"/>
              </WrapPanel>
            </StackPanel>
          </Border>

          <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="4,12,4,0">
            <Button Name="btnRetrieve" Content="Retrieve Data"  Style="{StaticResource PrimaryButton}"   Width="140" Height="34" Margin="0,0,10,0"/>
            <Button Name="btnExport"   Content="Export to CSV"   Style="{StaticResource SecondaryButton}" Width="140" Height="34" Margin="0,0,10,0"/>
            <Button Name="btnClose"    Content="Close"           Style="{StaticResource CloseButton}"     Width="100" Height="34"/>
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
              <TextBlock Text="Tip: Use Field:Value for column filters (example: PowerState:On)" FontSize="10"
                         Foreground="#7F8C8D" Margin="0,4,0,0"/>
            </StackPanel>
          </Border>

          <ProgressBar Name="progressBar" Grid.Row="3" Style="{StaticResource AccentProgress}"
                       Margin="4,10,4,0" Visibility="Hidden"/>
        </Grid>
      </TabItem>

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
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBox Name="txtCollectionSummary" Grid.Column="0" IsReadOnly="True" BorderThickness="0" Background="Transparent"
                       TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontSize="11"
                       Foreground="#2C3E50"/>
              <TextBox Name="txtDhcpSummary" Grid.Column="2" IsReadOnly="True" BorderThickness="0" Background="Transparent"
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
              <TextBlock Name="lblRecordCount" Text="" FontSize="11"
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

    $window = [Windows.Markup.XamlReader]::Parse($XAML)

    # Bind controls
    $tabMain       = $window.FindName('tabMain')
    $chkDevices    = $window.FindName('chkDevices')
    $chkVDisk      = $window.FindName('chkVDisk')
    $chkReboot     = $window.FindName('chkReboot')
    $chkCSR        = $window.FindName('chkCSR')
    $chkXDC        = $window.FindName('chkXDC')
    $chkPower      = $window.FindName('chkPower')
    $chkServerIP   = $window.FindName('chkServerIP')
    $chkSingleServer = $window.FindName('chkSingleServer')
    $chkDhcpCheck  = $window.FindName('chkDhcpCheck')
    $txtServerName = $window.FindName('txtServerName')
    $btnToggle     = $window.FindName('btnToggle')
    $btnRetrieve   = $window.FindName('btnRetrieve')
    $btnExport     = $window.FindName('btnExport')
    $btnClose      = $window.FindName('btnClose')
    $txtSearch     = $window.FindName('txtSearch')
    $dataGrid      = $window.FindName('dataGrid')
    $progressBar   = $window.FindName('progressBar')
    $txtCollectionSummary = $window.FindName('txtCollectionSummary')
    $txtDhcpSummary = $window.FindName('txtDhcpSummary')
    $statusText    = $window.FindName('statusText')
    $lblRecordCount = $window.FindName('lblRecordCount')
    $lblServer     = $window.FindName('lblServer')

    Set-DataGridDisplayFormat -DataGrid $dataGrid -Window $window

    $lblServer.Text = "$($env:COMPUTERNAME)"

    # Store full dataset for filtering
    $script:pvsFullData = $null

    # Toggle all checkboxes
    $btnToggle.Add_Click({
        $allChecks = @($chkDevices, $chkVDisk, $chkReboot, $chkCSR, $chkXDC, $chkPower, $chkServerIP)
        $anyUnchecked = $allChecks | Where-Object { -not $_.IsChecked }
        $newState = [bool]$anyUnchecked
        foreach ($cb in $allChecks) { $cb.IsChecked = $newState }
        $btnToggle.Content = if ($newState) { 'Deselect All' } else { 'Select All' }
    })

    $chkSingleServer.Add_Click({
        $txtServerName.IsEnabled = [bool]$chkSingleServer.IsChecked
        if (-not $chkSingleServer.IsChecked) {
            $txtServerName.Text = ''
        } else {
            $txtServerName.Focus()
        }
    })

    # Search / filter
    $txtSearch.Add_TextChanged({
        if ($null -eq $script:pvsFullData) { return }
        $filterResult = Apply-GridFilter -Data $script:pvsFullData -Term $txtSearch.Text
        $dataGrid.ItemsSource = $filterResult.Rows
        $lblRecordCount.Text = $filterResult.Summary
    })

    # Retrieve PVS Data
    $btnRetrieve.Add_Click({
        $originalRetrieveText = [string]$btnRetrieve.Content
        $progressBar.Visibility = 'Visible'
        $progressBar.Value = 0
        $dataGrid.ItemsSource = $null
        $script:pvsFullData = $null
        $txtSearch.Text = ''
        $txtCollectionSummary.Text = ''
        $txtDhcpSummary.Text = ''
        $statusText.Text = 'Retrieving data. Please wait...'
        $lblRecordCount.Text = ''
        $btnRetrieve.IsEnabled = $false
        $btnRetrieve.Content = 'Retrieving...'
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Invoke-UiRefresh -Window $window

        $wantVDisk  = $chkVDisk.IsChecked
        $wantReboot = $chkReboot.IsChecked
        $wantCSR    = $chkCSR.IsChecked
        $wantXDC    = $chkXDC.IsChecked
        $wantPower  = $chkPower.IsChecked
        $wantServerIP = $chkServerIP.IsChecked
        $wantDhcpCheck = [bool]$chkDhcpCheck.IsChecked
        $singleServerMode = [bool]$chkSingleServer.IsChecked
        $targetServer = $txtServerName.Text.Trim()

        if ($singleServerMode -and -not $targetServer) {
            [System.Windows.MessageBox]::Show(
                'Please enter a server name for single-server mode.',
                'Validation', 'OK', 'Warning')
            $progressBar.Visibility = 'Hidden'
            $btnRetrieve.IsEnabled = $true
            $btnRetrieve.Content = $originalRetrieveText
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $statusText.Text = 'Validation failed'
            return
        }
        $tabMain.SelectedIndex = 1

        # Log selected fields
        $selectedFields = @('DeviceName','Collection')
        if ($wantVDisk)  { $selectedFields += 'vDisk' }
        if ($wantPower)  { $selectedFields += 'PowerState' }
        if ($wantReboot) { $selectedFields += 'RebootDay' }
        if ($wantCSR)    { $selectedFields += 'CSRServer' }
        if ($wantXDC)    { $selectedFields += 'XDCServer' }
        if ($wantServerIP) { $selectedFields += 'ServerIP' }
        if ($wantDhcpCheck) { $selectedFields += 'DHCPCheck' }
        Write-Log "PVS retrieval requested — fields: $($selectedFields -join ', ')"
        if ($singleServerMode) { Write-Log "PVS retrieval scope: single server [$targetServer]" }

        try {
            if (-not (Get-PSSnapin -Name 'Citrix*' -ErrorAction SilentlyContinue)) {
                Write-Log 'Loading Citrix PVS snap-in...'
                Add-PSSnapin Citrix* -ErrorAction Stop
                Write-Log 'Citrix PVS snap-in loaded successfully'
            }

            $statusText.Text = 'Retrieving PVS devices...'
            Write-Log 'Retrieving PVS devices'

            $devices = Get-PVSDevice -ErrorAction Stop
            if ($singleServerMode) {
                $devices = @($devices | Where-Object {
                    $name = [string]$_.Name
                    $shortName = $name -replace '^.*\\', ''
                    $name -ieq $targetServer -or $shortName -ieq $targetServer
                })
            }
            if (-not $devices) {
                if ($singleServerMode) {
                    Write-Log "PVS device not found for server [$targetServer]" -Level WARN
                    [System.Windows.MessageBox]::Show(
                        "Server '$targetServer' was not found in PVS device list.",
                        'Information', 'OK', 'Information')
                } else {
                    Write-Log 'No PVS devices found on this server' -Level WARN
                    [System.Windows.MessageBox]::Show('No PVS devices found.', 'Information', 'OK', 'Information')
                }
                $progressBar.Visibility = 'Hidden'
                $statusText.Text = 'No devices found'
                $btnRetrieve.IsEnabled = $true
                return
            }

            $totalDevices = @($devices).Count
            Write-Log "Found $totalDevices PVS devices"
            $statusText.Text = "Processing $totalDevices devices..."

            $pingResults = @{}
            if ($wantPower) {
                Write-Log "Starting parallel ping for $totalDevices devices..."
                $statusText.Text = "Pinging $totalDevices devices in parallel..."
                $pingResults = Get-BulkPingResults -ComputerNames ($devices | ForEach-Object { $_.Name })
                $onCount  = @($pingResults.Values | Where-Object { $_ -eq 'On' }).Count
                $offCount = $totalDevices - $onCount
                Write-Log "Ping complete: $onCount online, $offCount offline"
            }

            $deviceInfo = [System.Collections.Generic.List[object]]::new($totalDevices)
            $counter = 0
            $dnsCache = @{}

            foreach ($device in $devices) {
                $counter++
                $serverName       = $device.Name
                $deviceCollection = $device.CollectionName

                $vDiskValue = $rebootValue = $csrValue = $xdcValue = $serverIpValue = 'N/A'

                # Retrieve assigned vDisk name
                if ($wantVDisk) {
                    try {
                        $diskInfo = @(Get-PVSDiskInfo -DeviceName $serverName -ErrorAction Stop)
                        if ($diskInfo) {
                            $vDiskValue = ($diskInfo | ForEach-Object { $_.Name }) -join ', '
                        }
                    } catch {
                        Write-Log "vDisk fetch failed for ${serverName}: $_" -Level WARN
                    }
                }

                if ($wantReboot -or $wantCSR -or $wantXDC) {
                    try {
                        # DevicePersonality stores key/value pairs such as Reboot, CSAServer, XDC_LIST.
                        $personality = @(Get-PVSDevicePersonality -DeviceName $serverName -ErrorAction Stop |
                                       Select-Object -ExpandProperty DevicePersonality)
                    } catch {
                        $personality = @()
                        Write-Log "Personality fetch failed for ${serverName}: $_" -Level WARN
                    }

                    if ($personality) {
                        if ($wantReboot) {
                            $val = ($personality | Where-Object { $_.Name -eq 'Reboot' }).Value
                            if ($val) { $rebootValue = $val -join ', ' }
                        }
                        if ($wantCSR) {
                            # CSR Server is stored in personality key 'CSAServer' in many PVS environments.
                            $val = ($personality | Where-Object { $_.Name -eq 'CSAServer' }).Value
                            if ($val) { $csrValue = $val -join ', ' }
                        }
                        if ($wantXDC) {
                            # XDC list is stored in personality key 'XDC_LIST'.
                            $val = ($personality | Where-Object { $_.Name -eq 'XDC_LIST' }).Value
                            if ($val) { $xdcValue = $val -join ', ' }
                        }
                    }
                }

                if ($wantServerIP) {
                    $serverIp = $null
                    $cacheKey = $serverName.ToLowerInvariant()

                    foreach ($propName in @('IPAddress','IpAddress','IP','DeviceIP','DeviceIPAddress')) {
                        if ($device.PSObject.Properties.Match($propName).Count -gt 0 -and $device.$propName) {
                            $rawValue = $device.$propName
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
                            try {
                                $resolved = [System.Net.Dns]::GetHostAddresses($serverName) |
                                    Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
                                    Select-Object -First 1
                                if ($resolved) { $serverIp = $resolved.IPAddressToString }
                            } catch {
                                # Keep non-terminating; unresolved names remain N/A.
                            }
                            $dnsCache[$cacheKey] = $serverIp
                        }
                    }

                    if ($serverIp) { $serverIpValue = [string]$serverIp }
                }

                $obj = [ordered]@{
                    SrNumber         = $counter
                    ServerName       = $serverName
                    DeviceCollection = $deviceCollection
                }
                if ($wantVDisk)  { $obj['vDisk']       = $vDiskValue }
                if ($wantPower)  { $obj['PowerState']  = if ($pingResults[$serverName]) { $pingResults[$serverName] } else { 'Unknown' } }
                if ($wantReboot) { $obj['RebootDay']   = $rebootValue }
                if ($wantCSR)    { $obj['CSRServer']   = $csrValue }
                if ($wantXDC)    { $obj['XDCServer']   = $xdcValue }
                if ($wantServerIP) { $obj['ServerIP']  = $serverIpValue }

                [void]$deviceInfo.Add([PSCustomObject]$obj)
                $progressBar.Value = ($counter / $totalDevices) * 100
            }

            $script:pvsFullData = $deviceInfo
            $dataGrid.ItemsSource = $deviceInfo
            $lblRecordCount.Text = "$totalDevices records"
            $statusText.Text = "Retrieved $totalDevices devices"
            if ($singleServerMode) {
                $txtCollectionSummary.Text = "Single Server Mode:`r`n - Collection aggregation skipped for [$targetServer]"
            } else {
                $pvsCollectionSummary = Get-GroupedCountSummary -Items $devices -PropertyName 'CollectionName' -Label 'Collections'
                $txtCollectionSummary.Text = Format-CountSection -Heading 'Collections' -Description 'No. of servers in each Device Collection:' -Detail $pvsCollectionSummary.Detail
                Write-Log "PVS grouped counts | $($pvsCollectionSummary.Detail)"
            }
            if ($wantDhcpCheck) {
                $dhcpSummary = Get-PvsDhcpSummary -TargetServer $(if ($singleServerMode) { $targetServer } else { $null })
                if ($singleServerMode -and -not [string]::IsNullOrWhiteSpace($dhcpSummary.SingleServerDetail)) {
                    $txtDhcpSummary.Text = @(
                        "DHCP Check:"
                        " - $($dhcpSummary.Detail)"
                        ""
                        "DHCP Consistency:"
                        " - $($dhcpSummary.ConsistencyDetail)"
                        ""
                        "DHCP Data (Single Server):"
                        $dhcpSummary.SingleServerDetail
                    ) -join "`r`n"
                } else {
                    $txtDhcpSummary.Text = @(
                        "DHCP Check:"
                        " - $($dhcpSummary.Detail)"
                        ""
                        "DHCP Consistency:"
                        " - $($dhcpSummary.ConsistencyDetail)"
                    ) -join "`r`n"
                }

                Write-Log "PVS DHCP summary | $($dhcpSummary.Detail)"
                Write-Log "PVS DHCP consistency | $($dhcpSummary.ConsistencyDetail)"
                if ($singleServerMode -and -not [string]::IsNullOrWhiteSpace($dhcpSummary.SingleServerDetail)) {
                    Write-Log "PVS DHCP single-server detail generated for [$targetServer]"
                }
            } else {
                $txtDhcpSummary.Text = @(
                    "DHCP Check:"
                    " - Skipped (checkbox not selected)"
                ) -join "`r`n"
                Write-Log "PVS DHCP check skipped (checkbox not selected)"
            }
            Write-Log "PVS retrieval complete: $totalDevices devices"
            if ($singleServerMode) {
                Write-Log "PVS grouped counts skipped (single-server mode: [$targetServer])"
            }

        } catch {
            [System.Windows.MessageBox]::Show("Error: $_", 'Error', 'OK', 'Error')
            Write-Log "PVS retrieval error: $_" -Level ERROR
            $statusText.Text = 'Error during retrieval'
        } finally {
            $progressBar.Visibility = 'Hidden'
            $btnRetrieve.IsEnabled = $true
            $btnRetrieve.Content = $originalRetrieveText
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })

    $btnExport.Add_Click({ Export-DataGridToCSV -DataGrid $dataGrid -DefaultFileName 'PVS_Devices.csv' })
    $btnClose.Add_Click({ $window.Close() })

    $window.ShowDialog() | Out-Null
}

# ─────────────────────────────────────────────────────────────
# Region: DDC Mode
# ─────────────────────────────────────────────────────────────
<#
.SYNOPSIS
    Displays the DDC window and handles broker machine retrieval/export.
#>
function Show-DDCWindow {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName 'System.Windows.Forms'

    $XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Citrix DDC — VDA Status Retrieval" Height="640" Width="940"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
        Background="#ECF0F1" MinWidth="760" MinHeight="500">
  $($script:SharedStyles)
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="{StaticResource HeaderBgBrush}" Padding="20,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0">
          <TextBlock Text="VDA Status Retrieval" FontSize="18" FontWeight="Bold"
                     Foreground="{StaticResource HeaderFgBrush}"/>
          <TextBlock Text="Retrieve broker machine data from Citrix Delivery Controller" FontSize="11"
                     Foreground="#AEB6BF" Margin="0,3,0,0"/>
        </StackPanel>
        <TextBlock Grid.Column="1" Name="lblServer" FontSize="11" Foreground="#AEB6BF"
                   VerticalAlignment="Center"/>
      </Grid>
    </Border>

    <TabControl Name="tabMain" Grid.Row="1" Margin="12,10,12,10">
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
              <DockPanel Margin="0,0,0,8">
                <TextBlock Text="Data Fields" FontSize="13" FontWeight="SemiBold"
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
              <WrapPanel Orientation="Horizontal" Margin="0,4,0,0" VerticalAlignment="Center">
                <CheckBox Name="chkSingleServer" Content="Single Server Only" IsChecked="False"
                          VerticalAlignment="Center" Margin="0,0,12,0"/>
                <TextBlock Text="Server Name:" FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <TextBox Name="txtServerName" Width="260" Style="{StaticResource StyledTextBox}"
                         IsEnabled="False" VerticalAlignment="Center" Margin="0,0,12,0"/>
              </WrapPanel>
            </StackPanel>
          </Border>

          <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center" Margin="4,12,4,0">
            <Button Name="btnRetrieve" Content="Retrieve Data"  Style="{StaticResource PrimaryButton}"   Width="140" Height="34" Margin="0,0,10,0"/>
            <Button Name="btnExport"   Content="Export to CSV"   Style="{StaticResource SecondaryButton}" Width="140" Height="34" Margin="0,0,10,0"/>
            <Button Name="btnClose"    Content="Close"           Style="{StaticResource CloseButton}"     Width="100" Height="34"/>
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
              <TextBlock Text="Tip: Use Field:Value for column filters (example: RegistrationState:Registered)" FontSize="10"
                         Foreground="#7F8C8D" Margin="0,4,0,0"/>
            </StackPanel>
          </Border>

          <ProgressBar Name="progressBar" Grid.Row="3" Style="{StaticResource AccentProgress}"
                       Margin="4,10,4,0" Visibility="Hidden"/>
        </Grid>
      </TabItem>

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
              <TextBlock Name="lblRecordCount" Text="" FontSize="11"
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

    $window = [Windows.Markup.XamlReader]::Parse($XAML)

    $tabMain              = $window.FindName('tabMain')
    $chkMachineName       = $window.FindName('chkMachineName')
    $chkMachineCatalog    = $window.FindName('chkMachineCatalog')
    $chkDeliveryGroup     = $window.FindName('chkDeliveryGroup')
    $chkRegistrationState = $window.FindName('chkRegistrationState')
    $chkInMaintenanceMode = $window.FindName('chkInMaintenanceMode')
    $chkLogonMode         = $window.FindName('chkLogonMode')
    $chkServerIP          = $window.FindName('chkServerIP')
    $chkSingleServer      = $window.FindName('chkSingleServer')
    $txtServerName        = $window.FindName('txtServerName')
    $btnToggle     = $window.FindName('btnToggle')
    $btnRetrieve   = $window.FindName('btnRetrieve')
    $btnExport     = $window.FindName('btnExport')
    $btnClose      = $window.FindName('btnClose')
    $txtSearch     = $window.FindName('txtSearch')
    $dataGrid      = $window.FindName('dataGrid')
    $progressBar   = $window.FindName('progressBar')
    $txtCatalogSummary = $window.FindName('txtCatalogSummary')
    $txtDeliverySummary = $window.FindName('txtDeliverySummary')
    $statusText    = $window.FindName('statusText')
    $lblRecordCount = $window.FindName('lblRecordCount')
    $lblServer     = $window.FindName('lblServer')

    Set-DataGridDisplayFormat -DataGrid $dataGrid -Window $window

    $lblServer.Text = "$($env:COMPUTERNAME)"

    $script:ddcFullData = $null

    $btnToggle.Add_Click({
        $allChecks = @($chkMachineName, $chkMachineCatalog, $chkDeliveryGroup,
                       $chkRegistrationState, $chkInMaintenanceMode, $chkLogonMode, $chkServerIP)
        $anyUnchecked = $allChecks | Where-Object { -not $_.IsChecked }
        $newState = [bool]$anyUnchecked
        foreach ($cb in $allChecks) { $cb.IsChecked = $newState }
        $btnToggle.Content = if ($newState) { 'Deselect All' } else { 'Select All' }
    })

    $chkSingleServer.Add_Click({
        $txtServerName.IsEnabled = [bool]$chkSingleServer.IsChecked
        if (-not $chkSingleServer.IsChecked) {
            $txtServerName.Text = ''
        } else {
            $txtServerName.Focus()
        }
    })

    # Search / filter
    $txtSearch.Add_TextChanged({
        if ($null -eq $script:ddcFullData) { return }
        $filterResult = Apply-GridFilter -Data $script:ddcFullData -Term $txtSearch.Text
        $dataGrid.ItemsSource = $filterResult.Rows
        $lblRecordCount.Text = $filterResult.Summary
    })

    $btnRetrieve.Add_Click({
        $originalRetrieveText = [string]$btnRetrieve.Content
        $progressBar.Visibility = 'Visible'
        $progressBar.Value = 0
        $dataGrid.ItemsSource = $null
        $script:ddcFullData = $null
        $txtSearch.Text = ''
        $txtCatalogSummary.Text = ''
        $txtDeliverySummary.Text = ''
        $statusText.Text = 'Retrieving data. Please wait...'
        $lblRecordCount.Text = ''
        $btnRetrieve.IsEnabled = $false
        $btnRetrieve.Content = 'Retrieving...'
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        Invoke-UiRefresh -Window $window

        $wantName     = $chkMachineName.IsChecked
        $wantCatalog  = $chkMachineCatalog.IsChecked
        $wantDG       = $chkDeliveryGroup.IsChecked
        $wantRegState = $chkRegistrationState.IsChecked
        $wantMaint    = $chkInMaintenanceMode.IsChecked
        $wantLogon    = $chkLogonMode.IsChecked
        $wantServerIP = $chkServerIP.IsChecked
        $singleServerMode = [bool]$chkSingleServer.IsChecked
        $targetServer = $txtServerName.Text.Trim()

        if ($singleServerMode -and -not $targetServer) {
            [System.Windows.MessageBox]::Show(
                'Please enter a server name for single-server mode.',
                'Validation', 'OK', 'Warning')
            $progressBar.Visibility = 'Hidden'
            $btnRetrieve.IsEnabled = $true
            $btnRetrieve.Content = $originalRetrieveText
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            $statusText.Text = 'Validation failed'
            return
        }
        $tabMain.SelectedIndex = 1

        # Log selected fields
        $selectedFields = @()
        if ($wantName)     { $selectedFields += 'MachineName' }
        if ($wantCatalog)  { $selectedFields += 'CatalogName' }
        if ($wantDG)       { $selectedFields += 'DeliveryGroup' }
        if ($wantRegState) { $selectedFields += 'RegistrationState' }
        if ($wantMaint)    { $selectedFields += 'MaintenanceMode' }
        if ($wantLogon)    { $selectedFields += 'LogonMode' }
        if ($wantServerIP) { $selectedFields += 'ServerIP' }
        Write-Log "DDC retrieval requested — fields: $($selectedFields -join ', ')"
        if ($singleServerMode) { Write-Log "DDC retrieval scope: single server [$targetServer]" }

        try {
            if (-not (Get-PSSnapin -Name 'Citrix*' -ErrorAction SilentlyContinue)) {
                Write-Log 'Loading Citrix Broker snap-in...'
                Add-PSSnapin Citrix* -ErrorAction Stop
                Write-Log 'Citrix Broker snap-in loaded successfully'
            }

            $statusText.Text = 'Retrieving broker machines...'
            Write-Log 'Retrieving broker machines'

            $machines = Get-BrokerMachine -MaxRecordCount 5000 -ErrorAction Stop
            if ($singleServerMode) {
                $machines = @($machines | Where-Object {
                    $fullName = [string]$_.MachineName
                    $shortName = $fullName -replace '^.*\\', ''
                    $dnsName = [string]$_.DNSName
                    $fullName -ieq $targetServer -or $shortName -ieq $targetServer -or $dnsName -ieq $targetServer
                })
            }

            if (-not $machines) {
                if ($singleServerMode) {
                    Write-Log "Broker machine not found for server [$targetServer]" -Level WARN
                    [System.Windows.MessageBox]::Show(
                        "Server '$targetServer' was not found in Broker machine list.",
                        'Information', 'OK', 'Information')
                } else {
                    Write-Log 'No Broker Machines found on this controller' -Level WARN
                    [System.Windows.MessageBox]::Show('No Broker Machines found.', 'Information', 'OK', 'Information')
                }
                $progressBar.Visibility = 'Hidden'
                $statusText.Text = 'No machines found'
                $btnRetrieve.IsEnabled = $true
                return
            }

            $totalMachines = @($machines).Count
            Write-Log "Found $totalMachines broker machines"
            $statusText.Text = "Processing $totalMachines machines..."

            $machineInfo = [System.Collections.Generic.List[object]]::new($totalMachines)
            $counter = 0
            $dnsCache = @{}

            foreach ($machine in $machines) {
                $counter++

                $obj = [ordered]@{
                    SrNumber = $counter
                }
                # Broker returns DOMAIN\Machine; strip domain for cleaner display.
                if ($wantName)     { $obj['MachineName']       = ($machine.MachineName -replace '^.*\\', '') }
                if ($wantCatalog)  { $obj['CatalogName']       = $machine.CatalogName }
                if ($wantDG)       { $obj['DeliveryGroup']     = $machine.DesktopGroupName }
                if ($wantRegState) { $obj['RegistrationState'] = $machine.RegistrationState }
                if ($wantMaint)    { $obj['MaintenanceMode']   = $machine.InMaintenanceMode }
                if ($wantLogon) {
                    # Most Broker SDK versions expose WindowsConnectionSetting (LogonEnabled/Draining/DrainingUntilRestart).
                    if ($machine.PSObject.Properties.Match('WindowsConnectionSetting').Count -gt 0 -and $null -ne $machine.WindowsConnectionSetting) {
                        $mode = [string]$machine.WindowsConnectionSetting
                        $obj['LogonMode'] = if ($mode -eq 'LogonEnabled') { 'Enabled' } else { 'Disabled' }
                    } else {
                        # Fallback heuristic when the explicit logon property is unavailable.
                        $obj['LogonMode'] = if ($machine.InMaintenanceMode) { 'Disabled' } else { 'Enabled' }
                    }
                }
                if ($wantServerIP) {
                    $serverIp = $null
                    $lookupName = if ($machine.DNSName) { [string]$machine.DNSName } else { [string]($machine.MachineName -replace '^.*\\', '') }
                    $cacheKey = $lookupName.ToLowerInvariant()

                    foreach ($propName in @('IPAddress','IPv4Address','IPV4Address')) {
                        if ($machine.PSObject.Properties.Match($propName).Count -gt 0 -and $machine.$propName) {
                            $rawValue = $machine.$propName
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
                            try {
                                $resolved = [System.Net.Dns]::GetHostAddresses($lookupName) |
                                    Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
                                    Select-Object -First 1
                                if ($resolved) { $serverIp = $resolved.IPAddressToString }
                            } catch {
                                # Keep non-terminating; unresolved names remain N/A.
                            }
                            $dnsCache[$cacheKey] = $serverIp
                        }
                    }

                    $obj['ServerIP'] = if ($serverIp) { [string]$serverIp } else { 'N/A' }
                }

                [void]$machineInfo.Add([PSCustomObject]$obj)
                $progressBar.Value = ($counter / $totalMachines) * 100
            }

            $script:ddcFullData = $machineInfo
            $dataGrid.ItemsSource = $machineInfo
            $lblRecordCount.Text = "$totalMachines records"
            $ddcSummary = Get-DdcHealthSummary -Machines $machines
            $statusText.Text = "Retrieved $totalMachines machines | Maintenance: $($ddcSummary.MaintenanceCount) | Unregistered: $($ddcSummary.UnregisteredCount) | Logon Disabled: $($ddcSummary.LogonDisabledCount)"
            if ($singleServerMode) {
                $txtCatalogSummary.Text = "Single Server Mode:`r`n - MC aggregation skipped for [$targetServer]"
                $txtDeliverySummary.Text = "Single Server Mode:`r`n - DG aggregation skipped for [$targetServer]"
            } else {
                $catalogSummary = Get-GroupedCountSummary -Items $machines -PropertyName 'CatalogName' -Label 'Catalogs'
                $deliveryGroupSummary = Get-GroupedCountSummary -Items $machines -PropertyName 'DesktopGroupName' -Label 'DeliveryGroups'
                $txtCatalogSummary.Text = Format-CountSection -Heading 'Catalogs' -Description 'No. of Machines in each MC:' -Detail $catalogSummary.Detail
                $txtDeliverySummary.Text = Format-CountSection -Heading 'DeliveryGroups' -Description 'No. of Machines in each DG:' -Detail $deliveryGroupSummary.Detail
                Write-Log "DDC grouped counts | $($catalogSummary.Detail)"
                Write-Log "DDC grouped counts | $($deliveryGroupSummary.Detail)"
            }
            Write-Log "DDC retrieval complete: $totalMachines machines"
            Write-Log "DDC summary counts | Maintenance=$($ddcSummary.MaintenanceCount) | Unregistered=$($ddcSummary.UnregisteredCount) | LogonDisabled=$($ddcSummary.LogonDisabledCount)"
            if ($singleServerMode) {
                Write-Log "DDC grouped counts skipped (single-server mode: [$targetServer])"
            }

        } catch {
            [System.Windows.MessageBox]::Show("Error: $_", 'Error', 'OK', 'Error')
            Write-Log "DDC retrieval error: $_" -Level ERROR
            $statusText.Text = 'Error during retrieval'
        } finally {
            $progressBar.Visibility = 'Hidden'
            $btnRetrieve.IsEnabled = $true
            $btnRetrieve.Content = $originalRetrieveText
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })

    $btnExport.Add_Click({ Export-DataGridToCSV -DataGrid $dataGrid -DefaultFileName 'Broker_MachineData.csv' })
    $btnClose.Add_Click({ $window.Close() })

    $window.ShowDialog() | Out-Null
}

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
    $details.Add("PowerShell: $psv") | Out-Null
    if ($psv -lt [version]'5.1') {
        $issues.Add("PowerShell 5.1 or later is required. Current version: $psv") | Out-Null
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

function Test-RolePrereqs {
    param(
        [ValidateSet('PVS','DDC')]
        [string]$Role
    )

    $issues = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $details = [System.Collections.Generic.List[string]]::new()

    try {
        if (-not (Get-PSSnapin -Name 'Citrix*' -ErrorAction SilentlyContinue)) {
            Add-PSSnapin Citrix* -ErrorAction SilentlyContinue
        }
    } catch {
        # Continue; command checks below determine final readiness.
    }

    if ($Role -eq 'PVS') {
        if (Get-Command Get-PVSDevice -ErrorAction SilentlyContinue) {
            $details.Add('Citrix PVS cmdlets: Available') | Out-Null
        } else {
            $issues.Add('Citrix PVS cmdlets not found (Get-PVSDevice missing). Install/load Citrix PVS PowerShell SDK.') | Out-Null
        }

        if ((Get-Command Get-DhcpServerv4Scope -ErrorAction SilentlyContinue) -and
            (Get-Command Get-DhcpServerv4OptionValue -ErrorAction SilentlyContinue) -and
            (Get-Command Get-DhcpServerv4Reservation -ErrorAction SilentlyContinue)) {
            $details.Add('DHCP cmdlets: Available') | Out-Null
        } else {
            $warnings.Add('DHCP cmdlets missing. Retrieval still works; DHCP check output may be limited.') | Out-Null
        }
    }

    if ($Role -eq 'DDC') {
        if (Get-Command Get-BrokerMachine -ErrorAction SilentlyContinue) {
            $details.Add('Citrix Broker cmdlets: Available') | Out-Null
        } else {
            $issues.Add('Citrix Broker cmdlets not found (Get-BrokerMachine missing). Install/load Citrix Broker PowerShell SDK.') | Out-Null
        }
    }

    return [PSCustomObject]@{
        Passed = ($issues.Count -eq 0)
        Issues = @($issues)
        Warnings = @($warnings)
        Details = @($details)
    }
}

# ─────────────────────────────────────────────────────────────
# Region: Entry Point
# ─────────────────────────────────────────────────────────────
# Global error trap — catches any unhandled terminating error
trap {
    Write-Log "UNHANDLED EXCEPTION: $_" -Level ERROR
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    [System.Windows.MessageBox]::Show(
        "An unexpected error occurred:`n$_`n`nSee log for details:`n$($script:LogPath)",
        'Fatal Error', 'OK', 'Error')
    continue
}

Write-Log '====== Script started ======'
Write-Log "Host: $($env:COMPUTERNAME) | User: $($env:USERNAME) | PS version: $($PSVersionTable.PSVersion)"

$startupCheck = Test-StartupPrereqs
foreach ($line in $startupCheck.Details) { Write-Log "Precheck | $line" }
foreach ($line in $startupCheck.Warnings) { Write-Log "Precheck warning | $line" -Level WARN }
if (-not $startupCheck.Passed) {
    foreach ($line in $startupCheck.Issues) { Write-Log "Precheck failed | $line" -Level ERROR }
    $msg = @(
        "Startup pre-check failed:"
        ($startupCheck.Issues | ForEach-Object { " - $_" })
        ""
        "See log for details:"
        " - $script:LogPath"
    ) -join "`r`n"
    Write-Host $msg -ForegroundColor Red
    [System.Windows.MessageBox]::Show($msg, 'Prerequisite Check Failed', 'OK', 'Error') | Out-Null
    Write-Log '====== Script ended ======'
    return
}

# Runtime flow:
# 1) Prefer PVS mode when StreamService is active.
# 2) Otherwise try Delivery Controller mode.
# 3) Exit with guidance if neither role is detected.
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
        Write-Log '====== Script ended ======'
        return
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
        Write-Log '====== Script ended ======'
        return
    }
    Show-DDCWindow
} else {
    Write-Host 'This script must be run on a Citrix PVS Server or Delivery Controller.' -ForegroundColor Yellow
    Write-Host "Log file: $script:LogPath" -ForegroundColor Cyan
    Write-Log 'Neither PVS nor DDC role detected on this server — exiting' -Level WARN
}

Write-Log '====== Script ended ======'

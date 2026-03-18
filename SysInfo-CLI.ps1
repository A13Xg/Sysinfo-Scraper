<#
.SYNOPSIS
    SysInfo-CLI - Interactive command-line interface for Sysinfo-Scraper.

.DESCRIPTION
    Provides a colourful, menu-driven CLI for collecting, viewing, and exporting
    comprehensive system information.  Dot-sources the shared SysInfo-Core.ps1
    module for all data-collection and export logic.

    Upgrade from the legacy "Scrape Info.bat" which only captured basic
    systeminfo / ipconfig output.

.AUTHOR
    @13X

.VERSION
    2.0

.PARAMETER ScanAndExport
    Run a full scan and export automatically (non-interactive mode).

.PARAMETER OutputFormat
    Export format when using -ScanAndExport.  Valid values: TXT, CSV, Both.
    Default: Both.

.PARAMETER OutputPath
    Output directory for exported files.  Default: current directory.

.EXAMPLE
    .\SysInfo-CLI.ps1
    Launches the interactive menu.

.EXAMPLE
    .\SysInfo-CLI.ps1 -ScanAndExport -OutputFormat Both -OutputPath C:\Reports
    Runs headless: scans, exports TXT + CSV to C:\Reports, then exits.
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$ScanAndExport,

    [ValidateSet('TXT', 'CSV', 'Both')]
    [string]$OutputFormat = 'Both',

    [string]$OutputPath = '.'
)

# ── Dot-source the core module ──────────────────────────────────────────────────
. "$PSScriptRoot\SysInfo-Core.ps1"

# ── Initialise session log ─────────────────────────────────────────────────────
$null = Initialize-SysInfoLog
Write-SysInfoLog "SysInfo-CLI started (Admin: $($script:IsAdmin))" -Level INFO

# ── Attempt UAC elevation ──────────────────────────────────────────────────────
# Try to re-launch with administrator privileges. Request-AdminElevation calls
# exit 0 on the current process if UAC is accepted; we only reach the code below
# if the user declined or if UAC is blocked by policy.
if (-not $script:IsAdmin) {
    Write-Host ''
    Write-Host '  [*] Requesting administrator privileges via UAC...' -ForegroundColor Cyan
    # Reconstruct any non-default parameters to forward to the elevated instance
    $fwdArgs = @()
    if ($ScanAndExport)                              { $fwdArgs += '-ScanAndExport' }
    if ($OutputFormat -and $OutputFormat -ne 'Both') { $fwdArgs += "-OutputFormat `"$OutputFormat`"" }
    if ($OutputPath   -and $OutputPath   -ne '.')    { $fwdArgs += "-OutputPath `"$OutputPath`"" }
    # PSCommandPath is the most reliable way to get the current script path;
    # fall back to MyInvocation.MyCommand.Path if unavailable.
    $elevScriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    # Will exit this process if UAC is accepted; returns $false otherwise
    $elevated = Request-AdminElevation -ScriptPath $elevScriptPath `
                                       -OriginalArgs $fwdArgs
    if (-not $elevated) {
        Write-Host "  [!] Elevation declined or unavailable – continuing as standard user." -ForegroundColor Yellow
        Write-Host '      Items requiring admin rights will be shown in ' -ForegroundColor Yellow -NoNewline
        Write-Host 'red' -ForegroundColor Red -NoNewline
        Write-Host " as: $($script:NeedsAdminPriv)" -ForegroundColor Yellow
        Write-Host ''
        Write-SysInfoLog 'Admin elevation denied. Continuing as standard user.' -Level WARN
    }
}

# ── Script-level state ──────────────────────────────────────────────────────────
$script:ScanData = $null

# Step labels used during scanning – single source of truth (must match the
# progress messages emitted by Get-SystemInfoData in SysInfo-Core.ps1).
$script:ScanStepLabels = @(
    'Collecting system overview...'
    'Collecting operating system info...'
    'Collecting processor info...'
    'Collecting memory info...'
    'Collecting storage info...'
    'Collecting graphics info...'
    'Collecting network adapter info...'
    'Collecting BIOS info...'
    'Collecting motherboard info...'
    'Collecting battery info...'
    'Collecting hotfixes and startup programs...'
    'Collecting security info...'
    'Collecting environment info...'
    'Collecting display info...'
    'Collecting installed software...'
    'Collecting sound device info...'
    'Collecting printer info...'
)

# ── Graceful Ctrl+C handling ────────────────────────────────────────────────────
[Console]::TreatControlCAsInput = $false
# Unregister any previous handler (prevents accumulation when dot-sourced)
Get-EventSubscriber -SourceIdentifier 'SysInfoCLI.Exiting' -ErrorAction SilentlyContinue |
    Unregister-Event -ErrorAction SilentlyContinue
$script:ExitSubscription = Register-EngineEvent -SourceIdentifier 'SysInfoCLI.Exiting' -Action {
    [Console]::ResetColor()
}

# ── Helper: write a horizontal rule ────────────────────────────────────────────
function Write-Banner {
    <#
    .SYNOPSIS
        Displays the application banner with admin status indicator.
    .DESCRIPTION
        Writes a styled box banner to the console, including the tool name,
        author, and whether the current session is running with administrator
        privileges.  If running as admin, a green [ADMIN] badge is shown;
        otherwise a yellow [STANDARD USER] badge is displayed together with a
        reminder that some values will show 'Needs Admin Priv'.
    #>
    $adminLabel = if ($script:IsAdmin) { '[ADMIN]' } else { '[STANDARD USER]' }
    $adminColor = if ($script:IsAdmin) { 'Green'  } else { 'Yellow'           }
    $banner = @(
        ''
        '  ╔══════════════════════════════════════════════╗'
        '  ║        SYSINFO SCRAPER v2.0 - CLI           ║'
        '  ║              by @13X                        ║'
        '  ╚══════════════════════════════════════════════╝'
        ''
    )
    foreach ($line in $banner) {
        Write-Host $line -ForegroundColor Cyan
    }
    Write-Host "  Privilege level: " -ForegroundColor DarkGray -NoNewline
    Write-Host $adminLabel -ForegroundColor $adminColor
    if ($script:LogFile) {
        Write-Host "  Log file       : $($script:LogFile)" -ForegroundColor DarkGray
    }
    Write-Host ''
}

function Write-Separator {
    <#
    .SYNOPSIS
        Writes a horizontal divider line to the console.
    #>
    Write-Host ('  ' + ('─' * 46)) -ForegroundColor DarkGray
}

# ── Helper: generate timestamped filename ──────────────────────────────────────
function Get-ExportFileName {
    <#
    .SYNOPSIS
        Generates a timestamped export filename for scan results.
    .DESCRIPTION
        Combines the scanned computer name (or the current $env:COMPUTERNAME if
        no scan has been performed) with a yyyyMMdd_HHmmss timestamp and the
        supplied file extension.
    .PARAMETER Extension
        File extension without the leading dot (e.g. 'txt' or 'csv').
    .OUTPUTS
        [string] Filename in the form SysInfo_<ComputerName>_<Timestamp>.<Ext>.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Extension
    )
    $computer = if ($script:ScanData -and $script:ScanData.SystemOverview.ComputerName) {
        $script:ScanData.SystemOverview.ComputerName
    } else {
        $env:COMPUTERNAME
    }
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    return "SysInfo_${computer}_${ts}.${Extension}"
}

# ── Helper: write formatted table with 'Needs Admin Priv' lines coloured red ───
function Write-ColoredTable {
    <#
    .SYNOPSIS
        Writes a pre-formatted table string to the console, colouring any line
        that contains the 'Needs Admin Priv' sentinel in red.
    .DESCRIPTION
        Format-SystemInfoTable returns a plain multi-line string.  This wrapper
        splits the string on newlines and writes each line individually, using
        Write-Host -ForegroundColor Red for lines that contain the sentinel value
        so that missing-privilege entries are visually distinct from normal data.
    .PARAMETER TableText
        The multi-line string produced by Format-SystemInfoTable.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$TableText
    )
    $sentinel = $script:NeedsAdminPriv
    foreach ($line in $TableText -split "`n") {
        if ($line -match [regex]::Escape($sentinel)) {
            Write-Host $line -ForegroundColor Red
        } else {
            Write-Host $line
        }
    }
}

# ── 1. Run Full System Scan ────────────────────────────────────────────────────
function Invoke-FullScan {
    <#
    .SYNOPSIS
        Runs a full system information scan and displays the results.
    .DESCRIPTION
        Invokes Get-SystemInfoData (from SysInfo-Core.ps1) with a progress
        callback that updates a live step-by-step status display.  After the scan
        completes, the results are formatted and printed to the console.  Values
        that could not be retrieved due to insufficient privileges are highlighted
        in red with the text 'Needs Admin Priv'.  The user is then offered the
        option to export the results.
    #>
    Write-Host ''
    Write-Host '  [*] Starting system information scan...' -ForegroundColor Cyan
    if (-not $script:IsAdmin) {
        Write-Host "  [!] Running as standard user – some values will show '" -ForegroundColor Yellow -NoNewline
        Write-Host $script:NeedsAdminPriv -ForegroundColor Red -NoNewline
        Write-Host "'." -ForegroundColor Yellow
    }
    Write-Host ''

    $stepLabels = $script:ScanStepLabels
    $totalSteps = $stepLabels.Count

    # Track which step we're on and which steps succeeded
    $script:CLIStep = 0
    $script:StepResults = @{}

    # Print all steps as pending first
    foreach ($i in 0..($totalSteps - 1)) {
        $num = $i + 1
        $label = $stepLabels[$i]
        $padded = $label.PadRight(42)
        Write-Host "  [$num/$totalSteps] $padded" -ForegroundColor White -NoNewline
        Write-Host '[PENDING ]' -ForegroundColor DarkGray
    }

    # Move cursor back up to overwrite
    try { [Console]::SetCursorPosition(0, [Console]::CursorTop - $totalSteps) } catch {}

    # Progress callback – the core calls this as each category begins
    $progressCB = {
        param([int]$Pct, [string]$Msg)

        $stepLabels = $script:ScanStepLabels
        $totalSteps = $stepLabels.Count

        # Find which step this message corresponds to
        $idx = -1
        for ($i = 0; $i -lt $stepLabels.Count; $i++) {
            if ($Msg -eq $stepLabels[$i]) { $idx = $i; break }
        }
        if ($idx -lt 0) { return }

        $num = $idx + 1
        $padded = $Msg.PadRight(42)

        # Mark previous step as DONE (if any)
        if ($idx -gt 0) {
            $prevIdx  = $idx - 1
            $prevNum  = $prevIdx + 1
            $prevLabel = $stepLabels[$prevIdx].PadRight(42)
            try {
                $savedTop = [Console]::CursorTop
                [Console]::SetCursorPosition(0, $savedTop - 1)
                Write-Host "  [$prevNum/$totalSteps] $prevLabel" -ForegroundColor White -NoNewline
                Write-Host '[DONE ' -ForegroundColor Green -NoNewline
                Write-Host ([char]0x2713) -ForegroundColor Green -NoNewline
                Write-Host ']' -ForegroundColor Green
            } catch {}
        }

        # Show current step as WORKING
        Write-Host "  [$num/$totalSteps] $padded" -ForegroundColor Yellow -NoNewline
        Write-Host '[WORKING...]' -ForegroundColor Yellow
    }

    # Run the scan
    $scanError = $null
    try {
        $script:ScanData = Get-SystemInfoData -ProgressCallback $progressCB -ErrorAction Stop
        Write-SysInfoLog 'Full scan completed successfully.' -Level INFO
    } catch {
        $scanError = $_
        Write-SysInfoLog "Full scan failed: $($_.Exception.Message)" -Level ERROR
    }

    # Mark the last step
    $lastLabel = $stepLabels[$totalSteps - 1].PadRight(42)
    if ($scanError) {
        Write-Host "  [$totalSteps/$totalSteps] $lastLabel" -ForegroundColor Red -NoNewline
        Write-Host '[FAILED ' -ForegroundColor Red -NoNewline
        Write-Host ([char]0x2717) -ForegroundColor Red -NoNewline
        Write-Host ']' -ForegroundColor Red
    } else {
        try { [Console]::SetCursorPosition(0, [Console]::CursorTop - 1) } catch {}
        Write-Host "  [$totalSteps/$totalSteps] $lastLabel" -ForegroundColor White -NoNewline
        Write-Host '[DONE ' -ForegroundColor Green -NoNewline
        Write-Host ([char]0x2713) -ForegroundColor Green -NoNewline
        Write-Host ']' -ForegroundColor Green
    }

    Write-Progress -Activity 'System Information Scan' -Completed

    # Summary
    Write-Host ''
    if ($script:ScanData) {
        $categories = @(
            'SystemOverview', 'OperatingSystem', 'Processor', 'Memory',
            'Storage', 'Graphics', 'NetworkAdapters', 'BIOS',
            'Motherboard', 'Battery', 'Hotfixes', 'StartupPrograms',
            'Security', 'Environment', 'Displays', 'InstalledSoftware',
            'SoundDevices', 'Printers'
        )
        # Count categories that exist and have no Error property (or Error is empty)
        $successCount = 0
        foreach ($cat in $categories) {
            $val = $script:ScanData.PSObject.Properties[$cat]
            if ($val -and $val.Value) {
                $errProp = $val.Value.PSObject.Properties['Error']
                if (-not $errProp -or [string]::IsNullOrEmpty($errProp.Value)) {
                    $successCount++
                }
            }
        }
        $totalCat = $categories.Count
        Write-Host "  Scan complete: $successCount of $totalCat categories collected successfully." -ForegroundColor Green
    } else {
        Write-Host '  Scan failed with an error.' -ForegroundColor Red
        if ($scanError) {
            Write-Host "  Error: $($scanError.Exception.Message)" -ForegroundColor Red
        }
        return
    }

    # Display formatted table – lines containing the admin sentinel are shown in red
    Write-Host ''
    $tableOutput = Format-SystemInfoTable -Data $script:ScanData
    Write-ColoredTable -TableText $tableOutput

    # Prompt to export
    Write-Host ''
    Write-Host '  Would you like to export the results now? (Y/N): ' -ForegroundColor Yellow -NoNewline
    $answer = Read-Host
    if ($answer -match '^[Yy]') {
        Invoke-ExportMenu
    }
}

# ── 2. Export Options ──────────────────────────────────────────────────────────
function Invoke-ExportMenu {
    <#
    .SYNOPSIS
        Presents the interactive export format and path selection menu.
    .DESCRIPTION
        Prompts the user to choose an export format (TXT, CSV, or Both) and an
        output directory, then delegates to Export-ScanResults.  If no scan data
        is available the user is informed and returned to the main menu.
    #>
    if (-not $script:ScanData) {
        Write-Host ''
        Write-Host '  [!] No scan data available. Please run a full scan first (Option 1).' -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  Press Enter to return to the main menu...' -ForegroundColor DarkGray -NoNewline
        $null = Read-Host
        return
    }

    Write-Host ''
    Write-Host '  Export Format:' -ForegroundColor Cyan
    Write-Host '  [1] TXT Only' -ForegroundColor White
    Write-Host '  [2] CSV Only' -ForegroundColor White
    Write-Host '  [3] Both TXT and CSV' -ForegroundColor White
    Write-Host ''
    Write-Host '  Select format (1-3): ' -ForegroundColor Yellow -NoNewline
    $fmt = Read-Host

    switch ($fmt) {
        '1' { $format = 'TXT'  }
        '2' { $format = 'CSV'  }
        '3' { $format = 'Both' }
        default {
            Write-Host '  [!] Invalid selection.' -ForegroundColor Red
            return
        }
    }

    Write-Host ''
    Write-Host "  Output directory (default: current directory): " -ForegroundColor Yellow -NoNewline
    $dir = Read-Host
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = (Get-Location).Path }

    if (-not (Test-Path $dir)) {
        try {
            $null = New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop
        } catch {
            Write-Host "  [!] Cannot create directory: $dir" -ForegroundColor Red
            return
        }
    }

    Export-ScanResults -Format $format -Directory $dir
}

function Export-ScanResults {
    <#
    .SYNOPSIS
        Exports the current scan data to TXT, CSV, or both file formats.
    .DESCRIPTION
        Calls Export-SystemInfoTXT and/or Export-SystemInfoCSV (from SysInfo-Core.ps1)
        with auto-generated timestamped filenames in the specified output directory.
        Each export attempt is individually wrapped in try/catch so a failure in one
        format does not prevent the other from completing.
    .PARAMETER Format
        One of 'TXT', 'CSV', or 'Both'.
    .PARAMETER Directory
        Full path to the output directory.  Must already exist.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('TXT', 'CSV', 'Both')]
        [string]$Format,

        [Parameter(Mandatory)]
        [string]$Directory
    )

    Write-Host ''

    if ($Format -eq 'TXT' -or $Format -eq 'Both') {
        $txtFile = Join-Path $Directory (Get-ExportFileName -Extension 'txt')
        try {
            Export-SystemInfoTXT -Data $script:ScanData -Path $txtFile
            Write-Host "  [+] TXT exported: $txtFile" -ForegroundColor Green
            Write-SysInfoLog "TXT exported to: $txtFile" -Level INFO
        } catch {
            Write-Host "  [!] TXT export failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-SysInfoLog "TXT export failed: $($_.Exception.Message)" -Level ERROR
        }
    }

    if ($Format -eq 'CSV' -or $Format -eq 'Both') {
        $csvFile = Join-Path $Directory (Get-ExportFileName -Extension 'csv')
        try {
            Export-SystemInfoCSV -Data $script:ScanData -Path $csvFile
            Write-Host "  [+] CSV exported: $csvFile" -ForegroundColor Green
            Write-SysInfoLog "CSV exported to: $csvFile" -Level INFO
        } catch {
            Write-Host "  [!] CSV export failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-SysInfoLog "CSV export failed: $($_.Exception.Message)" -Level ERROR
        }
    }

    Write-Host ''
    Write-Host '  Press Enter to continue...' -ForegroundColor DarkGray -NoNewline
    $null = Read-Host
}

# ── 3. View Last Scan Results ──────────────────────────────────────────────────
function Show-LastScanResults {
    <#
    .SYNOPSIS
        Displays the formatted results from the most recent scan.
    .DESCRIPTION
        Re-renders the formatted table from the cached scan data ($script:ScanData).
        Lines containing the 'Needs Admin Priv' sentinel are printed in red.
        If no scan has been performed the user is informed and returned to the
        main menu.
    #>
    if (-not $script:ScanData) {
        Write-Host ''
        Write-Host '  [!] No scan data available. Please run a full scan first (Option 1).' -ForegroundColor Yellow
        Write-Host ''
        Write-Host '  Press Enter to return to the main menu...' -ForegroundColor DarkGray -NoNewline
        $null = Read-Host
        return
    }

    Write-Host ''
    $tableOutput = Format-SystemInfoTable -Data $script:ScanData
    Write-ColoredTable -TableText $tableOutput
    Write-Host ''
    Write-Host '  Press Enter to return to the main menu...' -ForegroundColor DarkGray -NoNewline
    $null = Read-Host
}

# ── 4. About ───────────────────────────────────────────────────────────────────
function Show-About {
    <#
    .SYNOPSIS
        Displays the About screen with tool information and admin status.
    .DESCRIPTION
        Shows version, author, license, and a brief description of the tool.
        Also displays the current admin privilege level and the path to the
        active session log file so users can find it for troubleshooting.
    #>
    Write-Host ''
    Write-Host '  ╔══════════════════════════════════════════════╗' -ForegroundColor Cyan
    Write-Host '  ║              ABOUT THIS TOOL                ║' -ForegroundColor Cyan
    Write-Host '  ╠══════════════════════════════════════════════╣' -ForegroundColor Cyan
    Write-Host '  ║  SysInfo Scraper v2.0                      ║' -ForegroundColor Cyan
    Write-Host '  ║  Author : @13X                             ║' -ForegroundColor Cyan
    Write-Host '  ║  License: See LICENSE file                 ║' -ForegroundColor Cyan
    Write-Host '  ╚══════════════════════════════════════════════╝' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  A comprehensive system-information collection tool' -ForegroundColor White
    Write-Host "  built in PowerShell.  Collects hardware, OS, network," -ForegroundColor White
    Write-Host "  and software data across $($script:ScanStepLabels.Count) categories." -ForegroundColor White
    Write-Host ''
    Write-Host '  This project replaces the original "Scrape Info.bat"' -ForegroundColor DarkGray
    Write-Host '  which relied on systeminfo.exe and ipconfig.  The new' -ForegroundColor DarkGray
    Write-Host '  version uses CIM/WMI queries for richer, structured' -ForegroundColor DarkGray
    Write-Host '  data and supports TXT and CSV export.' -ForegroundColor DarkGray
    Write-Host ''
    $adminLabel = if ($script:IsAdmin) { '[ADMIN]' } else { '[STANDARD USER]' }
    $adminColor = if ($script:IsAdmin) { 'Green'  } else { 'Yellow' }
    Write-Host '  Current privilege level: ' -ForegroundColor DarkGray -NoNewline
    Write-Host $adminLabel -ForegroundColor $adminColor
    if ($script:LogFile) {
        Write-Host "  Session log file: $($script:LogFile)" -ForegroundColor DarkGray
    }
    Write-Host ''
    Write-Host '  Press Enter to return to the main menu...' -ForegroundColor DarkGray -NoNewline
    $null = Read-Host
}

# ── Non-interactive mode ────────────────────────────────────────────────────────
if ($ScanAndExport) {
    Write-Banner
    Write-Host '  Running in non-interactive mode...' -ForegroundColor Cyan
    Write-Host ''

    # Scan
    Write-Host '  [*] Starting system information scan...' -ForegroundColor Cyan
    $scanErr = $null
    try {
        $script:ScanData = Get-SystemInfoData -ProgressCallback {
            param([int]$Pct, [string]$Msg)
            Write-Host "  [$Pct%] $Msg" -ForegroundColor White
        } -ErrorAction Stop
        Write-SysInfoLog 'Non-interactive scan completed.' -Level INFO
    }
    catch {
        $scanErr = $_
        Write-SysInfoLog "Non-interactive scan failed: $($_.Exception.Message)" -Level ERROR
    }

    if (-not $script:ScanData -or $scanErr) {
        Write-Host '  [!] Scan failed. Exiting.' -ForegroundColor Red
        if ($scanErr) { Write-Host "  Error: $($scanErr.Exception.Message)" -ForegroundColor Red }
        exit 1
    }

    Write-Host '  [*] Scan complete.' -ForegroundColor Green

    # Resolve output path
    if (-not (Test-Path $OutputPath)) {
        try {
            $null = New-Item -ItemType Directory -Path $OutputPath -Force -ErrorAction Stop
        }
        catch {
            Write-Host "  [!] Cannot create output directory: $OutputPath" -ForegroundColor Red
            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    }

    Export-ScanResults -Format $OutputFormat -Directory $OutputPath
    Write-Host '  Done.' -ForegroundColor Green
    exit 0
}

# ── Interactive menu loop ───────────────────────────────────────────────────────
$keepRunning = $true
Clear-Host
Write-Banner

while ($keepRunning) {
    Write-Host '  [1] Run Full System Scan' -ForegroundColor White
    Write-Host '  [2] Export Options' -ForegroundColor White
    Write-Host '  [3] View Last Scan Results' -ForegroundColor White
    Write-Host '  [4] About' -ForegroundColor White
    Write-Host '  [5] Exit' -ForegroundColor White
    Write-Host ''
    Write-Host '  Select an option (1-5): ' -ForegroundColor Yellow -NoNewline
    $choice = Read-Host

    switch ($choice) {
        '1' {
            Invoke-FullScan
            Write-Host ''
            Write-Banner
        }
        '2' {
            Invoke-ExportMenu
            Clear-Host
            Write-Banner
        }
        '3' {
            Show-LastScanResults
            Write-Host ''
            Write-Banner
        }
        '4' {
            Show-About
            Clear-Host
            Write-Banner
        }
        '5' {
            Write-Host ''
            Write-Host '  Goodbye! Thanks for using SysInfo Scraper.' -ForegroundColor Cyan
            Write-Host ''
            $keepRunning = $false
        }
        default {
            Write-Host ''
            Write-Host '  [!] Invalid option. Please select 1-5.' -ForegroundColor Red
            Write-Host ''
        }
    }
}

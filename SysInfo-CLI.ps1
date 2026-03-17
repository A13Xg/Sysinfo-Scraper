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

# ── Script-level state ──────────────────────────────────────────────────────────
$script:ScanData = $null

# Step labels used during scanning – single source of truth
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
}

function Write-Separator {
    Write-Host ('  ' + ('─' * 46)) -ForegroundColor DarkGray
}

# ── Helper: generate timestamped filename ──────────────────────────────────────
function Get-ExportFileName {
    param(
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

# ── 1. Run Full System Scan ────────────────────────────────────────────────────
function Invoke-FullScan {
    Write-Host ''
    Write-Host '  [*] Starting system information scan...' -ForegroundColor Cyan
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
    } catch {
        $scanError = $_
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
            'Motherboard', 'Battery', 'Hotfixes', 'StartupPrograms'
        )
        # Count categories that exist and have no Error property (or the Error property is empty)
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

    # Display formatted table
    Write-Host ''
    $tableOutput = Format-SystemInfoTable -Data $script:ScanData
    Write-Host $tableOutput

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
    param(
        [string]$Format,
        [string]$Directory
    )

    Write-Host ''

    if ($Format -eq 'TXT' -or $Format -eq 'Both') {
        $txtFile = Join-Path $Directory (Get-ExportFileName -Extension 'txt')
        try {
            Export-SystemInfoTXT -Data $script:ScanData -Path $txtFile
            Write-Host "  [+] TXT exported: $txtFile" -ForegroundColor Green
        } catch {
            Write-Host "  [!] TXT export failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if ($Format -eq 'CSV' -or $Format -eq 'Both') {
        $csvFile = Join-Path $Directory (Get-ExportFileName -Extension 'csv')
        try {
            Export-SystemInfoCSV -Data $script:ScanData -Path $csvFile
            Write-Host "  [+] CSV exported: $csvFile" -ForegroundColor Green
        } catch {
            Write-Host "  [!] CSV export failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host ''
    Write-Host '  Press Enter to continue...' -ForegroundColor DarkGray -NoNewline
    $null = Read-Host
}

# ── 3. View Last Scan Results ──────────────────────────────────────────────────
function Show-LastScanResults {
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
    Write-Host $tableOutput
    Write-Host ''
    Write-Host '  Press Enter to return to the main menu...' -ForegroundColor DarkGray -NoNewline
    $null = Read-Host
}

# ── 4. About ───────────────────────────────────────────────────────────────────
function Show-About {
    Write-Host ''
    Write-Host '  ╔══════════════════════════════════════════════╗' -ForegroundColor Cyan
    Write-Host '  ║              ABOUT THIS TOOL                ║' -ForegroundColor Cyan
    Write-Host '  ╠══════════════════════════════════════════════╣' -ForegroundColor Cyan
    Write-Host '  ║  Sysinfo Scraper v2.0                      ║' -ForegroundColor Cyan
    Write-Host '  ║  Author : @13X                             ║' -ForegroundColor Cyan
    Write-Host '  ║  License: See LICENSE file                 ║' -ForegroundColor Cyan
    Write-Host '  ╚══════════════════════════════════════════════╝' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  A comprehensive system-information collection tool' -ForegroundColor White
    Write-Host '  built in PowerShell.  Collects hardware, OS, network,' -ForegroundColor White
    Write-Host '  and software data across 11 categories.' -ForegroundColor White
    Write-Host ''
    Write-Host '  This project replaces the original "Scrape Info.bat"' -ForegroundColor DarkGray
    Write-Host '  which relied on systeminfo.exe and ipconfig.  The new' -ForegroundColor DarkGray
    Write-Host '  version uses CIM/WMI queries for richer, structured' -ForegroundColor DarkGray
    Write-Host '  data and supports TXT and CSV export.' -ForegroundColor DarkGray
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
    $script:ScanData = Get-SystemInfoData -ProgressCallback {
        param([int]$Pct, [string]$Msg)
        Write-Host "  [$Pct%] $Msg" -ForegroundColor White
    }

    if (-not $script:ScanData) {
        Write-Host '  [!] Scan failed. Exiting.' -ForegroundColor Red
        exit 1
    }

    Write-Host '  [*] Scan complete.' -ForegroundColor Green

    # Resolve output path
    if (-not (Test-Path $OutputPath)) {
        $null = New-Item -ItemType Directory -Path $OutputPath -Force
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
            Write-Host '  Goodbye! Thanks for using Sysinfo Scraper.' -ForegroundColor Cyan
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

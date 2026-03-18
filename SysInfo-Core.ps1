<#
.SYNOPSIS
    SysInfo-Core - Shared PowerShell module for comprehensive system information collection.

.DESCRIPTION
    Core module dot-sourced by both CLI and GUI front-ends of Sysinfo-Scraper.
    Collects detailed hardware, software, and configuration data using Get-CimInstance
    and exposes export helpers for TXT, CSV, and formatted console output.

    Upgrade from the legacy "Scrape Info.bat" which only captured basic systeminfo/ipconfig.

.AUTHOR
    @13X

.VERSION
    2.0

.NOTES
    Requires PowerShell 5.1+.  All WMI queries use Get-CimInstance (CIM cmdlets)
    instead of the deprecated WMIC utility.
#>

#Requires -Version 5.1

# ── Admin privilege detection ──────────────────────────────────────────────────
# Evaluated once when the module is dot-sourced by CLI or GUI front-ends.
# Wrapped in try/catch in case the WindowsPrincipal API is unavailable
# (e.g. on non-Windows platforms or restricted environments).
try {
    $script:IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
catch {
    $script:IsAdmin = $false
}

# Sentinel value shown (in red in CLI/GUI) when a value cannot be read without
# elevated administrator privileges.
$script:NeedsAdminPriv = 'Needs Admin Priv'

# Path to the active session log file; set by Initialize-SysInfoLog.
$script:LogFile = $null

# ── Known-Models Hashtable ──────────────────────────────────────────────────────
# Maps normalised "MakeModel" keys to image filenames used by Get-HardwareImagePath.

$script:KnownModels = @{
    # Dell
    'DellOptiplex3080'       = 'DellOptiplex3080.png'
    'DellOptiplex5080'       = 'DellOptiplex5080.png'
    'DellOptiplex7080'       = 'DellOptiplex7080.png'
    'DellOptiplex7090'       = 'DellOptiplex7090.png'
    'DellOptiplex7010'       = 'DellOptiplex7010.png'
    'DellLatitude5520'       = 'DellLatitude5520.png'
    'DellLatitude5530'       = 'DellLatitude5530.png'
    'DellLatitude5540'       = 'DellLatitude5540.png'
    'DellLatitude7420'       = 'DellLatitude7420.png'
    'DellLatitude7430'       = 'DellLatitude7430.png'
    'DellPrecision3660'      = 'DellPrecision3660.png'
    'DellPrecision5570'      = 'DellPrecision5570.png'
    'DellXPS139310'          = 'DellXPS139310.png'
    'DellXPS159520'          = 'DellXPS159520.png'
    'DellInspiron155510'     = 'DellInspiron155510.png'

    # HP
    'HPEliteDesk800G6'       = 'HPEliteDesk800G6.png'
    'HPEliteDesk800G8'       = 'HPEliteDesk800G8.png'
    'HPEliteDesk800G9'       = 'HPEliteDesk800G9.png'
    'HPProDesk400G7'         = 'HPProDesk400G7.png'
    'HPProDesk600G6'         = 'HPProDesk600G6.png'
    'HPEliteBook840G8'       = 'HPEliteBook840G8.png'
    'HPEliteBook840G9'       = 'HPEliteBook840G9.png'
    'HPEliteBook850G8'       = 'HPEliteBook850G8.png'
    'HPProBook450G8'         = 'HPProBook450G8.png'
    'HPProBook450G9'         = 'HPProBook450G9.png'
    'HPZBook15G8'            = 'HPZBook15G8.png'
    'HPZBookFury16G9'        = 'HPZBookFury16G9.png'
    'HPEliteDragonfly'       = 'HPEliteDragonfly.png'
    'HPZ2TowerG9'            = 'HPZ2TowerG9.png'
    'HPZ4G5'                 = 'HPZ4G5.png'

    # Lenovo
    'LenovoThinkPadT14Gen3'  = 'LenovoThinkPadT14Gen3.png'
    'LenovoThinkPadT14sGen3' = 'LenovoThinkPadT14sGen3.png'
    'LenovoThinkPadX1CarbonGen10' = 'LenovoThinkPadX1CarbonGen10.png'
    'LenovoThinkPadX1CarbonGen11' = 'LenovoThinkPadX1CarbonGen11.png'
    'LenovoThinkPadL14Gen3'  = 'LenovoThinkPadL14Gen3.png'
    'LenovoThinkPadE14Gen4'  = 'LenovoThinkPadE14Gen4.png'
    'LenovoThinkCentreM70q'  = 'LenovoThinkCentreM70q.png'
    'LenovoThinkCentreM90q'  = 'LenovoThinkCentreM90q.png'
    'LenovoThinkCentreM720s' = 'LenovoThinkCentreM720s.png'
    'LenovoThinkStationP340' = 'LenovoThinkStationP340.png'
    'LenovoThinkStationP360' = 'LenovoThinkStationP360.png'
    'LenovoIdeaPad5Pro'      = 'LenovoIdeaPad5Pro.png'
    'LenovoLegion5Pro'       = 'LenovoLegion5Pro.png'
    'LenovoYoga9i'           = 'LenovoYoga9i.png'
    'LenovoIdeaCentre5'      = 'LenovoIdeaCentre5.png'

    # Acer
    'AcerAspire5'            = 'AcerAspire5.png'
    'AcerAspire7'            = 'AcerAspire7.png'
    'AcerSwift3'             = 'AcerSwift3.png'
    'AcerSwift5'             = 'AcerSwift5.png'
    'AcerSpin5'              = 'AcerSpin5.png'
    'AcerNitro5'             = 'AcerNitro5.png'
    'AcerPredatorHelios300'  = 'AcerPredatorHelios300.png'
    'AcerTravelMateP2'       = 'AcerTravelMateP2.png'
    'AcerTravelMateP6'       = 'AcerTravelMateP6.png'
    'AcerVeritonN4680G'      = 'AcerVeritonN4680G.png'
    'AcerVeritonX4680G'      = 'AcerVeritonX4680G.png'
    'AcerConceptD3'          = 'AcerConceptD3.png'

    # ASUS
    'ASUSZenBook14'          = 'ASUSZenBook14.png'
    'ASUSZenBook13'          = 'ASUSZenBook13.png'
    'ASUSVivoBook15'         = 'ASUSVivoBook15.png'
    'ASUSVivoBookS15'        = 'ASUSVivoBookS15.png'
    'ASUSROGZEPHYRUSG14'     = 'ASUSROGZEPHYRUSG14.png'
    'ASUSROGZEPHYRUSG15'     = 'ASUSROGZEPHYRUSG15.png'
    'ASUSROGStrixG15'        = 'ASUSROGStrixG15.png'
    'ASUSTUFGamingF15'       = 'ASUSTUFGamingF15.png'
    'ASUSProArtStudioBook16' = 'ASUSProArtStudioBook16.png'
    'ASUSExpertBookB5'       = 'ASUSExpertBookB5.png'
    'ASUSExpertBookB9'       = 'ASUSExpertBookB9.png'
    'ASUSChromebookFlip'     = 'ASUSChromebookFlip.png'

    # MSI
    'MSIPrestige14'          = 'MSIPrestige14.png'
    'MSIPrestige15'          = 'MSIPrestige15.png'
    'MSIModern14'            = 'MSIModern14.png'
    'MSIModern15'            = 'MSIModern15.png'
    'MSIGS66Stealth'         = 'MSIGS66Stealth.png'
    'MSIGS76Stealth'         = 'MSIGS76Stealth.png'
    'MSIGE76Raider'          = 'MSIGE76Raider.png'
    'MSICreator15'           = 'MSICreator15.png'
    'MSICreatorZ16'          = 'MSICreatorZ16.png'
    'MSISummitE16Flip'       = 'MSISummitE16Flip.png'
    'MSITridentX'            = 'MSITridentX.png'
    'MSIInfiniteRS'          = 'MSIInfiniteRS.png'

    # Apple
    'AppleMacBookPro14'      = 'AppleMacBookPro14.png'
    'AppleMacBookPro16'      = 'AppleMacBookPro16.png'
    'AppleMacBookAir13'      = 'AppleMacBookAir13.png'
    'AppleMacBookAir15'      = 'AppleMacBookAir15.png'
    'AppleiMac24'            = 'AppleiMac24.png'
    'AppleMacMini'           = 'AppleMacMini.png'
    'AppleMacStudio'         = 'AppleMacStudio.png'
    'AppleMacPro'            = 'AppleMacPro.png'
    'AppleiMacPro'           = 'AppleiMacPro.png'
    'AppleMacBookPro13'      = 'AppleMacBookPro13.png'
    'AppleMacBookAir'        = 'AppleMacBookAir.png'

    # Microsoft Surface
    'MicrosoftSurfacePro9'   = 'MicrosoftSurfacePro9.png'
    'MicrosoftSurfacePro8'   = 'MicrosoftSurfacePro8.png'
    'MicrosoftSurfacePro7'   = 'MicrosoftSurfacePro7.png'
    'MicrosoftSurfaceLaptop5'= 'MicrosoftSurfaceLaptop5.png'
    'MicrosoftSurfaceLaptop4'= 'MicrosoftSurfaceLaptop4.png'
    'MicrosoftSurfaceLaptopStudio' = 'MicrosoftSurfaceLaptopStudio.png'
    'MicrosoftSurfaceLaptopGo2'    = 'MicrosoftSurfaceLaptopGo2.png'
    'MicrosoftSurfaceBook3'  = 'MicrosoftSurfaceBook3.png'
    'MicrosoftSurfaceGo3'    = 'MicrosoftSurfaceGo3.png'
    'MicrosoftSurfaceStudio2'= 'MicrosoftSurfaceStudio2.png'
    'MicrosoftSurfaceHub2S'  = 'MicrosoftSurfaceHub2S.png'
}

# ── Utility Functions ───────────────────────────────────────────────────────────

function Get-ChassisTypeName {
    <#
    .SYNOPSIS
        Maps Win32_SystemEnclosure ChassisTypes integers to human-readable names.
    .DESCRIPTION
        Returns a PSCustomObject with properties 'Name' (e.g. "Desktop", "Notebook")
        and 'Category' (Desktop, Laptop, Server, or Other).
    .PARAMETER ChassisType
        One or more integer chassis-type values from Win32_SystemEnclosure.
    .OUTPUTS
        [PSCustomObject] with Name and Category properties.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int[]]$ChassisType
    )

    $map = @{
        1  = @{ Name = 'Other';                   Category = 'Other'   }
        2  = @{ Name = 'Unknown';                  Category = 'Other'   }
        3  = @{ Name = 'Desktop';                  Category = 'Desktop' }
        4  = @{ Name = 'Low Profile Desktop';      Category = 'Desktop' }
        5  = @{ Name = 'Pizza Box';                Category = 'Desktop' }
        6  = @{ Name = 'Mini Tower';               Category = 'Desktop' }
        7  = @{ Name = 'Tower';                    Category = 'Desktop' }
        8  = @{ Name = 'Portable';                 Category = 'Laptop'  }
        9  = @{ Name = 'Laptop';                   Category = 'Laptop'  }
        10 = @{ Name = 'Notebook';                 Category = 'Laptop'  }
        11 = @{ Name = 'Hand Held';                Category = 'Other'   }
        12 = @{ Name = 'Docking Station';          Category = 'Other'   }
        13 = @{ Name = 'All in One';               Category = 'Desktop' }
        14 = @{ Name = 'Sub Notebook';             Category = 'Laptop'  }
        15 = @{ Name = 'Space-Saving';             Category = 'Desktop' }
        16 = @{ Name = 'Lunch Box';                Category = 'Desktop' }
        17 = @{ Name = 'Main System Chassis';      Category = 'Desktop' }
        18 = @{ Name = 'Expansion Chassis';        Category = 'Other'   }
        19 = @{ Name = 'SubChassis';               Category = 'Other'   }
        20 = @{ Name = 'Bus Expansion Chassis';    Category = 'Other'   }
        21 = @{ Name = 'Peripheral Chassis';       Category = 'Other'   }
        22 = @{ Name = 'Storage Chassis';          Category = 'Server'  }
        23 = @{ Name = 'Rack Mount Chassis';       Category = 'Server'  }
        24 = @{ Name = 'Sealed-Case PC';           Category = 'Desktop' }
        25 = @{ Name = 'Multi-system Chassis';     Category = 'Server'  }
        26 = @{ Name = 'Compact PCI';              Category = 'Server'  }
        27 = @{ Name = 'Advanced TCA';             Category = 'Server'  }
        28 = @{ Name = 'Blade';                    Category = 'Server'  }
        29 = @{ Name = 'Blade Enclosure';          Category = 'Server'  }
        30 = @{ Name = 'Tablet';                   Category = 'Laptop'  }
        31 = @{ Name = 'Convertible';              Category = 'Laptop'  }
        32 = @{ Name = 'Detachable';               Category = 'Laptop'  }
        33 = @{ Name = 'IoT Gateway';              Category = 'Other'   }
        34 = @{ Name = 'Embedded PC';              Category = 'Desktop' }
        35 = @{ Name = 'Mini PC';                  Category = 'Desktop' }
        36 = @{ Name = 'Stick PC';                 Category = 'Desktop' }
    }

    # Use the first recognised type
    foreach ($ct in $ChassisType) {
        if ($map.ContainsKey($ct)) {
            return [PSCustomObject]@{
                Name     = $map[$ct].Name
                Category = $map[$ct].Category
            }
        }
    }

    return [PSCustomObject]@{
        Name     = 'Unknown'
        Category = 'Other'
    }
}

function Get-FormattedUptime {
    <#
    .SYNOPSIS
        Converts a TimeSpan into a friendly "X days, Y hours, Z minutes" string.
    .PARAMETER TimeSpan
        The TimeSpan value to format.
    .OUTPUTS
        [string]
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [TimeSpan]$TimeSpan
    )

    $parts = @()
    if ($TimeSpan.Days   -gt 0) { $parts += "$($TimeSpan.Days) day$(if ($TimeSpan.Days   -ne 1) {'s'})" }
    if ($TimeSpan.Hours  -gt 0) { $parts += "$($TimeSpan.Hours) hour$(if ($TimeSpan.Hours  -ne 1) {'s'})" }
    if ($TimeSpan.Minutes -gt 0) { $parts += "$($TimeSpan.Minutes) minute$(if ($TimeSpan.Minutes -ne 1) {'s'})" }

    if ($parts.Count -eq 0) { return '0 minutes' }
    return $parts -join ', '
}

function Get-DiskMediaType {
    <#
    .SYNOPSIS
        Determines whether a physical disk is SSD, HDD, or NVMe.
    .DESCRIPTION
        Uses MSFT_PhysicalDisk MediaType / BusType when available, falls back to
        model-string heuristics.
    .PARAMETER MediaType
        The MediaType integer from MSFT_PhysicalDisk (0=Unspecified, 3=HDD, 4=SSD).
    .PARAMETER BusType
        The BusType integer from MSFT_PhysicalDisk (17=NVMe).
    .PARAMETER Model
        The model string of the disk for heuristic detection.
    .OUTPUTS
        [string] "SSD", "HDD", "NVMe", or "Unknown"
    #>
    [CmdletBinding()]
    param(
        [int]$MediaType = 0,
        [int]$BusType   = 0,
        [string]$Model  = ''
    )

    # NVMe bus type takes priority
    if ($BusType -eq 17) { return 'NVMe' }

    switch ($MediaType) {
        4       { return 'SSD' }
        3       { return 'HDD' }
        default {
            # Heuristic fallback
            if ($Model -match 'NVMe|NVME') { return 'NVMe' }
            if ($Model -match 'SSD|Solid.State') { return 'SSD' }
            if ($Model -match 'HDD|Hard.Disk') { return 'HDD' }
            return 'Unknown'
        }
    }
}

# ── Hardware Image Resolution ───────────────────────────────────────────────────

function Get-HardwareImagePath {
    <#
    .SYNOPSIS
        Resolves the best-matching hardware image path for a given manufacturer and model.
    .DESCRIPTION
        Searches the hardwareImages directory using a four-tier fallback:
          1. Exact model match  – hardwareImages/{Make}{Model}.png
          2. Type-generic match – hardwareImages/{Make}Generic{Type}.png
          3. Make-generic match – hardwareImages/{Make}Generic.png
          4. Global fallback    – hardwareImages/GenericComputer.png
        Strings are normalised (spaces and special characters removed) before lookup.
    .PARAMETER Manufacturer
        System manufacturer / make string.
    .PARAMETER Model
        System model string.
    .PARAMETER ChassisCategory
        Device category: Desktop, Laptop, Server, or Other.  Used for tier-2 lookup.
    .PARAMETER BasePath
        Root directory of the application.  Defaults to the script's own directory.
    .OUTPUTS
        [string] Full path to the resolved image file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Manufacturer,

        [Parameter(Mandatory)]
        [string]$Model,

        [string]$ChassisCategory = 'Desktop',

        [string]$BasePath = $PSScriptRoot
    )

    # Normalise: strip everything except alphanumerics
    $normMake  = ($Manufacturer -replace '[^A-Za-z0-9]', '')
    $normModel = ($Model        -replace '[^A-Za-z0-9]', '')
    $imageDir  = Join-Path $BasePath 'hardwareImages'

    # Tier 1 – exact known-model lookup then file check
    $exactKey = "$normMake$normModel"
    if ($script:KnownModels.ContainsKey($exactKey)) {
        $candidate = Join-Path $imageDir $script:KnownModels[$exactKey]
        if (Test-Path $candidate) { return $candidate }
    }

    # Also check raw filename on disk even if not in the hashtable
    $rawFile = Join-Path $imageDir "$exactKey.png"
    if (Test-Path $rawFile) { return $rawFile }

    # Tier 2 – manufacturer + generic type
    $typeFile = Join-Path $imageDir "${normMake}Generic${ChassisCategory}.png"
    if (Test-Path $typeFile) { return $typeFile }

    # Tier 3 – manufacturer generic
    $makeFile = Join-Path $imageDir "${normMake}Generic.png"
    if (Test-Path $makeFile) { return $makeFile }

    # Tier 4 – global fallback
    $fallback = Join-Path $imageDir 'GenericComputer.png'
    return $fallback
}

# ── Logging Functions ──────────────────────────────────────────────────────────

function Initialize-SysInfoLog {
    <#
    .SYNOPSIS
        Creates and initialises a timestamped session log file.
    .DESCRIPTION
        Creates a log directory under $env:TEMP\SysInfoScraper\ if it does not
        exist, then creates a new log file named SysInfo_<timestamp>.log and
        writes a header block containing computer name, username, admin status,
        and PowerShell version.  The path is stored in $script:LogFile for use
        by Write-SysInfoLog.
    .OUTPUTS
        [string] Full path to the created log file.
    #>
    [CmdletBinding()]
    param()

    # Determine log directory – fall back to system temp or script root if $env:TEMP is unavailable
    $tempRoot = if ($env:TEMP) { $env:TEMP } elseif ($env:TMPDIR) { $env:TMPDIR } else { $PSScriptRoot }
    $logDir = Join-Path $tempRoot 'SysInfoScraper'
    if (-not (Test-Path $logDir)) {
        try {
            $null = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop
        }
        catch {
            # Cannot create log directory; log directly in temp root as fallback
            $logDir = $tempRoot
        }
    }

    $timestamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:LogFile = Join-Path $logDir "SysInfo_${timestamp}.log"

    # Resolve current user name – WindowsIdentity.GetCurrent() is Windows-only
    $currentUser = try { [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }
                   catch { "$env:USERNAME" }

    $header = @(
        '=' * 70
        '  SysInfo Scraper - Session Log'
        "  Started : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "  Computer: $env:COMPUTERNAME"
        "  User    : $currentUser"
        "  Admin   : $($script:IsAdmin)"
        "  PS Ver  : $($PSVersionTable.PSVersion)"
        '=' * 70
        ''
    )

    try {
        $header | Out-File -FilePath $script:LogFile -Encoding UTF8 -Force
    }
    catch {
        # Silently suppress – logging is non-critical
    }

    return $script:LogFile
}

function Write-SysInfoLog {
    <#
    .SYNOPSIS
        Appends a timestamped entry to the active session log file.
    .DESCRIPTION
        Writes a formatted log entry to the file initialised by Initialize-SysInfoLog.
        If the log file has not been initialised the call is silently discarded.
        Log failures are suppressed so that they never interrupt normal operation.
    .PARAMETER Message
        The message text to record.
    .PARAMETER Level
        Severity level: INFO, WARN, ERROR, or DEBUG.  Defaults to INFO.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    if (-not $script:LogFile) { return }

    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$Level] $Message"

    try {
        $entry | Out-File -FilePath $script:LogFile -Encoding UTF8 -Append
    }
    catch {
        # Suppress write failures to avoid cascading errors
    }
}

function Test-IsAdministrator {
    <#
    .SYNOPSIS
        Returns $true if the current process has administrator privileges.
    .DESCRIPTION
        Uses the .NET WindowsPrincipal class to check membership in the built-in
        Administrators group.  This is the canonical, reliable way to detect
        elevated privileges on Windows.
    .OUTPUTS
        [bool] $true when running as administrator, $false otherwise.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    try {
        return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        # WindowsPrincipal is not available on non-Windows platforms; default to $false
        return $false
    }
}

function Request-AdminElevation {
    <#
    .SYNOPSIS
        Attempts to re-launch the calling script with UAC-elevated privileges.
    .DESCRIPTION
        If the current process is already elevated this function returns $true
        immediately.  Otherwise it tries to start a new PowerShell process using
        the 'RunAs' verb (which triggers UAC).

        If the user approves UAC the elevated process takes over and this function
        calls exit 0 on the original (non-admin) process – so it will NEVER return
        on success.

        If the user declines UAC, or if UAC is disabled / blocked by policy, the
        function catches the resulting error and returns $false so the caller can
        continue running with standard user privileges.
    .PARAMETER ScriptPath
        Full path to the script to re-launch.  Callers should pass
        $MyInvocation.MyCommand.Path or $PSCommandPath.
    .PARAMETER OriginalArgs
        Array of command-line arguments to forward to the elevated instance.
    .OUTPUTS
        [bool] $true if already admin, $false if elevation was declined or blocked.
              Does not return if elevation was successful (exits current process).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$ScriptPath,

        [string[]]$OriginalArgs = @()
    )

    # Already elevated – nothing to do
    if (Test-IsAdministrator) {
        Write-SysInfoLog 'Process is already running with administrator privileges.' -Level INFO
        return $true
    }

    # Guard: cannot elevate without a script path
    if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
        Write-SysInfoLog 'Request-AdminElevation: ScriptPath is empty; cannot re-launch for elevation.' -Level WARN
        return $false
    }

    # Guard: script file must exist
    if (-not (Test-Path $ScriptPath)) {
        Write-SysInfoLog "Request-AdminElevation: ScriptPath '$ScriptPath' does not exist; cannot elevate." -Level WARN
        return $false
    }

    Write-SysInfoLog "Requesting UAC elevation for: $ScriptPath" -Level INFO

    # Build the argument list for the elevated PowerShell instance
    $argList = "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""
    if ($OriginalArgs.Count -gt 0) {
        $argList += ' ' + ($OriginalArgs -join ' ')
    }

    try {
        Start-Process -FilePath 'powershell.exe' `
                      -ArgumentList $argList `
                      -Verb RunAs `
                      -ErrorAction Stop

        Write-SysInfoLog 'UAC elevation accepted; elevated process launched. Exiting non-admin instance.' -Level INFO
        exit 0   # Elevated process takes over; exit this non-admin process
    }
    catch {
        Write-SysInfoLog "UAC elevation failed or was denied: $($_.Exception.Message)" -Level WARN
        return $false
    }
}

# ── Main Data-Collection Function ──────────────────────────────────────────────

function Get-SystemInfoData {
    <#
    .SYNOPSIS
        Collects comprehensive system information and returns it as a PSCustomObject.
    .DESCRIPTION
        Queries CIM/WMI classes to gather data across 11 categories: System Overview,
        Operating System, Processor, Memory, Storage, Graphics, Network, BIOS,
        Motherboard, Battery, Hotfixes, and Startup Programs.

        Each category is wrapped in try/catch so a single failure does not abort the
        entire scan.  Progress is reported through an optional scriptblock callback.
    .PARAMETER ProgressCallback
        Optional scriptblock invoked with two arguments:
          [int]    PercentComplete  (0-100)
          [string] StatusMessage
        The caller (CLI or GUI) can use this to drive a progress bar or log.
    .OUTPUTS
        [PSCustomObject] containing all collected data.
    .EXAMPLE
        $data = Get-SystemInfoData -ProgressCallback { param($pct,$msg) Write-Progress -Activity 'Scan' -Status $msg -PercentComplete $pct }
    #>
    [CmdletBinding()]
    param(
        [scriptblock]$ProgressCallback,

        [bool]$IsAdmin = $script:IsAdmin
    )

    # Helper to report progress
    $totalSteps = 17
    $currentStep = 0
    $reportProgress = {
        param([string]$Message)
        $script:currentStep++
        $pct = [math]::Min(100, [math]::Round(($script:currentStep / $totalSteps) * 100))
        if ($ProgressCallback) {
            & $ProgressCallback $pct $Message
        }
        Write-Progress -Activity 'System Information Scan' -Status $Message -PercentComplete $pct
    }

    # Result container
    $info = [ordered]@{}

    # ── 1. System Overview ──────────────────────────────────────────────────────
    & $reportProgress 'Collecting system overview...'
    try {
        $cs  = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $enc = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction Stop

        $chassisTypes = @($enc.ChassisTypes | Where-Object { $_ })
        if ($chassisTypes.Count -eq 0) { $chassisTypes = @(1) }
        $chassisInfo = Get-ChassisTypeName -ChassisType $chassisTypes

        # VM detection
        $systemType = $chassisInfo.Category
        $vmIndicators = @('Virtual', 'VMware', 'VirtualBox', 'Hyper-V', 'KVM', 'Xen', 'QEMU', 'Parallels')
        $combinedStr = "$($cs.Manufacturer) $($cs.Model)"
        foreach ($vi in $vmIndicators) {
            if ($combinedStr -match $vi) { $systemType = 'Virtual Machine'; break }
        }

        $info['SystemOverview'] = [PSCustomObject]@{
            ComputerName       = $cs.Name
            CurrentUser        = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            Domain             = if ($cs.PartOfDomain) { $cs.Domain } else { $cs.Workgroup }
            IsDomain           = [bool]$cs.PartOfDomain
            SystemManufacturer = $cs.Manufacturer
            SystemModel        = $cs.Model
            SystemType         = $systemType
            ChassisType        = $chassisInfo.Name
            SerialNumber       = $enc.SerialNumber
            AssetTag           = $enc.SMBIOSAssetTag
        }
        Write-SysInfoLog 'SystemOverview: collected.' -Level INFO
    }
    catch {
        Write-SysInfoLog "SystemOverview collection failed: $($_.Exception.Message)" -Level ERROR
        $info['SystemOverview'] = [PSCustomObject]@{
            ComputerName = $env:COMPUTERNAME
            CurrentUser  = $env:USERNAME
            Error        = $_.Exception.Message
        }
    }

    # ── 2. Operating System ─────────────────────────────────────────────────────
    & $reportProgress 'Collecting operating system info...'
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $uptime = (Get-Date) - $os.LastBootUpTime

        $info['OperatingSystem'] = [PSCustomObject]@{
            OSName           = $os.Caption
            OSVersion        = $os.Version
            OSBuild          = $os.BuildNumber
            OSArchitecture   = $os.OSArchitecture
            InstallDate      = $os.InstallDate
            LastBootTime     = $os.LastBootUpTime
            Uptime           = Get-FormattedUptime -TimeSpan $uptime
            RegisteredOwner  = $os.RegisteredUser
            ProductID        = $os.SerialNumber
            WindowsDirectory = $os.WindowsDirectory
            SystemDirectory  = $os.SystemDirectory
        }
        Write-SysInfoLog 'OperatingSystem: collected.' -Level INFO
    }
    catch {
        Write-SysInfoLog "OperatingSystem collection failed: $($_.Exception.Message)" -Level ERROR
        $info['OperatingSystem'] = [PSCustomObject]@{ Error = $_.Exception.Message }
    }

    # ── 3. Processor ────────────────────────────────────────────────────────────
    & $reportProgress 'Collecting processor info...'
    try {
        $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Select-Object -First 1

        $archMap = @{
            0 = 'x86'; 1 = 'MIPS'; 2 = 'Alpha'; 3 = 'PowerPC'; 5 = 'ARM'
            6 = 'ia64'; 9 = 'x64'; 12 = 'ARM64'
        }

        $info['Processor'] = [PSCustomObject]@{
            ProcessorName            = $cpu.Name.Trim()
            Manufacturer             = $cpu.Manufacturer
            NumberOfCores            = $cpu.NumberOfCores
            NumberOfLogicalProcessors = $cpu.NumberOfLogicalProcessors
            MaxClockSpeedMHz         = $cpu.MaxClockSpeed
            CurrentClockSpeedMHz     = $cpu.CurrentClockSpeed
            Architecture             = if ($archMap.ContainsKey([int]$cpu.Architecture)) { $archMap[[int]$cpu.Architecture] } else { "Unknown ($($cpu.Architecture))" }
            L2CacheSizeKB            = $cpu.L2CacheSize
            L3CacheSizeKB            = $cpu.L3CacheSize
            SocketDesignation        = $cpu.SocketDesignation
        }
        Write-SysInfoLog 'Processor: collected.' -Level INFO
    }
    catch {
        Write-SysInfoLog "Processor collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Processor'] = [PSCustomObject]@{ Error = $_.Exception.Message }
    }

    # ── 4. Memory (RAM) ────────────────────────────────────────────────────────
    & $reportProgress 'Collecting memory info...'
    try {
        $os2 = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $totalGB = [math]::Round($os2.TotalVisibleMemorySize / 1MB, 2)
        $freeGB  = [math]::Round($os2.FreePhysicalMemory     / 1MB, 2)
        $usedGB  = [math]::Round($totalGB - $freeGB, 2)
        $usePct  = if ($totalGB -gt 0) { [math]::Round(($usedGB / $totalGB) * 100, 1) } else { 0 }

        $formFactorMap = @{
            0 = 'Unknown'; 1 = 'Other'; 2 = 'SIP'; 3 = 'DIP'; 4 = 'ZIP'; 5 = 'SOJ'
            6 = 'Proprietary'; 7 = 'SIMM'; 8 = 'DIMM'; 9 = 'TSOP'; 10 = 'PGA'
            11 = 'RIMM'; 12 = 'SODIMM'; 13 = 'SRIMM'; 14 = 'SMD'; 15 = 'SSMP'
            16 = 'QFP'; 17 = 'TQFP'; 18 = 'SOIC'; 19 = 'LCC'; 20 = 'PLCC'
            21 = 'BGA'; 22 = 'FPBGA'; 23 = 'LGA'
        }

        $physMem  = @(Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction Stop)
        $memArray = @(Get-CimInstance -ClassName Win32_PhysicalMemoryArray -ErrorAction Stop)

        $totalSlots = 0
        foreach ($ma in $memArray) { $totalSlots += $ma.MemoryDevices }
        $usedSlots = $physMem.Count

        $dimms = foreach ($stick in $physMem) {
            $ff = if ($formFactorMap.ContainsKey([int]$stick.FormFactor)) { $formFactorMap[[int]$stick.FormFactor] } else { 'Unknown' }
            [PSCustomObject]@{
                Manufacturer = ($stick.Manufacturer -replace '^\s+|\s+$', '')
                PartNumber   = ($stick.PartNumber -replace '^\s+|\s+$', '')
                SerialNumber = ($stick.SerialNumber -replace '^\s+|\s+$', '')
                CapacityGB   = [math]::Round($stick.Capacity / 1GB, 2)
                Speed        = $stick.Speed
                FormFactor   = $ff
                BankLabel    = $stick.BankLabel
            }
        }

        $info['Memory'] = [PSCustomObject]@{
            TotalPhysicalMemoryGB = $totalGB
            AvailableMemoryGB     = $freeGB
            UsedMemoryGB          = $usedGB
            MemoryUsagePercent    = $usePct
            TotalSlots            = $totalSlots
            UsedSlots             = $usedSlots
            DIMMs                 = @($dimms)
        }
        Write-SysInfoLog 'Memory: collected.' -Level INFO
    }
    catch {
        Write-SysInfoLog "Memory collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Memory'] = [PSCustomObject]@{ Error = $_.Exception.Message }
    }

    # ── 5. Storage ──────────────────────────────────────────────────────────────
    & $reportProgress 'Collecting storage info...'
    try {
        $physicalDisks = @()
        if ($IsAdmin) {
            # MSFT_PhysicalDisk requires administrator privileges
            try {
                $msftDisks = @(Get-CimInstance -Namespace root/Microsoft/Windows/Storage `
                               -ClassName MSFT_PhysicalDisk -ErrorAction Stop)
                foreach ($d in $msftDisks) {
                    $physicalDisks += [PSCustomObject]@{
                        Model          = $d.FriendlyName
                        SerialNumber   = $d.SerialNumber
                        SizeGB         = [math]::Round($d.Size / 1GB, 2)
                        MediaType      = Get-DiskMediaType -MediaType $d.MediaType -BusType $d.BusType -Model $d.FriendlyName
                        InterfaceType  = switch ($d.BusType) {
                            1  { 'SCSI' }; 3  { 'ATA' }; 7  { 'USB' }
                            11 { 'SATA' }; 17 { 'NVMe' }; default { "Other ($($d.BusType))" }
                        }
                        Status         = $d.HealthStatus
                        PartitionCount = ($d | Get-CimAssociatedInstance -ResultClassName MSFT_Partition -ErrorAction SilentlyContinue | Measure-Object).Count
                    }
                }
                Write-SysInfoLog "Storage: collected $($physicalDisks.Count) physical disk(s) via MSFT_PhysicalDisk." -Level INFO
            }
            catch {
                Write-SysInfoLog "Storage: MSFT_PhysicalDisk query failed ($($_.Exception.Message)); falling back to Win32_DiskDrive." -Level WARN
                # Fallback to Win32_DiskDrive
                $wmiDisks = @(Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction Stop)
                foreach ($d in $wmiDisks) {
                    $physicalDisks += [PSCustomObject]@{
                        Model          = $d.Model
                        SerialNumber   = ($d.SerialNumber -replace '^\s+|\s+$', '')
                        SizeGB         = [math]::Round($d.Size / 1GB, 2)
                        MediaType      = Get-DiskMediaType -Model $d.Model
                        InterfaceType  = $d.InterfaceType
                        Status         = $d.Status
                        PartitionCount = $d.Partitions
                    }
                }
            }
        }
        else {
            # Not running as admin – use Win32_DiskDrive (available without elevation)
            Write-SysInfoLog 'Storage: not admin; using Win32_DiskDrive (MSFT_PhysicalDisk requires admin).' -Level INFO
            try {
                $wmiDisks = @(Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction Stop)
                foreach ($d in $wmiDisks) {
                    $physicalDisks += [PSCustomObject]@{
                        Model          = $d.Model
                        SerialNumber   = $script:NeedsAdminPriv
                        SizeGB         = [math]::Round($d.Size / 1GB, 2)
                        MediaType      = Get-DiskMediaType -Model $d.Model
                        InterfaceType  = $d.InterfaceType
                        Status         = $d.Status
                        PartitionCount = $d.Partitions
                    }
                }
            }
            catch {
                Write-SysInfoLog "Storage: Win32_DiskDrive query failed: $($_.Exception.Message)" -Level ERROR
                throw
            }
        }

        # Logical volumes
        $logicalVolumes = @()
        $vols = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop)
        foreach ($v in $vols) {
            $sizeGB = if ($v.Size -gt 0) { [math]::Round($v.Size / 1GB, 2) } else { 0 }
            $freeGB = if ($v.FreeSpace) { [math]::Round($v.FreeSpace / 1GB, 2) } else { 0 }
            $usedPct = if ($sizeGB -gt 0) { [math]::Round((($sizeGB - $freeGB) / $sizeGB) * 100, 1) } else { 0 }

            $logicalVolumes += [PSCustomObject]@{
                DriveLetter = $v.DeviceID
                Label       = $v.VolumeName
                FileSystem  = $v.FileSystem
                SizeGB      = $sizeGB
                FreeSpaceGB = $freeGB
                UsedPercent = $usedPct
            }
        }

        $info['Storage'] = [PSCustomObject]@{
            PhysicalDisks  = @($physicalDisks)
            LogicalVolumes = @($logicalVolumes)
        }
        Write-SysInfoLog 'Storage: collected.' -Level INFO
    }
    catch {
        Write-SysInfoLog "Storage collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Storage'] = [PSCustomObject]@{ Error = $_.Exception.Message }
    }

    # ── 6. Graphics / GPU ──────────────────────────────────────────────────────
    & $reportProgress 'Collecting graphics info...'
    try {
        $gpus = @(Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop)
        $gpuList = foreach ($g in $gpus) {
            $ramMB = if ($g.AdapterRAM -and $g.AdapterRAM -gt 0) {
                [math]::Round($g.AdapterRAM / 1MB, 0)
            } else { 0 }

            [PSCustomObject]@{
                Name              = $g.Name
                DriverVersion     = $g.DriverVersion
                AdapterRAMMB      = $ramMB
                VideoProcessor    = $g.VideoProcessor
                CurrentResolution = "$($g.CurrentHorizontalResolution)x$($g.CurrentVerticalResolution)"
                RefreshRate       = $g.CurrentRefreshRate
            }
        }

        $info['Graphics'] = @($gpuList)
        Write-SysInfoLog "Graphics: collected $($gpuList.Count) GPU(s)." -Level INFO
    }
    catch {
        Write-SysInfoLog "Graphics collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Graphics'] = @([PSCustomObject]@{ Error = $_.Exception.Message })
    }

    # ── 7. Network Adapters ────────────────────────────────────────────────────
    & $reportProgress 'Collecting network adapter info...'
    try {
        $adapters = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop)
        $adapterInfo = @(Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction Stop)

        $netList = foreach ($a in $adapters) {
            $parent = $adapterInfo | Where-Object { $_.Index -eq $a.Index } | Select-Object -First 1
            $speed  = if ($parent.Speed) { $parent.Speed } else { $null }

            [PSCustomObject]@{
                Name           = if ($parent) { $parent.NetConnectionID } else { $a.Description }
                Description    = $a.Description
                MACAddress     = $a.MACAddress
                Speed          = $speed
                Status         = if ($parent) { $parent.NetConnectionStatus } else { $null }
                IPv4Address    = ($a.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ', '
                SubnetMask     = ($a.IPSubnet | Select-Object -First 1)
                DefaultGateway = ($a.DefaultIPGateway -join ', ')
                DNSServers     = ($a.DNSServerSearchOrder -join ', ')
                DHCPEnabled    = $a.DHCPEnabled
                DHCPServer     = $a.DHCPServer
            }
        }

        $info['NetworkAdapters'] = @($netList)
        Write-SysInfoLog "NetworkAdapters: collected $($netList.Count) adapter(s)." -Level INFO
    }
    catch {
        Write-SysInfoLog "NetworkAdapters collection failed: $($_.Exception.Message)" -Level ERROR
        $info['NetworkAdapters'] = @([PSCustomObject]@{ Error = $_.Exception.Message })
    }

    # ── 8. BIOS ─────────────────────────────────────────────────────────────────
    & $reportProgress 'Collecting BIOS info...'
    try {
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop

        $info['BIOS'] = [PSCustomObject]@{
            BIOSManufacturer = $bios.Manufacturer
            BIOSVersion      = ($bios.BIOSVersion -join '; ')
            BIOSDate         = $bios.ReleaseDate
            SMBIOSVersion    = $bios.SMBIOSBIOSVersion
        }
        Write-SysInfoLog 'BIOS: collected.' -Level INFO
    }
    catch {
        Write-SysInfoLog "BIOS collection failed: $($_.Exception.Message)" -Level ERROR
        $info['BIOS'] = [PSCustomObject]@{ Error = $_.Exception.Message }
    }

    # ── 9. Motherboard ──────────────────────────────────────────────────────────
    & $reportProgress 'Collecting motherboard info...'
    try {
        $mb = Get-CimInstance -ClassName Win32_BaseBoard -ErrorAction Stop

        $info['Motherboard'] = [PSCustomObject]@{
            Manufacturer = $mb.Manufacturer
            Product      = $mb.Product
            Version      = $mb.Version
            SerialNumber = $mb.SerialNumber
        }
        Write-SysInfoLog 'Motherboard: collected.' -Level INFO
    }
    catch {
        Write-SysInfoLog "Motherboard collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Motherboard'] = [PSCustomObject]@{ Error = $_.Exception.Message }
    }

    # ── 10. Battery ─────────────────────────────────────────────────────────────
    & $reportProgress 'Collecting battery info...'
    try {
        $bat = @(Get-CimInstance -ClassName Win32_Battery -ErrorAction Stop)

        if ($bat.Count -gt 0) {
            $b = $bat[0]
            $designCap    = $b.DesignCapacity
            $fullChargeCap = $b.FullChargeCapacity
            $health = if ($designCap -and $designCap -gt 0 -and $fullChargeCap) {
                [math]::Round(($fullChargeCap / $designCap) * 100, 1)
            } else { $null }

            $info['Battery'] = [PSCustomObject]@{
                HasBattery          = $true
                Status              = $b.Status
                ChargePercent       = $b.EstimatedChargeRemaining
                EstimatedRuntime    = $b.EstimatedRunTime
                DesignCapacity      = $designCap
                FullChargeCapacity  = $fullChargeCap
                BatteryHealth       = $health
            }
            Write-SysInfoLog 'Battery: collected.' -Level INFO
        }
        else {
            $info['Battery'] = [PSCustomObject]@{ HasBattery = $false }
            Write-SysInfoLog 'Battery: no battery detected.' -Level INFO
        }
    }
    catch {
        Write-SysInfoLog "Battery collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Battery'] = [PSCustomObject]@{ HasBattery = $false; Error = $_.Exception.Message }
    }

    # ── 11. Installed Hotfixes & Startup Programs ───────────────────────────────
    & $reportProgress 'Collecting hotfixes and startup programs...'
    try {
        $hf = @(Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop)
        $info['Hotfixes'] = @(foreach ($h in $hf) {
            [PSCustomObject]@{
                HotfixID    = $h.HotFixID
                Description = $h.Description
                InstalledOn = $h.InstalledOn
            }
        })
        Write-SysInfoLog "Hotfixes: collected $($info['Hotfixes'].Count) hotfix(es)." -Level INFO
    }
    catch {
        Write-SysInfoLog "Hotfixes collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Hotfixes'] = @([PSCustomObject]@{ Error = $_.Exception.Message })
    }

    try {
        $startup = @(Get-CimInstance -ClassName Win32_StartupCommand -ErrorAction Stop)
        $info['StartupPrograms'] = @(foreach ($s in $startup) {
            [PSCustomObject]@{
                Name     = $s.Name
                Command  = $s.Command
                Location = $s.Location
            }
        })
        Write-SysInfoLog "StartupPrograms: collected $($info['StartupPrograms'].Count) entry(ies)." -Level INFO
    }
    catch {
        Write-SysInfoLog "StartupPrograms collection failed: $($_.Exception.Message)" -Level ERROR
        $info['StartupPrograms'] = @([PSCustomObject]@{ Error = $_.Exception.Message })
    }

    # ── 12. Security Information ────────────────────────────────────────────────
    & $reportProgress 'Collecting security info...'
    try {
        # Antivirus products (via Windows Security Center)
        # productState encoding (simplified): mask 0x1000 = product enabled;
        # mask 0x0010 = signatures out of date (0 = current, non-zero = outdated).
        $avProducts = @()
        try {
            $avRaw = @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop)
            foreach ($av in $avRaw) {
                $stateHex = '{0:X}' -f $av.productState
                $enabled  = ($av.productState -band 0x1000) -ne 0
                $upToDate = ($av.productState -band 0x0010) -eq 0
                $avProducts += [PSCustomObject]@{
                    Name      = $av.displayName
                    Enabled   = $enabled
                    UpToDate  = $upToDate
                    StateCode = $stateHex
                }
            }
        }
        catch {
            Write-SysInfoLog "Security: AntiVirusProduct query failed: $($_.Exception.Message)" -Level WARN
        }

        # Firewall products
        $fwProducts = @()
        try {
            $fwRaw = @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName FirewallProduct -ErrorAction Stop)
            foreach ($fw in $fwRaw) {
                $fwProducts += [PSCustomObject]@{
                    Name    = $fw.displayName
                    Enabled = ($fw.productState -band 0x1000) -ne 0
                }
            }
        }
        catch {
            Write-SysInfoLog "Security: FirewallProduct query failed: $($_.Exception.Message)" -Level WARN
        }

        # Windows Defender real-time protection status via registry
        $defenderRealTime = 'Unknown'
        try {
            $defReg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection' `
                                       -Name 'DisableRealtimeMonitoring' -ErrorAction Stop
            $defenderRealTime = if ($defReg.DisableRealtimeMonitoring -eq 0) { 'Enabled' } else { 'Disabled' }
        }
        catch {
            $defenderRealTime = 'Unknown'
        }

        # BitLocker status (requires admin for full details)
        $bitLockerStatus = 'Unknown'
        if ($IsAdmin) {
            try {
                $bl = Get-BitLockerVolume -ErrorAction Stop | Select-Object -First 1
                $bitLockerStatus = if ($bl) { "$($bl.VolumeStatus) ($($bl.ProtectionStatus))" } else { 'Not configured' }
            }
            catch {
                try {
                    # Fallback: query via manage-bde command line tool
                    $bdOut = & manage-bde -status C: 2>&1
                    if ($bdOut -match 'Protection On')  { $bitLockerStatus = 'Protection On' }
                    elseif ($bdOut -match 'Protection Off') { $bitLockerStatus = 'Protection Off' }
                    else { $bitLockerStatus = 'Unknown' }
                }
                catch { $bitLockerStatus = 'Unknown' }
            }
        }
        else {
            $bitLockerStatus = $script:NeedsAdminPriv
        }

        $info['Security'] = [PSCustomObject]@{
            AntiVirusProducts    = @($avProducts)
            FirewallProducts     = @($fwProducts)
            DefenderRealTime     = $defenderRealTime
            BitLockerStatus      = $bitLockerStatus
        }
        Write-SysInfoLog 'Security info collected.' -Level INFO
    }
    catch {
        Write-SysInfoLog "Security info collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Security'] = [PSCustomObject]@{ Error = $_.Exception.Message }
    }

    # ── 13. Environment / Locale ───────────────────────────────────────────────
    & $reportProgress 'Collecting environment info...'
    try {
        $tz      = [System.TimeZoneInfo]::Local
        $culture = [System.Globalization.CultureInfo]::CurrentCulture
        $uiCulture = [System.Globalization.CultureInfo]::CurrentUICulture

        $info['Environment'] = [PSCustomObject]@{
            TimeZone        = $tz.StandardName
            UTCOffset       = $tz.GetUtcOffset((Get-Date)).ToString()
            SystemLocale    = $culture.Name
            SystemLanguage  = $uiCulture.EnglishName
            OSLanguage      = $uiCulture.Name
            PowerShellVersion = $PSVersionTable.PSVersion.ToString()
            CLRVersion      = $PSVersionTable.CLRVersion.ToString()
            ExecutionPolicy = (Get-ExecutionPolicy).ToString()
        }
        Write-SysInfoLog 'Environment info collected.' -Level INFO
    }
    catch {
        Write-SysInfoLog "Environment info collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Environment'] = [PSCustomObject]@{ Error = $_.Exception.Message }
    }

    # ── 14. Display Monitors ───────────────────────────────────────────────────
    & $reportProgress 'Collecting display info...'
    try {
        $monitors = @(Get-CimInstance -ClassName Win32_DesktopMonitor -ErrorAction Stop)
        $displayList = foreach ($m in $monitors) {
            [PSCustomObject]@{
                Name             = $m.Name
                Manufacturer     = $m.MonitorManufacturer
                ScreenWidth      = $m.ScreenWidth
                ScreenHeight     = $m.ScreenHeight
                PixelsPerXLogicalInch = $m.PixelsPerXLogicalInch
                PixelsPerYLogicalInch = $m.PixelsPerYLogicalInch
                MonitorType      = $m.MonitorType
                Status           = $m.Status
            }
        }
        $info['Displays'] = @($displayList)
        Write-SysInfoLog "Displays: collected $($displayList.Count) monitor(s)." -Level INFO
    }
    catch {
        Write-SysInfoLog "Display info collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Displays'] = @([PSCustomObject]@{ Error = $_.Exception.Message })
    }

    # ── 15. Installed Software ─────────────────────────────────────────────────
    & $reportProgress 'Collecting installed software...'
    try {
        $regPaths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )

        $softwareHash = @{}  # Deduplicate by DisplayName+Version
        foreach ($path in $regPaths) {
            try {
                $items = @(Get-ItemProperty -Path $path -ErrorAction SilentlyContinue)
                foreach ($item in $items) {
                    if (-not [string]::IsNullOrWhiteSpace($item.DisplayName)) {
                        $key = "$($item.DisplayName)|$($item.DisplayVersion)"
                        if (-not $softwareHash.ContainsKey($key)) {
                            $softwareHash[$key] = [PSCustomObject]@{
                                Name        = $item.DisplayName
                                Version     = if ($item.DisplayVersion) { $item.DisplayVersion } else { '' }
                                Publisher   = if ($item.Publisher) { $item.Publisher } else { '' }
                                InstallDate = if ($item.InstallDate) { $item.InstallDate } else { '' }
                            }
                        }
                    }
                }
            }
            catch {
                Write-SysInfoLog "InstalledSoftware: could not read path $path : $($_.Exception.Message)" -Level WARN
            }
        }

        $info['InstalledSoftware'] = @($softwareHash.Values | Sort-Object Name)
        Write-SysInfoLog "InstalledSoftware: found $($info['InstalledSoftware'].Count) installed application(s)." -Level INFO
    }
    catch {
        Write-SysInfoLog "InstalledSoftware collection failed: $($_.Exception.Message)" -Level ERROR
        $info['InstalledSoftware'] = @([PSCustomObject]@{ Error = $_.Exception.Message })
    }

    # ── 16. Sound Devices ─────────────────────────────────────────────────────
    & $reportProgress 'Collecting sound device info...'
    try {
        $soundDevs = @(Get-CimInstance -ClassName Win32_SoundDevice -ErrorAction Stop)
        $soundList = foreach ($s in $soundDevs) {
            [PSCustomObject]@{
                Name         = $s.Name
                Manufacturer = $s.Manufacturer
                Status       = $s.Status
                DeviceID     = $s.DeviceID
            }
        }
        $info['SoundDevices'] = @($soundList)
        Write-SysInfoLog "SoundDevices: collected $($soundList.Count) device(s)." -Level INFO
    }
    catch {
        Write-SysInfoLog "Sound device collection failed: $($_.Exception.Message)" -Level ERROR
        $info['SoundDevices'] = @([PSCustomObject]@{ Error = $_.Exception.Message })
    }

    # ── 17. Printers ──────────────────────────────────────────────────────────
    & $reportProgress 'Collecting printer info...'
    try {
        $printers = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop)
        $printerList = foreach ($p in $printers) {
            [PSCustomObject]@{
                Name           = $p.Name
                PortName       = $p.PortName
                DriverName     = $p.DriverName
                Default        = $p.Default
                NetworkPrinter = $p.Network
                Status         = $p.PrinterStatus
                Shared         = $p.Shared
            }
        }
        $info['Printers'] = @($printerList)
        Write-SysInfoLog "Printers: collected $($printerList.Count) printer(s)." -Level INFO
    }
    catch {
        Write-SysInfoLog "Printer collection failed: $($_.Exception.Message)" -Level ERROR
        $info['Printers'] = @([PSCustomObject]@{ Error = $_.Exception.Message })
    }

    Write-Progress -Activity 'System Information Scan' -Completed

    return [PSCustomObject]$info
}

# ── Export Functions ────────────────────────────────────────────────────────────

function Export-SystemInfoTXT {
    <#
    .SYNOPSIS
        Writes a formatted plain-text system information report.
    .PARAMETER Data
        The PSCustomObject returned by Get-SystemInfoData.
    .PARAMETER Path
        File path for the output TXT file.
    .EXAMPLE
        Export-SystemInfoTXT -Data $data -Path 'C:\Reports\sysinfo.txt'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Data,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $sep  = '=' * 70
    $sep2 = '-' * 70
    $lines = [System.Collections.Generic.List[string]]::new()

    $lines.Add($sep)
    $lines.Add('  SYSTEM INFORMATION REPORT')
    $lines.Add("  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $lines.Add($sep)
    $lines.Add('')

    # Helper: write a section header
    $writeSection = {
        param([string]$Title)
        $lines.Add($sep2)
        $lines.Add("  $Title")
        $lines.Add($sep2)
    }

    # Helper: write key-value pairs from an object
    $writeProps = {
        param([PSCustomObject]$Obj, [string[]]$Exclude)
        foreach ($prop in $Obj.PSObject.Properties) {
            if ($Exclude -contains $prop.Name) { continue }
            if ($prop.Value -is [System.Collections.IEnumerable] -and $prop.Value -isnot [string]) { continue }
            $lines.Add("  {0,-30} : {1}" -f $prop.Name, $prop.Value)
        }
        $lines.Add('')
    }

    # 1. System Overview
    if ($Data.SystemOverview) {
        & $writeSection 'SYSTEM OVERVIEW'
        & $writeProps $Data.SystemOverview
    }

    # 2. Operating System
    if ($Data.OperatingSystem) {
        & $writeSection 'OPERATING SYSTEM'
        & $writeProps $Data.OperatingSystem
    }

    # 3. Processor
    if ($Data.Processor) {
        & $writeSection 'PROCESSOR'
        & $writeProps $Data.Processor
    }

    # 4. Memory
    if ($Data.Memory) {
        & $writeSection 'MEMORY (RAM)'
        & $writeProps $Data.Memory @('DIMMs')
        if ($Data.Memory.DIMMs) {
            $lines.Add("  Installed DIMMs:")
            $i = 1
            foreach ($dimm in $Data.Memory.DIMMs) {
                $lines.Add("    [$i] $($dimm.CapacityGB) GB  $($dimm.Manufacturer)  $($dimm.PartNumber)  Speed:$($dimm.Speed)  Form:$($dimm.FormFactor)  Bank:$($dimm.BankLabel)  S/N:$($dimm.SerialNumber)")
                $i++
            }
            $lines.Add('')
        }
    }

    # 5. Storage
    if ($Data.Storage) {
        & $writeSection 'STORAGE'
        if ($Data.Storage.PhysicalDisks) {
            $lines.Add('  Physical Disks:')
            $i = 1
            foreach ($disk in $Data.Storage.PhysicalDisks) {
                $lines.Add("    [$i] $($disk.Model)  $($disk.SizeGB) GB  Type:$($disk.MediaType)  Interface:$($disk.InterfaceType)  Status:$($disk.Status)  S/N:$($disk.SerialNumber)")
                $i++
            }
            $lines.Add('')
        }
        if ($Data.Storage.LogicalVolumes) {
            $lines.Add('  Logical Volumes:')
            foreach ($vol in $Data.Storage.LogicalVolumes) {
                $lines.Add("    $($vol.DriveLetter) [$($vol.Label)]  $($vol.FileSystem)  $($vol.SizeGB) GB total  $($vol.FreeSpaceGB) GB free  ($($vol.UsedPercent)% used)")
            }
            $lines.Add('')
        }
    }

    # 6. Graphics
    if ($Data.Graphics) {
        & $writeSection 'GRAPHICS / GPU'
        $i = 1
        foreach ($gpu in $Data.Graphics) {
            $lines.Add("    [$i] $($gpu.Name)")
            $lines.Add("        Driver: $($gpu.DriverVersion)  VRAM: $($gpu.AdapterRAMMB) MB  Resolution: $($gpu.CurrentResolution) @ $($gpu.RefreshRate) Hz")
            $i++
        }
        $lines.Add('')
    }

    # 7. Network
    if ($Data.NetworkAdapters) {
        & $writeSection 'NETWORK ADAPTERS'
        $i = 1
        foreach ($nic in $Data.NetworkAdapters) {
            $lines.Add("    [$i] $($nic.Name)  ($($nic.Description))")
            $lines.Add("        MAC: $($nic.MACAddress)  Speed: $($nic.Speed)")
            $lines.Add("        IPv4: $($nic.IPv4Address)  Subnet: $($nic.SubnetMask)  GW: $($nic.DefaultGateway)")
            $lines.Add("        DNS: $($nic.DNSServers)")
            $lines.Add("        DHCP: $($nic.DHCPEnabled)  Server: $($nic.DHCPServer)")
            $i++
        }
        $lines.Add('')
    }

    # 8. BIOS
    if ($Data.BIOS) {
        & $writeSection 'BIOS'
        & $writeProps $Data.BIOS
    }

    # 9. Motherboard
    if ($Data.Motherboard) {
        & $writeSection 'MOTHERBOARD'
        & $writeProps $Data.Motherboard
    }

    # 10. Battery
    if ($Data.Battery) {
        & $writeSection 'BATTERY'
        & $writeProps $Data.Battery
    }

    # 11. Hotfixes
    if ($Data.Hotfixes) {
        & $writeSection 'INSTALLED HOTFIXES'
        foreach ($hf in $Data.Hotfixes) {
            $lines.Add("    $($hf.HotfixID)  $($hf.Description)  Installed: $($hf.InstalledOn)")
        }
        $lines.Add('')
    }

    # 12. Startup Programs
    if ($Data.StartupPrograms) {
        & $writeSection 'STARTUP PROGRAMS'
        foreach ($sp in $Data.StartupPrograms) {
            $lines.Add("    $($sp.Name)  [$($sp.Location)]")
            $lines.Add("        $($sp.Command)")
        }
        $lines.Add('')
    }

    # 13. Security
    if ($Data.Security) {
        & $writeSection 'SECURITY'
        if ($Data.Security.Error) {
            $lines.Add("  Error: $($Data.Security.Error)")
            $lines.Add('')
        } else {
            $lines.Add("  Windows Defender Real-Time Protection: $($Data.Security.DefenderRealTime)")
            $lines.Add("  BitLocker Status                     : $($Data.Security.BitLockerStatus)")
            $lines.Add('')
            if ($Data.Security.AntiVirusProducts) {
                $lines.Add('  Antivirus Products:')
                foreach ($av in $Data.Security.AntiVirusProducts) {
                    $lines.Add("    - $($av.Name)  Enabled:$($av.Enabled)  UpToDate:$($av.UpToDate)")
                }
                $lines.Add('')
            }
            if ($Data.Security.FirewallProducts) {
                $lines.Add('  Firewall Products:')
                foreach ($fw in $Data.Security.FirewallProducts) {
                    $lines.Add("    - $($fw.Name)  Enabled:$($fw.Enabled)")
                }
                $lines.Add('')
            }
        }
    }

    # 14. Environment
    if ($Data.Environment) {
        & $writeSection 'ENVIRONMENT / LOCALE'
        & $writeProps $Data.Environment
    }

    # 15. Displays
    if ($Data.Displays) {
        & $writeSection 'DISPLAY MONITORS'
        $i = 1
        foreach ($disp in $Data.Displays) {
            $lines.Add("    [$i] $($disp.Name)  Resolution: $($disp.ScreenWidth)x$($disp.ScreenHeight)  Status: $($disp.Status)")
            $i++
        }
        $lines.Add('')
    }

    # 16. Installed Software
    if ($Data.InstalledSoftware) {
        & $writeSection 'INSTALLED SOFTWARE'
        foreach ($sw in $Data.InstalledSoftware) {
            $lines.Add("    $($sw.Name)  v$($sw.Version)  [$($sw.Publisher)]")
        }
        $lines.Add('')
    }

    # 17. Sound Devices
    if ($Data.SoundDevices) {
        & $writeSection 'SOUND DEVICES'
        foreach ($s in $Data.SoundDevices) {
            $lines.Add("    $($s.Name)  Manufacturer: $($s.Manufacturer)  Status: $($s.Status)")
        }
        $lines.Add('')
    }

    # 18. Printers
    if ($Data.Printers) {
        & $writeSection 'PRINTERS'
        foreach ($p in $Data.Printers) {
            $lines.Add("    $($p.Name)  Driver: $($p.DriverName)  Default: $($p.Default)  Network: $($p.NetworkPrinter)")
        }
        $lines.Add('')
    }

    $lines.Add($sep)
    $lines.Add('  End of Report')
    $lines.Add($sep)

    $lines | Out-File -FilePath $Path -Encoding UTF8 -Force
}

function Export-SystemInfoCSV {
    <#
    .SYNOPSIS
        Exports system information as a Category,Property,Value CSV file.
    .PARAMETER Data
        The PSCustomObject returned by Get-SystemInfoData.
    .PARAMETER Path
        File path for the output CSV file.
    .EXAMPLE
        Export-SystemInfoCSV -Data $data -Path 'C:\Reports\sysinfo.csv'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Data,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Helper: flatten an object into rows
    $addObject = {
        param([string]$Category, [PSCustomObject]$Obj, [string]$Prefix)
        foreach ($prop in $Obj.PSObject.Properties) {
            $val = $prop.Value
            if ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) { continue }
            $name = if ($Prefix) { "$Prefix.$($prop.Name)" } else { $prop.Name }
            $rows.Add([PSCustomObject]@{
                Category = $Category
                Property = $name
                Value    = "$val"
            })
        }
    }

    # Scalar sections
    $scalarSections = @(
        @{ Category = 'System Overview';  Obj = $Data.SystemOverview  }
        @{ Category = 'Operating System'; Obj = $Data.OperatingSystem }
        @{ Category = 'Processor';        Obj = $Data.Processor       }
        @{ Category = 'BIOS';             Obj = $Data.BIOS            }
        @{ Category = 'Motherboard';      Obj = $Data.Motherboard     }
        @{ Category = 'Battery';          Obj = $Data.Battery         }
    )

    foreach ($sec in $scalarSections) {
        if ($sec.Obj) { & $addObject $sec.Category $sec.Obj '' }
    }

    # Memory (scalar + DIMMs)
    if ($Data.Memory) {
        & $addObject 'Memory' $Data.Memory ''
        if ($Data.Memory.DIMMs) {
            $idx = 0
            foreach ($dimm in $Data.Memory.DIMMs) {
                & $addObject 'Memory' $dimm "DIMM[$idx]"
                $idx++
            }
        }
    }

    # Storage arrays
    if ($Data.Storage) {
        if ($Data.Storage.PhysicalDisks) {
            $idx = 0
            foreach ($disk in $Data.Storage.PhysicalDisks) {
                & $addObject 'Storage' $disk "PhysicalDisk[$idx]"
                $idx++
            }
        }
        if ($Data.Storage.LogicalVolumes) {
            $idx = 0
            foreach ($vol in $Data.Storage.LogicalVolumes) {
                & $addObject 'Storage' $vol "Volume[$idx]"
                $idx++
            }
        }
    }

    # Array sections
    $arraySections = @(
        @{ Category = 'Graphics';         Arr = $Data.Graphics;        Prefix = 'GPU'     }
        @{ Category = 'Network Adapters'; Arr = $Data.NetworkAdapters; Prefix = 'Adapter' }
        @{ Category = 'Hotfixes';         Arr = $Data.Hotfixes;        Prefix = 'Hotfix'  }
        @{ Category = 'Startup Programs'; Arr = $Data.StartupPrograms; Prefix = 'Startup' }
    )

    foreach ($sec in $arraySections) {
        if ($sec.Arr) {
            $idx = 0
            foreach ($item in $sec.Arr) {
                & $addObject $sec.Category $item "$($sec.Prefix)[$idx]"
                $idx++
            }
        }
    }

    # Security (scalar + nested arrays)
    if ($Data.Security) {
        $rows.Add([PSCustomObject]@{ Category = 'Security'; Property = 'DefenderRealTime'; Value = "$($Data.Security.DefenderRealTime)" })
        $rows.Add([PSCustomObject]@{ Category = 'Security'; Property = 'BitLockerStatus';  Value = "$($Data.Security.BitLockerStatus)" })
        if ($Data.Security.AntiVirusProducts) {
            $idx = 0
            foreach ($av in $Data.Security.AntiVirusProducts) {
                & $addObject 'Security' $av "AV[$idx]"; $idx++
            }
        }
        if ($Data.Security.FirewallProducts) {
            $idx = 0
            foreach ($fw in $Data.Security.FirewallProducts) {
                & $addObject 'Security' $fw "FW[$idx]"; $idx++
            }
        }
    }

    # Environment
    if ($Data.Environment) { & $addObject 'Environment' $Data.Environment '' }

    # Displays
    if ($Data.Displays) {
        $idx = 0
        foreach ($disp in $Data.Displays) { & $addObject 'Displays' $disp "Display[$idx]"; $idx++ }
    }

    # InstalledSoftware
    if ($Data.InstalledSoftware) {
        $idx = 0
        foreach ($sw in $Data.InstalledSoftware) { & $addObject 'InstalledSoftware' $sw "SW[$idx]"; $idx++ }
    }

    # SoundDevices
    if ($Data.SoundDevices) {
        $idx = 0
        foreach ($s in $Data.SoundDevices) { & $addObject 'SoundDevices' $s "Sound[$idx]"; $idx++ }
    }

    # Printers
    if ($Data.Printers) {
        $idx = 0
        foreach ($p in $Data.Printers) { & $addObject 'Printers' $p "Printer[$idx]"; $idx++ }
    }

    $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
}

function Format-SystemInfoTable {
    <#
    .SYNOPSIS
        Returns formatted string blocks suitable for console display with box-drawing characters.
    .PARAMETER Data
        The PSCustomObject returned by Get-SystemInfoData.
    .OUTPUTS
        [string] Multi-line formatted text.
    .EXAMPLE
        Write-Host (Format-SystemInfoTable -Data $data)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Data
    )

    $sb = [System.Text.StringBuilder]::new()

    # Box-drawing characters
    $tl = [char]0x250C  # ┌
    $tr = [char]0x2510  # ┐
    $bl = [char]0x2514  # └
    $br = [char]0x2518  # ┘
    $hz = [char]0x2500  # ─
    $vt = [char]0x2502  # │
    $lj = [char]0x251C  # ├
    $rj = [char]0x2524  # ┤

    $boxWidth = 72
    $innerW   = $boxWidth - 2  # space inside vertical bars

    $hzLine    = "$hz" * $innerW
    $topBorder = "$tl$hzLine$tr"
    $botBorder = "$bl$hzLine$br"
    $midBorder = "$lj$hzLine$rj"

    # Helper: title line centred
    $titleLine = {
        param([string]$Title)
        $pad = $innerW - $Title.Length
        $left  = [math]::Floor($pad / 2)
        $right = $pad - $left
        "$vt$(' ' * $left)$Title$(' ' * $right)$vt"
    }

    # Helper: key-value line
    $kvLine = {
        param([string]$Key, [string]$Value)
        $content = "  {0,-28} {1}" -f "${Key}:", $Value
        if ($content.Length -gt $innerW) { $content = $content.Substring(0, $innerW) }
        $pad = $innerW - $content.Length
        "$vt$content$(' ' * $pad)$vt"
    }

    # Helper: plain text line
    $textLine = {
        param([string]$Text)
        if ($Text.Length -gt $innerW) { $Text = $Text.Substring(0, $innerW) }
        $pad = $innerW - $Text.Length
        "$vt$Text$(' ' * $pad)$vt"
    }

    # Helper: render a section
    $renderSection = {
        param([string]$Title, [scriptblock]$ContentBlock)
        [void]$sb.AppendLine($topBorder)
        [void]$sb.AppendLine((& $titleLine $Title))
        [void]$sb.AppendLine($midBorder)
        & $ContentBlock
        [void]$sb.AppendLine($botBorder)
        [void]$sb.AppendLine('')
    }

    $renderProps = {
        param([PSCustomObject]$Obj, [string[]]$Exclude)
        foreach ($prop in $Obj.PSObject.Properties) {
            if ($Exclude -contains $prop.Name) { continue }
            if ($prop.Value -is [System.Collections.IEnumerable] -and $prop.Value -isnot [string]) { continue }
            [void]$sb.AppendLine((& $kvLine $prop.Name "$($prop.Value)"))
        }
    }

    # 1. System Overview
    if ($Data.SystemOverview) {
        & $renderSection 'SYSTEM OVERVIEW' { & $renderProps $Data.SystemOverview }
    }

    # 2. Operating System
    if ($Data.OperatingSystem) {
        & $renderSection 'OPERATING SYSTEM' { & $renderProps $Data.OperatingSystem }
    }

    # 3. Processor
    if ($Data.Processor) {
        & $renderSection 'PROCESSOR' { & $renderProps $Data.Processor }
    }

    # 4. Memory
    if ($Data.Memory) {
        & $renderSection 'MEMORY (RAM)' {
            & $renderProps $Data.Memory @('DIMMs')
            if ($Data.Memory.DIMMs -and $Data.Memory.DIMMs.Count -gt 0) {
                [void]$sb.AppendLine($midBorder)
                [void]$sb.AppendLine((& $textLine '  Installed DIMMs:'))
                $i = 1
                foreach ($dimm in $Data.Memory.DIMMs) {
                    [void]$sb.AppendLine((& $textLine "   [$i] $($dimm.CapacityGB)GB $($dimm.Manufacturer) Speed:$($dimm.Speed)"))
                    $i++
                }
            }
        }
    }

    # 5. Storage
    if ($Data.Storage) {
        & $renderSection 'STORAGE' {
            if ($Data.Storage.PhysicalDisks) {
                [void]$sb.AppendLine((& $textLine '  Physical Disks:'))
                $i = 1
                foreach ($disk in $Data.Storage.PhysicalDisks) {
                    [void]$sb.AppendLine((& $textLine "   [$i] $($disk.Model) $($disk.SizeGB)GB $($disk.MediaType)"))
                    $i++
                }
            }
            if ($Data.Storage.LogicalVolumes) {
                [void]$sb.AppendLine((& $textLine '  Logical Volumes:'))
                foreach ($vol in $Data.Storage.LogicalVolumes) {
                    [void]$sb.AppendLine((& $textLine "   $($vol.DriveLetter) $($vol.SizeGB)GB ($($vol.UsedPercent)% used)"))
                }
            }
        }
    }

    # 6. Graphics
    if ($Data.Graphics) {
        & $renderSection 'GRAPHICS / GPU' {
            $i = 1
            foreach ($gpu in $Data.Graphics) {
                [void]$sb.AppendLine((& $textLine "   [$i] $($gpu.Name)"))
                [void]$sb.AppendLine((& $textLine "       VRAM: $($gpu.AdapterRAMMB)MB  $($gpu.CurrentResolution)"))
                $i++
            }
        }
    }

    # 7. Network
    if ($Data.NetworkAdapters) {
        & $renderSection 'NETWORK ADAPTERS' {
            $i = 1
            foreach ($nic in $Data.NetworkAdapters) {
                [void]$sb.AppendLine((& $textLine "   [$i] $($nic.Name)"))
                [void]$sb.AppendLine((& $textLine "       IP: $($nic.IPv4Address)  MAC: $($nic.MACAddress)"))
                $i++
            }
        }
    }

    # 8. BIOS
    if ($Data.BIOS) {
        & $renderSection 'BIOS' { & $renderProps $Data.BIOS }
    }

    # 9. Motherboard
    if ($Data.Motherboard) {
        & $renderSection 'MOTHERBOARD' { & $renderProps $Data.Motherboard }
    }

    # 10. Battery
    if ($Data.Battery) {
        & $renderSection 'BATTERY' { & $renderProps $Data.Battery }
    }

    # 11. Security
    if ($Data.Security -and -not $Data.Security.Error) {
        & $renderSection 'SECURITY' {
            [void]$sb.AppendLine((& $kvLine 'Defender Real-Time' "$($Data.Security.DefenderRealTime)"))
            [void]$sb.AppendLine((& $kvLine 'BitLocker Status' "$($Data.Security.BitLockerStatus)"))
            if ($Data.Security.AntiVirusProducts) {
                [void]$sb.AppendLine((& $textLine '  Antivirus:'))
                foreach ($av in $Data.Security.AntiVirusProducts) {
                    [void]$sb.AppendLine((& $textLine "   - $($av.Name) [Enabled:$($av.Enabled)]"))
                }
            }
        }
    }

    # 12. Environment
    if ($Data.Environment) {
        & $renderSection 'ENVIRONMENT' { & $renderProps $Data.Environment }
    }

    # 13. Displays
    if ($Data.Displays) {
        & $renderSection 'DISPLAY MONITORS' {
            $i = 1
            foreach ($disp in $Data.Displays) {
                [void]$sb.AppendLine((& $textLine "   [$i] $($disp.Name) $($disp.ScreenWidth)x$($disp.ScreenHeight)"))
                $i++
            }
        }
    }

    # 14. Sound Devices
    if ($Data.SoundDevices) {
        & $renderSection 'SOUND DEVICES' {
            $i = 1
            foreach ($s in $Data.SoundDevices) {
                [void]$sb.AppendLine((& $textLine "   [$i] $($s.Name) [$($s.Status)]"))
                $i++
            }
        }
    }

    # 15. Printers
    if ($Data.Printers) {
        & $renderSection 'PRINTERS' {
            $i = 1
            foreach ($p in $Data.Printers) {
                [void]$sb.AppendLine((& $textLine "   [$i] $($p.Name) [Default:$($p.Default)]"))
                $i++
            }
        }
    }

    return $sb.ToString()
}

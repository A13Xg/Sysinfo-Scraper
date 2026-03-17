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
        [scriptblock]$ProgressCallback
    )

    # Helper to report progress
    $totalSteps = 11
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
    }
    catch {
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
    }
    catch {
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
    }
    catch {
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
    }
    catch {
        $info['Memory'] = [PSCustomObject]@{ Error = $_.Exception.Message }
    }

    # ── 5. Storage ──────────────────────────────────────────────────────────────
    & $reportProgress 'Collecting storage info...'
    try {
        # Physical disks via Storage namespace (preferred) with fallback
        $physicalDisks = @()
        try {
            $msftDisks = @(Get-CimInstance -Namespace root/Microsoft/Windows/Storage -ClassName MSFT_PhysicalDisk -ErrorAction Stop)
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
        }
        catch {
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
    }
    catch {
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
    }
    catch {
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
    }
    catch {
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
    }
    catch {
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
    }
    catch {
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
        }
        else {
            $info['Battery'] = [PSCustomObject]@{ HasBattery = $false }
        }
    }
    catch {
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
    }
    catch {
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
    }
    catch {
        $info['StartupPrograms'] = @([PSCustomObject]@{ Error = $_.Exception.Message })
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

    return $sb.ToString()
}

# Sysinfo-Scraper

A system information collection tool for Windows. Started as a small batch script and now includes comprehensive PowerShell scripts with CLI and GUI interfaces.

## Scripts

| Script | Description |
|---|---|
| `Scrape Info.bat` | Legacy batch script — collects basic systeminfo and ipconfig output to a text file |
| `SysInfo-CLI.ps1` | **PowerShell CLI** — interactive menu-driven interface with verbose progress, stylized console table output, and TXT/CSV export |
| `SysInfo-GUI.ps1` | **PowerShell GUI** — WPF-based graphical interface with tabbed data display, hardware image support, status bar, and export options |
| `SysInfo-Core.ps1` | Shared core module — dot-sourced by both CLI and GUI scripts; contains all data collection, export, and utility functions |

## Requirements

- **Windows** with PowerShell 5.1 or later
- Administrator privileges are recommended for full data access but not required

## Quick Start

### CLI (Interactive)
```powershell
.\SysInfo-CLI.ps1
```

### CLI (Non-Interactive / Headless)
```powershell
.\SysInfo-CLI.ps1 -ScanAndExport -OutputFormat Both -OutputPath C:\Reports
```

### GUI
```powershell
.\SysInfo-GUI.ps1
```

Or right-click `SysInfo-GUI.ps1` → **Run with PowerShell**.

### Legacy Batch
```
"Scrape Info.bat"
```

## Data Collected (v2.0)

The PowerShell scripts collect **11 categories** of system information:

1. **System Overview** — Computer name, user, domain, manufacturer, model, serial number, asset tag, chassis type
2. **Operating System** — Name, version, build, architecture, install date, last boot, uptime
3. **Processor** — Name, cores, threads, clock speeds, cache sizes, socket
4. **Memory (RAM)** — Total/available/used, usage percent, slot count, individual DIMM details
5. **Storage** — Physical disks (model, size, type SSD/HDD/NVMe, health) and logical volumes
6. **Graphics / GPU** — Name, driver, VRAM, resolution, refresh rate
7. **Network Adapters** — Name, MAC, IP, gateway, DNS, DHCP configuration
8. **BIOS** — Manufacturer, version, date, SMBIOS version
9. **Motherboard** — Manufacturer, product, version, serial number
10. **Battery** — Status, charge, health, capacity (if present)
11. **Hotfixes & Startup Programs**

## Hardware Images (GUI)

The GUI displays a product image for the detected hardware. Place PNG images in the `hardwareImages/` directory. The image resolver uses a four-tier fallback:

1. Exact model match (e.g., `DellOptiplex7080.png`)
2. Manufacturer + type generic (e.g., `DellGenericLaptop.png`)
3. Manufacturer generic (e.g., `DellGeneric.png`)
4. Global fallback (`GenericComputer.png`)

See [`hardwareImages/README.md`](hardwareImages/README.md) for the full list of supported models and required images.

## Export Formats

- **TXT** — Formatted plain-text report with headers and separators
- **CSV** — Structured `Category, Property, Value` format for spreadsheet import

Files are named `SysInfo_{ComputerName}_{Timestamp}.txt/csv`.

# — @13X —




## License Information: ##

This repository and these scripts are brought to you and the general public under the 'Creative Commons > Attribution-NonCommercial-ShareAlike 4.0 International License' (CC BY-NC-SA 4.0)


You are free to:

    Share — copy and redistribute the material in any medium or format
    Adapt — remix, transform, and build upon the material

    The licensor cannot revoke these freedoms as long as you follow the license terms.

Under the following terms:

    Attribution — You must give appropriate credit, provide a link to the license, and indicate if changes were made. You may do so in any reasonable manner, but not in any way that suggests the licensor endorses you or your use.

    NonCommercial — You may not use the material for commercial purposes.

    ShareAlike — If you remix, transform, or build upon the material, you must distribute your contributions under the same license as the original.

    No additional restrictions — You may not apply legal terms or technological measures that legally restrict others from doing anything the license permits.



Creative Commons Corporation (“Creative Commons”) is not a law firm and does not provide legal services or legal advice. Distribution of Creative Commons public licenses does not create a lawyer-client or other relationship. Creative Commons makes its licenses and related information available on an “as-is” basis. Creative Commons gives no warranties regarding its licenses, any material licensed under their terms and conditions, or any related information. Creative Commons disclaims all liability for damages resulting from their use to the fullest extent possible.

Using Creative Commons Public Licenses

Creative Commons public licenses provide a standard set of terms and conditions that creators and other rights holders may use to share original works of authorship and other material subject to copyright and certain other rights specified in the public license below. The following considerations are for informational purposes only, are not exhaustive, and do not form part of our licenses.

    Considerations for licensors: Our public licenses are intended for use by those authorized to give the public permission to use material in ways otherwise restricted by copyright and certain other rights. Our licenses are irrevocable. Licensors should read and understand the terms and conditions of the license they choose before applying it. Licensors should also secure all rights necessary before applying our licenses so that the public can reuse the material as expected. Licensors should clearly mark any material not subject to the license. This includes other CC-licensed material, or material used under an exception or limitation to copyright. More considerations for licensors.

    Considerations for the public: By using one of our public licenses, a licensor grants the public permission to use the licensed material under specified terms and conditions. If the licensor’s permission is not necessary for any reason–for example, because of any applicable exception or limitation to copyright–then that use is not regulated by the license. Our licenses grant only permissions under copyright and certain other rights that a licensor has authority to grant. Use of the licensed material may still be restricted for other reasons, including because others have copyright or other rights in the material. A licensor may make special requests, such as asking that all changes be marked or described. Although not required by our licenses, you are encouraged to respect those requests where reasonable. More considerations for the public.

### MD5 Hashes: ###

 > v0.2 'MD5: 6214a9aac4773c8f0863113793e978b3'

 > v0.3 'MD5: fc3e11101f33b540274f819903bda4c2'

 > v0.4 'MD5: c895090b636384d877b3e4db50defc2a'

# Hardware Images Directory

This directory contains hardware product images used by the **SysInfo-GUI** to display a visual representation of the detected system.

## Fallback Logic

The image resolver (`Get-HardwareImagePath` in `SysInfo-Core.ps1`) uses a four-tier fallback:

1. **Exact model match** — `{Make}{Model}.png` (e.g., `DellOptiplex7080.png`)
2. **Type-generic match** — `{Make}Generic{Type}.png` (e.g., `DellGenericLaptop.png`, `HPGenericDesktop.png`)
3. **Make-generic match** — `{Make}Generic.png` (e.g., `DellGeneric.png` — just the manufacturer logo)
4. **Global fallback** — `GenericComputer.png`

All filenames are normalized: spaces and special characters are stripped before lookup.  
Images should be **PNG** format, ideally **200×200 pixels** or larger (the GUI scales to fit).

---

## Required Generic Images

At minimum, place these files to cover the fallback tiers:

| Filename | Description |
|---|---|
| `GenericComputer.png` | Universal fallback — generic computer icon |
| **Dell** | |
| `DellGeneric.png` | Dell logo |
| `DellGenericDesktop.png` | Generic Dell desktop |
| `DellGenericLaptop.png` | Generic Dell laptop |
| **HP** | |
| `HPGeneric.png` | HP logo |
| `HPGenericDesktop.png` | Generic HP desktop |
| `HPGenericLaptop.png` | Generic HP laptop |
| **Lenovo** | |
| `LenovoGeneric.png` | Lenovo logo |
| `LenovoGenericDesktop.png` | Generic Lenovo desktop |
| `LenovoGenericLaptop.png` | Generic Lenovo laptop |
| **Acer** | |
| `AcerGeneric.png` | Acer logo |
| `AcerGenericDesktop.png` | Generic Acer desktop |
| `AcerGenericLaptop.png` | Generic Acer laptop |
| **ASUS** | |
| `ASUSGeneric.png` | ASUS logo |
| `ASUSGenericDesktop.png` | Generic ASUS desktop |
| `ASUSGenericLaptop.png` | Generic ASUS laptop |
| **MSI** | |
| `MSIGeneric.png` | MSI logo |
| `MSIGenericDesktop.png` | Generic MSI desktop |
| `MSIGenericLaptop.png` | Generic MSI laptop |
| **Apple** | |
| `AppleGeneric.png` | Apple logo |
| `AppleGenericDesktop.png` | Generic Apple desktop |
| `AppleGenericLaptop.png` | Generic Apple laptop |
| **Microsoft** | |
| `MicrosoftGeneric.png` | Microsoft logo |
| `MicrosoftGenericDesktop.png` | Generic Microsoft desktop |
| `MicrosoftGenericLaptop.png` | Generic Microsoft laptop |

---

## Supported Exact-Model Images

Place any of these for an exact match (see `$script:KnownModels` in `SysInfo-Core.ps1`):

### Dell
- `DellOptiplex3080.png`
- `DellOptiplex5080.png`
- `DellOptiplex7080.png`
- `DellOptiplex7090.png`
- `DellOptiplex7010.png`
- `DellLatitude5520.png`
- `DellLatitude5530.png`
- `DellLatitude5540.png`
- `DellLatitude7420.png`
- `DellLatitude7430.png`
- `DellPrecision3660.png`
- `DellPrecision5570.png`
- `DellXPS139310.png`
- `DellXPS159520.png`
- `DellInspiron155510.png`

### HP
- `HPEliteDesk800G6.png`
- `HPEliteDesk800G8.png`
- `HPEliteDesk800G9.png`
- `HPProDesk400G7.png`
- `HPProDesk600G6.png`
- `HPEliteBook840G8.png`
- `HPEliteBook840G9.png`
- `HPEliteBook850G8.png`
- `HPProBook450G8.png`
- `HPProBook450G9.png`
- `HPZBook15G8.png`
- `HPZBookFury16G9.png`
- `HPEliteDragonfly.png`
- `HPZ2TowerG9.png`
- `HPZ4G5.png`

### Lenovo
- `LenovoThinkPadT14Gen3.png`
- `LenovoThinkPadT14sGen3.png`
- `LenovoThinkPadX1CarbonGen10.png`
- `LenovoThinkPadX1CarbonGen11.png`
- `LenovoThinkPadL14Gen3.png`
- `LenovoThinkPadE14Gen4.png`
- `LenovoThinkCentreM70q.png`
- `LenovoThinkCentreM90q.png`
- `LenovoThinkCentreM720s.png`
- `LenovoThinkStationP340.png`
- `LenovoThinkStationP360.png`
- `LenovoIdeaPad5Pro.png`
- `LenovoLegion5Pro.png`
- `LenovoYoga9i.png`
- `LenovoIdeaCentre5.png`

### Acer
- `AcerAspire5.png`
- `AcerAspire7.png`
- `AcerSwift3.png`
- `AcerSwift5.png`
- `AcerSpin5.png`
- `AcerNitro5.png`
- `AcerPredatorHelios300.png`
- `AcerTravelMateP2.png`
- `AcerTravelMateP6.png`
- `AcerVeritonN4680G.png`
- `AcerVeritonX4680G.png`
- `AcerConceptD3.png`

### ASUS
- `ASUSZenBook14.png`
- `ASUSZenBook13.png`
- `ASUSVivoBook15.png`
- `ASUSVivoBookS15.png`
- `ASUSROGZEPHYRUSG14.png`
- `ASUSROGZEPHYRUSG15.png`
- `ASUSROGStrixG15.png`
- `ASUSTUFGamingF15.png`
- `ASUSProArtStudioBook16.png`
- `ASUSExpertBookB5.png`
- `ASUSExpertBookB9.png`
- `ASUSChromebookFlip.png`

### MSI
- `MSIPrestige14.png`
- `MSIPrestige15.png`
- `MSIModern14.png`
- `MSIModern15.png`
- `MSIGS66Stealth.png`
- `MSIGS76Stealth.png`
- `MSIGE76Raider.png`
- `MSICreator15.png`
- `MSICreatorZ16.png`
- `MSISummitE16Flip.png`
- `MSITridentX.png`
- `MSIInfiniteRS.png`

### Apple
- `AppleMacBookPro14.png`
- `AppleMacBookPro16.png`
- `AppleMacBookAir13.png`
- `AppleMacBookAir15.png`
- `AppleiMac24.png`
- `AppleMacMini.png`
- `AppleMacStudio.png`
- `AppleMacPro.png`
- `AppleiMacPro.png`
- `AppleMacBookPro13.png`
- `AppleMacBookAir.png`

### Microsoft Surface
- `MicrosoftSurfacePro9.png`
- `MicrosoftSurfacePro8.png`
- `MicrosoftSurfacePro7.png`
- `MicrosoftSurfaceLaptop5.png`
- `MicrosoftSurfaceLaptop4.png`
- `MicrosoftSurfaceLaptopStudio.png`
- `MicrosoftSurfaceLaptopGo2.png`
- `MicrosoftSurfaceBook3.png`
- `MicrosoftSurfaceGo3.png`
- `MicrosoftSurfaceStudio2.png`
- `MicrosoftSurfaceHub2S.png`

---

## Adding New Models

To add support for a new model:

1. Add the normalized key → filename mapping to `$script:KnownModels` in `SysInfo-Core.ps1`
2. Place the corresponding PNG image in this directory
3. The key format is `{Make}{Model}` with all spaces and special characters removed

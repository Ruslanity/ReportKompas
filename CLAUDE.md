# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**ReportKompas** is a Windows Forms desktop application for integrating with KOMPAS-3D CAD software. It generates reports about assembly structures, processes DXF files for laser cutting calculations, and manages material/equipment coatings.

**Technology Stack:**
- C# .NET Framework 4.0
- Windows Forms (WinForms)
- COM Interop with KOMPAS-3D API (versions 5 and 7)
- Excel generation via ClosedXML
- DXF processing via netDxf library

## External API Documentation

### KOMPAS-3D SDK Documentation
**Official ASCON Documentation**: https://help.ascon.ru/KOMPAS_SDK/22/ru-RU/index.html

This project uses COM Interop with KOMPAS-3D API (both v5 and v7). Key sections:
- **API v5 (Kompas6API5)**: Used for core COM integration in [KompasReport.cs](KompasReport.cs)
- **API v7 (KompasAPI7)**: Referenced for newer functionality
- **COM Constants**: Kompas6Constants and Kompas6Constants3D provide enumeration values

### Related Documentation
- **netDxf Library**: Used in [Классы обработки DXF\DxfProcessor.cs](Классы обработки DXF\DxfProcessor.cs) for DXF file parsing
- **ClosedXML**: Used for Excel report generation

### Quick Reference
When working with KOMPAS-3D integration:
1. COM registration patterns → See `KompasReport.RegisterKompasLib()` and Windows Registry setup
2. License validation → See `KompasReport.ExternalRunCommand()`
3. Assembly tree structure → See [ObjectAssemblyKompas.cs](Классы структуры данных\ObjectAssemblyKompas.cs) hierarchical data model
4. DXF processing → See [DxfProcessor.cs](Классы обработки DXF\DxfProcessor.cs) with speed calculations from `SpeedCut.xml`

## Building and Running

### Build Commands
```bash
# Build Debug version (includes full debug symbols, COM registration)
msbuild ReportKompas.sln /p:Configuration=Debug

# Build Release version (optimized, COM registration enabled)
msbuild ReportKompas.sln /p:Configuration=Release
```

### Output Locations
- Debug builds: `bin\Debug\ReportKompas.exe`
- Release builds: `bin\Release\ReportKompas.exe`

### COM Registration
The application registers itself for COM interop automatically during build. Both Debug and Release configurations have `RegisterForComInterop` enabled, which allows KOMPAS-3D to load the application as a library.

To manually register/unregister:
```bash
# Register for COM
regasm /codebase bin\Debug\ReportKompas.exe

# Unregister
regasm /u bin\Debug\ReportKompas.exe
```

### Running the Application
The application can be launched in two ways:
1. **Standalone**: Run `ReportKompas.exe` directly (requires license file `\bug` in application directory)
2. **From KOMPAS-3D**: Load as an external library through KOMPAS-3D's library manager

## Architecture

### Core Data Model: Tree Structure

The application was restructured (commit `5e67ce0`) from a flat list to a **tree-based architecture**. The central data structure is `ObjectAssemblyKompas`, which represents assembly components in a hierarchical parent-child relationship.

**Key class: `ObjectAssemblyKompas`** ([Классы структуры данных\ObjectAssemblyKompas.cs](Классы структуры данных\ObjectAssemblyKompas.cs))
- Represents a single component (assembly, detail, standard part, etc.)
- Contains properties: Designation, Name, Quantity, Material, Mass, Dimensions, DXF paths, coating info, etc.
- Tree relationships: `ParentK` (parent reference), `Children` (list of child components)
- Important methods:
  - `AddChild()`: Adds child components; automatically merges duplicates by increasing `Quantity`
  - `FindChild()`: Recursive search through tree by designation/name
  - `SortChildrenBySpecificationSection()`: Sorts children by specification section ("Сборочные единицы", "Детали", "Стандартные изделия", "Прочие изделия")
  - `ReplaceMaterial()`: Recursively processes and updates material fields based on designation patterns (e.g., "1.5mm_Zn", "1.5mm_Aisi")

### Main Components

**1. Main UI Form: `ReportKompas.cs`** ([ReportKompas.cs](ReportKompas.cs))
- Singleton pattern: `GetInstance()` returns shared instance
- Uses `ObjectListView` (TreeListView) to display hierarchical assembly structure
- Handles KOMPAS-3D integration, DXF processing, report generation
- Primary entry point for user interactions

**2. COM Integration: `KompasReport.cs`** ([KompasReport.cs](KompasReport.cs))
- COM-callable wrapper for KOMPAS-3D integration
- `ExternalRunCommand()`: Entry point called by KOMPAS-3D
- License verification: Validates license file (`\bug`) against machine fingerprint (Username + OS + MachineName)
- COM registration/unregistration functions for Windows Registry

**3. DXF Processing: `DxfProcessor.cs`** ([Классы обработки DXF\DxfProcessor.cs](Классы обработки DXF\DxfProcessor.cs))
- Parses DXF files using netDxf library
- Calculates:
  - Overall dimensions (length, width, height)
  - Cutting path lengths (segments, arcs, circles)
  - Laser cutting times based on material/thickness from `SpeedCut.xml`
  - Engraving and idle times
- Returns `DimensionsDXF` objects with calculated data

**4. Coating Management: `Coating.cs`** ([Классы покрытий\Coating.cs](Классы покрытий\Coating.cs))
- UI form for managing surface treatments
- Tracks coating type and coverage area
- Integrates with tree structure to update component coating properties

**5. Settings Management** ([Классы настроек\](Классы настроек\))
- `Settings.cs`: Main XML serializer for application configuration
- `SettingsForm.cs`: UI for user settings
- `CodeEquip.cs`: Equipment code mappings (serializes to `Settings\CodeEquip.xml`)
- `CodeMaterial.cs`: Material code mappings (serializes to `Settings\CodeMaterial.xml`)
- `SpeedCut.cs`: Laser cutting speed dictionary (serializes to `Settings\SpeedCut.xml`)

### Configuration Files (Settings\ directory)

All configuration files are XML-based and copied to output directory during build:
- **Settings.xml**: Application paths, feature flags, user preferences
- **CodeEquip.xml**: Equipment type codes and descriptions (34KB database)
- **CodeMaterial.xml**: Material type codes (17KB database)
- **SpeedCut.xml**: Laser cutting speed parameters indexed by material/thickness
- **Сolumns.xml**: UI column definitions for TreeListView

### Data Flow

```
KOMPAS-3D → KompasReport (COM) → ReportKompas (Singleton UI)
                                      ↓
                                  ObjectAssemblyKompas (Tree Structure)
                                      ↓
                     ┌────────────────┼────────────────┐
                     ↓                ↓                ↓
              DxfProcessor    Settings/Dictionaries   Coating
                     ↓                ↓                ↓
                     └────────────────┴────────────────┘
                                      ↓
                              Report/Excel Export
```

### Specification Sections

Components are categorized into specification sections (`SpecificationSection` property):
1. **"Сборочные единицы"** (Assembly Units) - Priority 1
2. **"Детали"** (Details/Parts) - Priority 2
3. **"Стандартные изделия"** (Standard Products) - Priority 3
4. **"Прочие изделия"** (Other Products) - Priority 4

Sorting within sections:
- Assembly Units & Details: sorted by `Designation`
- Standard Products & Other Products: sorted by `Name`

## Key Dependencies

**NuGet Packages:**
- `ClosedXML.Signed 0.95.4` - Excel file generation
- `DocumentFormat.OpenXml 2.7.2` - Office Open XML support
- `ObjectListView.Official 2.9.1` - Enhanced TreeListView control

**Local Libraries:**
- `netDxf.dll` - DXF file parsing (not from NuGet, included in repository root)

**COM References:**
- Kompas6API5, Kompas6Constants, Kompas6Constants3D, KompasAPI7 - KOMPAS-3D APIs

## Important Notes

### COM Interop
- The application is registered as a COM component during build
- KOMPAS-3D expects specific registry keys under `HKLM\SOFTWARE\Classes\CLSID\{GUID}\Kompas Report`
- `KompasReport.RegisterKompasLib()` and `UnregisterKompasLib()` handle COM registration

### KOMPAS AddIns Registration (from Руководство KOMPAS-Invisible.txt)

For automatic library connection in KOMPAS-3D, AddIn registration is required in the Windows Registry.

**Registry Paths:**
- **KOMPAS v22+**: `HKLM\SOFTWARE\ASCON\KOMPAS-3D\AddIns\{LibraryName}` (general path without version)
- **KOMPAS v18**: `HKLM\SOFTWARE\ASCON\KOMPAS-3D\18.0\AddIns\{LibraryName}`
- Alternative: `HKCU\SOFTWARE\ASCON\...` for per-user registration (no admin rights needed)

**Required Registry Values:**
| Value Name | Type | Description |
|------------|------|-------------|
| `ProgID` | REG_SZ | COM component identifier (e.g., `"ReportKompas.KompasReport"`) |
| `Path` | REG_SZ | Full path to the DLL/library file |
| `AutoConnect` | REG_DWORD | `1` = auto-connect on KOMPAS startup, `0` = manual connection |

**Note:** If both `ProgID` and `Path` are present, `Path` takes priority. ActiveX components must be registered via regasm.

**Implementation:** See `RegisterKompasAddIn()` and `UnregisterKompasAddIn()` in [KompasReport.cs](KompasReport.cs)

**Example .reg file format:**
```reg
[HKEY_LOCAL_MACHINE\SOFTWARE\ASCON\KOMPAS-3D\AddIns\Kompas Report]
"ProgID"="ReportKompas.KompasReport"
"Path"="C:\\Program Files\\...\\ReportKompas.dll"
"AutoConnect"=dword:00000001
```

### Licensing System
- License file location: Application directory + `\bug`
- License format: Hex-encoded → Base64-encoded machine fingerprint
- Fingerprint formula: `Environment.UserName + Environment.OSVersion + Environment.MachineName`
- License validation occurs in `KompasReport.ExternalRunCommand()`

### Material Auto-Detection
When `ObjectAssemblyKompas.ReplaceMaterial()` is called, it automatically assigns materials based on designation patterns:
- `"1.5mm_*_Zn"` → Galvanized sheet (ГОСТ 19904-90)
- `"1.5mm_*_Aisi"` → Stainless steel AISI430
- `"1.5mm_*_Aisi Bronze"` → Bronze-colored stainless steel
- `"1mm_*_Zn"` → 1mm galvanized sheet
- `"1.5mm_*_AL"` → Aluminum checker plate
- `"1.5mm_*_Forbo"` → Forbo flooring material

### Duplicate Handling
The tree structure automatically handles duplicate components: when adding a child with the same `Name` and `Designation`, the `Quantity` is incremented rather than creating a duplicate node. This is critical for accurate assembly reporting.

## Directory Structure (Russian folder names)

```
ReportKompas/
├── Program.cs                               # Application entry point
├── ReportKompas.cs                          # Main UI form (Singleton)
├── KompasReport.cs                          # COM wrapper for KOMPAS-3D
├── Классы структуры данных/                 # Data Structure Classes
│   └── ObjectAssemblyKompas.cs              # Tree-based assembly model
├── Классы настроек/                         # Settings Classes
│   ├── Settings.cs, SettingsForm.cs         # Settings management
│   ├── CodeEquip.cs, CodeMaterial.cs        # Code dictionaries
│   └── SpeedCut.cs                          # Laser cutting speeds
├── Классы обработки DXF/                    # DXF Processing Classes
│   ├── DxfProcessor.cs                      # DXF parsing & calculations
│   └── DimensionsDXF.cs                     # DXF dimensions model
├── Классы покрытий/                         # Coating Classes
│   └── Coating.cs                           # Coating management form
└── Settings/                                # XML configuration files
```

## Development Notes

### When modifying the tree structure
- Always use `AddChild()` to maintain duplicate merging behavior
- Call `SortTreeNodes()` after building the tree to ensure proper ordering
- Update both `ParentK` and `Children` references to maintain bidirectional links

### When adding new materials
- Update `ObjectAssemblyKompas.ReplaceMaterial()` with new designation patterns
- Add material codes to `Settings\CodeMaterial.xml`

### When modifying DXF processing
- Laser cutting speeds are loaded from `SpeedCut.xml` as a dictionary
- Key format for speed lookup: `"{MaterialThickness}_{MaterialType}"`
- All calculations assume metric units (mm, seconds)

### Testing COM integration
- Requires KOMPAS-3D installation for full testing
- Standalone testing: Place a valid license file named `bug` in the output directory
- COM registration requires administrator privileges

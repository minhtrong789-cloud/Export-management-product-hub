# Export Management & Product Hub

A desktop management application built on Microsoft Access (VBA / DAO) tailored for export workflows. The system manages product catalogs, calculates packaging/shipping specifications, handles customer accounts, and automates commercial export document generation (Commercial Invoices, Packing Lists).

Designed with a modular codebase exported and tracked via Version Control System (VCS) conventions.

---

## Key Modules & Technical Scope

### 1. Product Hub & Catalog Architecture
- **Hierarchical Specifications**: Tracks net weight, gross weight, carton dimensions, and CBM (Cubic Meters) per item.
- **Image Integration**: Local asset resolution and media binding supporting dynamic preview without bloating database file size.
- **Staging Buffer Pattern**: Implements an in-memory staging mechanism (`StagingBuffer.cls`) and transaction validation before committing modified catalog rows to disk.

### 2. Standalone Windows GUI & UX
- **Decoupled Shell Window**: Uses Win32 API (`user32.dll`: `SetWindowLongPtr`, `ShowWindow`, `IsWindowVisible`) to detach the primary dashboard from the standard MS Access parent workspace, providing a clean standalone desktop experience.
- **Watchdog Liveness Controller**: Maintains application responsiveness, window state, and clean teardown on exit.
- **Programmatic GUI Builder**: Custom module layout routines (`BuildGUI_*`, `GUIEnforcer_*`) separating structural wireframing from control styling and theme enforcement.

### 3. Commercial Export Document Automation
- **Direct Excel Automation**: Automated export rendering (`ExcelRender_Orders_CommercialInvoice`, `ExcelRender_Products_PrintProducts`) via COM interface (`Excel.Application`).
- **In-Memory Workbook Generation**: Builds formatted commercial invoices, billing calculations, and catalog sheets dynamically in RAM without relying on hardcoded template files on disk.

### 4. Data Layer & Security
- **Dynamic QueryDefs**: Centralized query builders (`BuildSQL`, `BuildSQL_Products`, `BuildSQL_Customers`) managing DAO QueryDef creation, parameterized filtering, and search indices.
- **Data Protection**: Implementation of Windows Cryptography API: Next Generation (CNG / `bcrypt.dll`) via `Extension_Cryptography.bas` for secure token/hash handling.

---

## Architecture & Directory Structure

```text
├── ExportManagement.accde       # Compiled executable binary (Ready to run)
├── ExportManagement.bmp         # Native application splash screen
├── Export_Management.ico        # Multi-resolution application icon
├── MediaStorage/                # Catalog product image assets
└── ExportManagement.accdb.src/  # Deconstructed VCS source code
    ├── forms/                   # Form definitions (.bas) and code-behind (.cls)
    ├── modules/                 # Core procedural modules and class controllers
    │   ├── App_Entry.bas        # Bootstrapper and application lifecycle
    │   ├── Extension_Window.bas # Win32 API window shell controller
    │   ├── Extension_Cryptography.bas # Windows CNG BCrypt integration
    │   ├── StagingBuffer.cls    # In-memory buffer management
    │   ├── DataHub_Products.cls # Catalog transaction service
    │   ├── BuildSQL_*.bas       # DAO QueryDef builders
    │   └── ExcelRender_*.bas    # In-memory Excel automation engines
    ├── queries/                 # SQL query source definitions
    ├── tbldefs/                 # DAO table schema definitions
    └── macros/                  # AutoExec macro definitions
```

---

## Technical Specifications

- **Runtime Target**: Microsoft Access 2016 / 2019 / 2021 / Office 365 (32-bit & 64-bit compatible via `PtrSafe` & `LongPtr`).
- **Database Engine**: Microsoft Access Database Engine (ACE / DAO 12.0+).
- **External Dependencies**:
  - `user32.dll` (Win32 Window Management)
  - `bcrypt.dll` (Windows Cryptography API: Next Generation)
  - `Excel.Application` COM Object (for export reporting)
- **Source Management**: msaccess-vcs-addin compatible schema.

---

## Running the Application

1. **Quick Run (Binary)**:
   - Double-click `ExportManagement.accde` (requires Microsoft Access or free Microsoft Access 365 Runtime installed).
2. **From Source**:
   - Build into an `.accdb` file using the Access VCS Add-in against the `ExportManagement.accdb.src/` directory.

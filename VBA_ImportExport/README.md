# pyVBA – Automated VBA & Access SQL Source Code Sync for Git & SVN

A lightweight, zero-dependency VBA utility that automatically exports and imports all VBA modules, classes, UserForms (`.bas`, `.cls`, `.frm`), and **Microsoft Access queries (`.sql`)** between Office files (**Excel, Word & Access**) and clean plain-text folders for **Git** or **SVN** version control.

📺 **Watch the full tutorial on YouTube:** https://www.youtube.com/watch?v=x08LZku653Y

**Binary Office files (`.xlsm`, `.docm`, `.accdb`) are a nightmare for version control** – no readable diffs, no clean merge history, and constant binary conflicts. **`cVBA_ImportExport`** solves this with a single call: it extracts your entire VBA project (plus Access SQL queries) into readable text files and syncs external changes right back into your open Office host without losing hidden attributes or worksheet bindings.

---

## 🚀 Key Features & Highlights

* **Smart Hybrid Import Engine:** Updates Standard modules (`.bas`) and Document modules (`ThisWorkbook`, worksheets, Access form/report code) *in-place* so existing bindings and events stay intact, while cleanly removing and re-importing Classes (`.cls`) and UserForms (`.frm`) to preserve hidden attributes (`VB_UserMemId = 0`) and binary `.frx` layouts.
* **Microsoft Access Query Sync (`src\Queries\*.sql`):** Automatically exports all saved Access queries (`QueryDefs`) as clean, diff-friendly `.sql` files into a `Queries` subfolder and updates or creates them on import (skipping internal `~` and `MSys` queries). Can be toggled anytime via optional parameters.
* **Access Duplicate-Name Prevention:** Safely renames existing Access class modules (`...OLD`) prior to removal so the original module name is immediately released for VBE `.Import` without triggering Access's automatic `Name1` suffix collision.
* **Zero-Config `src.lnk` Shortcut Support:** Drop a standard Windows folder shortcut named `src.lnk` next to your Office file to export directly into any local Git or SVN repository – no hardcoded paths in your VBA code.
* **Cross-Host Compatibility (Excel, Word & Access):** Automatically detects the active Office host application and resolves local drives as well as OneDrive-synced paths.
* **100% Native Late Binding:** Uses Late Binding for `VBIDE`, `WScript.Shell`, and host-specific `Application` methods – no manual references (`Tools -> References`) required.
* **Self-Protection & Clean Filtering:** Automatically skips itself (`cVBA_ImportExport`) and ignores empty boilerplate modules (2 lines or fewer) during export.

---

## 📦 Core Component Files

1. **`cVBA_ImportExport.cls`**: The main class that handles host detection, OneDrive/shortcut path resolution (`setScrDir`), full component & Access SQL query export (`Export`), and the two-pass hybrid import (`Import`).
2. **`Develop.bas`**: Optional helper module providing ready-to-run `VBA_Export` and `VBA_Import` macros that can be triggered directly from the VBA Editor or a Quick Access Toolbar button.

---

## 🛠️ Critical Installation & Setup Instructions

### 1. Enable Trust Center Access (One-Time Setup)
Because `cVBA_ImportExport` reads and writes components inside the VBA project, Office requires programmatic access to the VBA object model to be enabled once:

1. Open **Excel**, **Word**, or **Access** and go to **File** -> **Options** -> **Trust Center** -> **Trust Center Settings...**
2. Select **Macro Settings** and check **"Trust access to the VBA project object model"**.

### 2. Import the Files into Your Project
1. Download `cVBA_ImportExport.cls` (and optionally `Develop.bas`) from this repository.
2. Open the VBA Editor (`Alt + F11`) and go to **File** -> **Import File...** (`Ctrl + M`).
3. Import **`cVBA_ImportExport.cls`** and **`Develop.bas`**.
*No additional VBA References are required.*

---

## 💻 Real-World Usage

### Trigger Export & Import via Helper Module (`Develop.bas`)
```vba
Sub VBA_Export()
    Dim VBA_ImportExport As New cVBA_ImportExport
    
    ' Exports all .bas, .cls, .frm files into "src" (and Access queries into "src\Queries\*.sql")
    VBA_ImportExport.Export
    
    ' Optional (Access only): Export VBA modules without exporting SQL queries
    ' VBA_ImportExport.Export ExportQueries_:=False
End Sub

Sub VBA_Import()
    Dim VBA_ImportExport As New cVBA_ImportExport
    
    ' Syncs all external .bas, .cls, .frm changes (and Access "src\Queries\*.sql") back into the host
    VBA_ImportExport.Import
    
    ' Optional (Access only): Import VBA modules without overwriting SQL queries
    ' VBA_ImportExport.Import ImportQueries_:=False
End Sub
```

### Target Directory Resolution Order
When calling `.Export` or `.Import`, the target folder is resolved automatically in the following order:
1. **`.\src\`** – Uses an existing `src` subfolder right next to your Office document (Access queries are stored in `.\src\Queries\`).
2. **`.\src.lnk`** – Resolves a Windows folder shortcut named `src.lnk` pointing directly to your Git or SVN working copy.
3. **Auto-Create `.\src\`** – Automatically creates a new local `src` folder if neither exists.

---

## 🖖 Star Trek Rationale
> *"Software is structure, not syntax – so that Picard knows what Kirk did, even though Kirk has been dead for a long time."*

---

## 📄 License

This framework is licensed under the **MIT License**. Feel free to use, fork, alter, and embed these classes into private or commercial macro projects.

<p align="center">
  <img src="docs/banniere.png" alt="CustomUI Studio — Your Excel ribbons, designed with the mouse" width="100%">
</p>

<h1 align="center">CustomUI Studio</h1>

<p align="center">
  <a href="README.md">Français</a> · <b>English</b>
</p>

<p align="center">
  <b>Visual ribbon editor for Microsoft Excel — free, portable, no installation required.</b><br>
  Design your tabs, groups, buttons, and right-click menus with your mouse, without writing a single line of XML.
</p>

<p align="center">
  <a href="https://github.com/CustomUI-Studio/CustomUI-Studio/releases/latest"><img src="https://img.shields.io/github/v/release/CustomUI-Studio/CustomUI-Studio?label=version&color=1f7a4a" alt="Latest release"></a>
  <a href="https://github.com/CustomUI-Studio/CustomUI-Studio/releases"><img src="https://img.shields.io/github/downloads/CustomUI-Studio/CustomUI-Studio/total?label=downloads&color=1f7a4a" alt="Downloads"></a>
  <img src="https://img.shields.io/badge/Windows-64%20bits-1f7a4a" alt="Windows 64-bit">
  <img src="https://img.shields.io/badge/license-freeware-1f7a4a" alt="License: freeware">
</p>

<p align="center">
  <a href="https://github.com/CustomUI-Studio/CustomUI-Studio/releases/latest/download/CustomUI.Studio.1.18.0.exe"><b>⬇️ Download CustomUI Studio 1.18.0</b></a>
  · <a href="https://apps.microsoft.com/detail/9PH54K3T7TCT"><b>Microsoft Store</b></a>
</p>

![CustomUI Studio Visual Builder](docs/capture-constructeur.png)[cite: 1]

## Why CustomUI Studio?[cite: 1]

Writing an Excel ribbon XML by hand is tedious: a single typo, empty attribute, or duplicate ID, and the ribbon fails to display without any error message.[cite: 1]

CustomUI Studio solves the problem at the root. You build the ribbon visually, the application only offers what Excel accepts, and it validates everything before writing to your workbook.[cite: 1]

## In Three Steps[cite: 1]

1. **Import** your `.xlsx`, `.xlsm` workbook, or `.xlam` add-in, or start from a blank ribbon or a template.[cite: 1]
2. **Design** the ribbon: add tabs, groups, and controls, choose their icons, and adjust their properties. The preview updates live.[cite: 1]
3. **Save**: the ribbon is written into the workbook, and the accompanying VBA module is ready to import.[cite: 1]

## Features[cite: 1]

### Building the Ribbon[cite: 1]

- **All standard controls**: buttons, toggle buttons, checkboxes, edit boxes, drop-down lists, combo boxes, galleries, menus, dynamic menus, split buttons.[cite: 1]
- **Only allowed elements are offered**: it is impossible to build a ribbon that Excel would reject. The order required by the format is automatically respected.[cite: 1]
- **Drag and drop**: reorganize tabs, groups, and controls with the mouse within the Ribbon Structure.[cite: 1]
- **Duplicate, copy, paste** (Ctrl+D, Ctrl+C, Ctrl+V) a tab, group, or control, including across projects: copies receive their own IDs and VBA procedures.[cite: 1]
- **Live preview, just like in Excel**: test lists, menus, and tooltips on hover.[cite: 1]
- **Icons**: categorized imageMso library with search, custom images, and native Excel commands catalog (Save, Copy, AutoSum...). With Excel installed, the preview displays actual Excel icons.[cite: 1]
- **Two ribbon formats**: CustomUI 2006 (Office 2007) and CustomUI 2009 (Office 2010 and later).[cite: 1]
- **Contextual tabs and Backstage**: complete the tabs Excel displays for a table, chart, or image, and build the File menu (Backstage) directly in the editor, with a preview close to Excel 365: Excel commands, hiding, renaming, and placement of each added item. A File menu can also be stored in the library as a component; the "Backstage – Administration" component is provided.[cite: 1]
- **Quick Access Toolbar**: add Excel commands, your own buttons, and separators, displayed when the workbook is active (or across all workbooks for an `.xlam` add-in). The application reminds you of what Excel requires (`startFromScratch`) and offers to keep Undo and Redo, which the workbook's toolbar replaces.[cite: 1]
- **Excel tabs and groups**: add your groups to the Home tab or another Excel tab, hide an Excel tab or group, copy any Excel group into your own tab (all groups from every tab, rendered in the preview just like in Excel), and place your tabs and groups before or after Excel's. Keyboard shortcuts and display settings (grayed out, image only, gallery columns...) are available in the inspector.[cite: 1]
- **Workbook display**: hide the formula bar, status bar, ribbon, sheet tabs, headers, or gridlines, and set a window title; the generated VBA code applies these settings and restores Excel defaults that affect all workbooks when yours loses focus.[cite: 1]
- **Undo / Redo** (Ctrl+Z, Ctrl+Y) for every modification.[cite: 1]
- **Displayed compatibility**: for the targeted Excel version, every command, icon, and menu carries a "verified", "documented", or "unknown" badge, and a report summarizes the project.[cite: 1]
- **In French or English**: language is chosen at first launch and can be changed at any time.[cite: 1]

![Live ribbon preview](docs/apercu-ruban.png)[cite: 1]

### Projects, Components, and Templates[cite: 1]

- **`.cuis` projects**: save a ribbon independently of any workbook, reopen it, and apply it to the workbook of your choice.[cite: 1]
- **`.cuic` reusable components**: a group, tab, or menu with its VBA code, ready to be inserted into any project. A local library organizes them, and several components are provided (PDF export, sheet navigation...).[cite: 1]
- **Ready-to-use templates**: inventory management, business management, dashboard, complete with their ribbon, VBA code, and the "Prepare Workbook" macro, which creates expected sheets and tables without touching existing data.[cite: 1]

### Right-Click Menus[cite: 1]

- **Your commands in Excel's right-click menu**: all 89 menus from Microsoft's official list (cell, row, column, table, sheet tab, charts, shapes, slicers...). Place them before or after an Excel command.[cite: 1]
- **Entries displayed only when meaningful**: on a specific sheet, in a specific table or column, on a non-empty cell... The conditional VBA code is generated for you.[cite: 1]
- **Copy an entry to other menus** and organize it into submenus.[cite: 1]
- **Hide Excel commands in one click**: for cells, rows, columns, and tables, check the commands to keep or apply a preset ("Minimal Menu", "No Copy-Paste"...).[cite: 1]
- **Faithful preview**: the menu displays exactly as Excel will show it.[cite: 1]
- Commands that the ribbon cannot hide are handled via VBA, with procedures generated for you.[cite: 1]

Right-click menus require the CustomUI 2009 format.[cite: 1]

![Right-click menu preview](docs/capture-menu-clic-droit.png)[cite: 1]

### Generated VBA Code[cite: 1]

- **Ready-to-use `RibbonDESIGNER` module**: every callback has the correct signature, with commented examples.[cite: 1]
- **One-macro import**: `ImporterModuleRibbonDESIGNER` places the module into your VBA project without duplicates.[cite: 1]
- **"Apply to a new workbook..." and "Finish in Excel" guide**: the application creates the workbook, then shows remaining steps (importing modules, pasting a snippet, reopening). In the desktop version, modules are saved alongside the workbook and a button opens it in Excel.[cite: 1]
- **Ribbon driven by your macros**: check the properties of a control that need to change dynamically (label, disabled, hidden, image...), then a single line is enough:[cite: 1]

  ```vb
  SetControlState "btnExporter", "enabled", False
  ```[cite: 1]

### Add-ins and Batch Processing[cite: 1]

- **Excel `.xlam` add-ins**: open, edit, and re-save an add-in; the generated VBA code adapts (acting on the active workbook). A guided procedure explains how to create the add-in in Excel.[cite: 1]
- **Apply to multiple workbooks** (desktop version): choose workbooks or a folder, inspect them before writing, and get copies in an output folder along with a CSV report. Originals are never modified.[cite: 1]

### Controlled AI Assistant[cite: 1]

- **Describe what you want**: the application prepares a prompt to copy into your AI tool of choice, then checks its response before any modification.[cite: 1]
- **Three targets**: the ribbon and its right-click menus, a contextual tab (choose the set: table, chart, image...), or the File menu (Backstage), with the assistant displaying the list as Excel will show it. Each request targets a single component.[cite: 1]
- **Nothing invented**: any attribute, control, or command unknown to the format is rejected; each change is presented, deletions are confirmed, and "Undo" reverts everything in a single step.[cite: 1]
- **No connection required**: CustomUI Studio transmits nothing. VBA code, images, and filenames never appear in the prompt.[cite: 1]

### Pre-Save Validation[cite: 1]

- **Offline verification using Microsoft schemas**: the XML editor lists every issue with its line number, a clear explanation, and the required fix.[cite: 1]
- **Saving blocked while errors remain**: unknown or empty attributes, duplicate IDs, misplaced controls, incorrect order...[cite: 1]
- **"Issues" panel in the Visual Builder**: errors appear during design; a click selects the affected element.[cite: 1]
- **XML editor synchronized** with the interface.[cite: 1]

![XML editor and Issues panel](docs/capture-editeur-xml.png)[cite: 1]

## Your Workbook is Respected[cite: 1]

When you click **Save Office**, your workbook is not rebuilt from scratch:[cite: 1]

- all original parts are copied identically and verified before writing;[cite: 1]
- only the ribbon changes: your sheets, data, and VBA macros remain intact;[cite: 1]
- any image deleted from the library is removed from the workbook along with its link;[cite: 1]
- writing goes through a temporary file, which is read back and substituted. In case of issues, the original file is untouched.[cite: 1]

## Installation[cite: 1]

### Portable Version[cite: 1]

1. Download [`CustomUI.Studio.1.18.0.exe`](https://github.com/CustomUI-Studio/CustomUI-Studio/releases/latest/download/CustomUI.Studio.1.18.0.exe) from the [Releases](https://github.com/CustomUI-Studio/CustomUI-Studio/releases/latest) page.[cite: 1]
2. Double-click it. No installation required.[cite: 1]

> **"Windows protected your PC" message?**[cite: 1]
> The portable executable is not digitally signed, triggering Windows SmartScreen upon first launch. Click **More info**, then **Run anyway**.[cite: 1]

<details>
<summary>Verify downloaded file (optional)</summary>[cite: 1]

<br>[cite: 1]

SHA-256 hash for `CustomUI.Studio.1.18.0.exe`:[cite: 1]

To verify in PowerShell:[cite: 1]

```powershell
Get-FileHash .\CustomUI.Studio.1.18.0.exe -Algorithm SHA256
```[cite: 1]

</details>[cite: 1]

### Microsoft Store[cite: 1]

CustomUI Studio is available on the [Microsoft Store](https://apps.microsoft.com/detail/9PH54K3T7TCT). The Store version is signed by Microsoft, installs without Windows warnings, and updates automatically.[cite: 1]

## System Requirements[cite: 1]

- Windows 64-bit[cite: 1]
- Desktop Microsoft Excel, to use the generated workbooks[cite: 1]
- Supported files: `.xlsx`, `.xlsm`, and `.xlam` add-ins[cite: 1]
- For the VBA module import macro: Excel option **Trust access to the VBA project object model** must be enabled[cite: 1]

## Good to Know[cite: 1]

- Modifying the ribbon of a digitally signed workbook invalidates its signature: it must be re-signed.[cite: 1]
- A workbook downloaded from the Internet may be blocked by Windows (red banner). Right-click the file, select **Properties**, and check **Unblock**.[cite: 1]
- Validation applies to the ribbon XML. Your macros and command IDs are not checked: always test the ribbon inside Excel.[cite: 1]
- Drag and drop is done within the Ribbon Structure, not from the palette or the preview.[cite: 1]
- Right-click menus: certain Excel commands can only be hidden via VBA ("VBA" label). This hiding affects all of Excel while the workbook is active; procedures to paste into `ThisWorkbook` are provided at the top of the `RibbonDESIGNER` module.[cite: 1]
- The list of Excel right-click commands was captured from Microsoft 365 in French. Excel may display others depending on context and version.[cite: 1]
- In English, Excel right-click command names and the generated VBA module remain in French.[cite: 1]
- In an `.xlam` add-in, `ThisWorkbook` refers to the add-in itself: generated code uses the active workbook, and the application flags component code that uses `ThisWorkbook`.[cite: 1]
- The AI assistant does not manage library components or custom images; in the File menu, it only adjusts the 13 Excel commands captured from Microsoft 365.[cite: 1]

## Versions[cite: 1]

The current version is 1.18.0. Details for each version are in the [changelog](CHANGELOG.md).[cite: 1]

## Questions, Suggestions, Bugs[cite: 1]

Open an issue in the [Issues](https://github.com/CustomUI-Studio/CustomUI-Studio/issues) tab of this repository.[cite: 1]

## Privacy[cite: 1]

CustomUI Studio does not collect or transmit any data and operates without an internet connection, including the AI assistant: you copy the prompt into your own AI tool yourself. See the [privacy policy](CONFIDENTIALITE.md).[cite: 1]

## Third-Party Components[cite: 1]

XML validation uses libxml2 (MIT license) via xmllint-wasm (MIT license). CustomUI schemas belong to Microsoft Corporation.[cite: 1]

## License[cite: 1]

Freeware, source code not published. See [LICENCE.txt](LICENCE.txt).[cite: 1]

© 2026 Régis DEVAUX — All rights reserved.[cite: 1]

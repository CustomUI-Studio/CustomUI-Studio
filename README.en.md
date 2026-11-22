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

![CustomUI Studio Visual Builder](docs/capture-constructeur.png)

## Why CustomUI Studio?

Writing an Excel ribbon XML by hand is tedious: a single typo, empty attribute, or duplicate ID, and the ribbon fails to display without any error message.

CustomUI Studio solves the problem at the root. You build the ribbon visually, the application only offers what Excel accepts, and it validates everything before writing to your workbook.

## In Three Steps

1. **Import** your `.xlsx`, `.xlsm` workbook, or `.xlam` add-in, or start from a blank ribbon or a template.
2. **Design** the ribbon: add tabs, groups, and controls, choose their icons, and adjust their properties. The preview updates live.
3. **Save**: the ribbon is written into the workbook, and the accompanying VBA module is ready to import.

## Features

### Building the Ribbon

- **All standard controls**: buttons, toggle buttons, checkboxes, edit boxes, drop-down lists, combo boxes, galleries, menus, dynamic menus, split buttons.
- **Only allowed elements are offered**: it is impossible to build a ribbon that Excel would reject. The order required by the format is automatically respected.
- **Drag and drop**: reorganize tabs, groups, and controls with the mouse within the Ribbon Structure.
- **Duplicate, copy, paste** (Ctrl+D, Ctrl+C, Ctrl+V) a tab, group, or control, including across projects: copies receive their own IDs and VBA procedures.
- **Live preview, just like in Excel**: test lists, menus, and tooltips on hover.
- **Icons**: categorized imageMso library with search, custom images, and native Excel commands catalog (Save, Copy, AutoSum...). With Excel installed, the preview displays actual Excel icons.
- **Two ribbon formats**: CustomUI 2006 (Office 2007) and CustomUI 2009 (Office 2010 and later).
- **Contextual tabs and Backstage**: complete the tabs Excel displays for a table, chart, or image, and build the File menu (Backstage) directly in the editor, with a preview close to Excel 365: Excel commands, hiding, renaming, and placement of each added item. A File menu can also be stored in the library as a component; the "Backstage – Administration" component is provided.
- **Quick Access Toolbar**: add Excel commands, your own buttons, and separators, displayed when the workbook is active (or across all workbooks for an `.xlam` add-in). The application reminds you of what Excel requires (`startFromScratch`) and offers to keep Undo and Redo, which the workbook's toolbar replaces.
- **Excel tabs and groups**: add your groups to the Home tab or another Excel tab, hide an Excel tab or group, copy any Excel group into your own tab (all groups from every tab, rendered in the preview just like in Excel), and place your tabs and groups before or after Excel's. Keyboard shortcuts and display settings (grayed out, image only, gallery columns...) are available in the inspector.
- **Workbook display**: hide the formula bar, status bar, ribbon, sheet tabs, headers, or gridlines, and set a window title; the generated VBA code applies these settings and restores Excel defaults that affect all workbooks when yours loses focus.
- **Undo / Redo** (Ctrl+Z, Ctrl+Y) for every modification.
- **Displayed compatibility**: for the targeted Excel version, every command, icon, and menu carries a "verified", "documented", or "unknown" badge, and a report summarizes the project.
- **In French or English**: language is chosen at first launch and can be changed at any time.

![Live ribbon preview](docs/apercu-ruban.png)

### Projects, Components, and Templates

- **`.cuis` projects**: save a ribbon independently of any workbook, reopen it, and apply it to the workbook of your choice.
- **`.cuic` reusable components**: a group, tab, or menu with its VBA code, ready to be inserted into any project. A local library organizes them, and several components are provided (PDF export, sheet navigation...).
- **Ready-to-use templates**: inventory management, business management, dashboard, complete with their ribbon, VBA code, and the "Prepare Workbook" macro, which creates expected sheets and tables without touching existing data.

### Right-Click Menus

- **Your commands in Excel's right-click menu**: all 89 menus from Microsoft's official list (cell, row, column, table, sheet tab, charts, shapes, slicers...). Place them before or after an Excel command.
- **Entries displayed only when meaningful**: on a specific sheet, in a specific table or column, on a non-empty cell... The conditional VBA code is generated for you.
- **Copy an entry to other menus** and organize it into submenus.
- **Hide Excel commands in one click**: for cells, rows, columns, and tables, check the commands to keep or apply a preset ("Minimal Menu", "No Copy-Paste"...).
- **Faithful preview**: the menu displays exactly as Excel will show it.
- Commands that the ribbon cannot hide are handled via VBA, with procedures generated for you.

Right-click menus require the CustomUI 2009 format.

![Right-click menu preview](docs/capture-menu-clic-droit.png)

### Generated VBA Code

- **Ready-to-use `RibbonDESIGNER` module**: every callback has the correct signature, with commented examples.
- **One-macro import**: `ImporterModuleRibbonDESIGNER` places the module into your VBA project without duplicates.
- **"Apply to a new workbook..." and "Finish in Excel" guide**: the application creates the workbook, then shows remaining steps (importing modules, pasting a snippet, reopening). In the desktop version, modules are saved alongside the workbook and a button opens it in Excel.
- **Ribbon driven by your macros**: check the properties of a control that need to change dynamically (label, disabled, hidden, image...), then a single line is enough:

  ```vb
  SetControlState "btnExporter", "enabled", False

  Add-ins and Batch Processing

    Excel .xlam add-ins: open, edit, and re-save an add-in; the generated VBA code adapts (acting on the active workbook). A guided procedure explains how to create the add-in in Excel.

    Apply to multiple workbooks (desktop version): choose workbooks or a folder, inspect them before writing, and get copies in an output folder along with a CSV report. Originals are never modified.

Controlled AI Assistant

    Describe what you want: the application prepares a prompt to copy into your AI tool of choice, then checks its response before any modification.

    Three targets: the ribbon and its right-click menus, a contextual tab (choose the set: table, chart, image...), or the File menu (Backstage), with the assistant displaying the list as Excel will show it. Each request targets a single component.

    Nothing invented: any attribute, control, or command unknown to the format is rejected; each change is presented, deletions are confirmed, and "Undo" reverts everything in a single step.

    No connection required: CustomUI Studio transmits nothing. VBA code, images, and filenames never appear in the prompt.

Pre-Save Validation

    Offline verification using Microsoft schemas: the XML editor lists every issue with its line number, a clear explanation, and the required fix.

    Saving blocked while errors remain: unknown or empty attributes, duplicate IDs, misplaced controls, incorrect order...

    "Issues" panel in the Visual Builder: errors appear during design; a click selects the affected element.

    XML editor synchronized with the interface.

Your Workbook is Respected

When you click Save Office, your workbook is not rebuilt from scratch:

    all original parts are copied identically and verified before writing;

    only the ribbon changes: your sheets, data, and VBA macros remain intact;

    any image deleted from the library is removed from the workbook along with its link;

    writing goes through a temporary file, which is read back and substituted. In case of issues, the original file is untouched.

Installation
Portable Version

    Download CustomUI.Studio.1.18.0.exe from the Releases page.

    Double-click it. No installation required.

    "Windows protected your PC" message?
    The portable executable is not digitally signed, triggering Windows SmartScreen upon first launch. Click More info, then Run anyway.

SHA-256 hash for CustomUI.Studio.1.18.0.exe:
cecf15c36ef918ce0c3845e04f0ff0372f71bf7be7439e7fce813753e60cf1bc

To verify in PowerShell:
Get-FileHash .\CustomUI.Studio.1.18.0.exe -Algorithm SHA256

Microsoft StoreCustomUI Studio is available on the Microsoft Store. The Store version is signed by Microsoft, installs without Windows warnings, and updates automatically.System RequirementsWindows 64-bitDesktop Microsoft Excel, to use the generated workbooksSupported files: .xlsx, .xlsm, and .xlam add-insFor the VBA module import macro: Excel option Trust access to the VBA project object model must be enabledGood to KnowModifying the ribbon of a digitally signed workbook invalidates its signature: it must be re-signed.A workbook downloaded from the Internet may be blocked by Windows (red banner). Right-click the file, select Properties, and check Unblock.Validation applies to the ribbon XML. Your macros and command IDs are not checked: always test the ribbon inside Excel.Drag and drop is done within the Ribbon Structure, not from the palette or the preview.Right-click menus: certain Excel commands can only be hidden via VBA ("VBA" label). This hiding affects all of Excel while the workbook is active; procedures to paste into ThisWorkbook are provided at the top of the RibbonDESIGNER module.The list of Excel right-click commands was captured from Microsoft 365 in French. Excel may display others depending on context and version.In English, Excel right-click command names and the generated VBA module remain in French.In an .xlam add-in, ThisWorkbook refers to the add-in itself: generated code uses the active workbook, and the application flags component code that uses ThisWorkbook.The AI assistant does not manage library components or custom images; in the File menu, it only adjusts the 13 Excel commands captured from Microsoft 365.VersionsThe current version is 1.18.0. Details for each version are in the changelog.Questions, Suggestions, BugsOpen an issue in the Issues tab of this repository.PrivacyCustomUI Studio does not collect or transmit any data and operates without an internet connection, including the AI assistant: you copy the prompt into your own AI tool yourself. See the privacy policy.Third-Party ComponentsXML validation uses libxml2 (MIT license) via xmllint-wasm (MIT license). CustomUI schemas belong to Microsoft Corporation.LicenseFreeware, source code not published. See LICENCE.txt.© 2026 Régis DEVAUX — All rights reserved.   


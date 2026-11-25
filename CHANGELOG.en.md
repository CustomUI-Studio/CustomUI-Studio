# Version History

The latest version can be downloaded from the [Releases](https://github.com/CustomUI-Studio/CustomUI-Studio/releases/latest) page.

## 1.20.1

Live preview fix: it shows the ribbon as Excel does. Projects and workbooks from previous versions open and save as before.

### Fixes

- **Large buttons**: the label is written on two balanced lines, as in Excel ("Nouvel / Article", "Envoyer par / courriel ▾"), and the button takes the width of the longest line, without "…". Same rule for the Excel groups drawn in the preview.
- **Label** (`labelControl`): text only, no image.
- **Control without an image** (button, toggle button, menu, split button, gallery): no image, as in Excel; in large size, the icon space stays empty.
- **Menu or gallery with an imageMso**: its image is shown.
- **Check box**: unchecked at first, as in Excel.
- **Missing icon**: a generic picture replaces the icon in the preview, the structure, the inspector and the catalogs.
- **Ribbon structure**: File menu (Backstage) items and each right-click menu can be collapsed.
- **Properties panel**: its header stays on one line, with the element type in capitals (translated); collapsed, the whole strip reopens it, even at the bottom of a long page; open, a click on its title bar collapses it.

## 1.20.0

Keyboard shortcuts and an AI installed on your computer. Projects and workbooks from previous versions open and save as before.

### New

- **Keyboard shortcuts** (VBA Callbacks panel, "Keyboard shortcuts" section): disable an Excel key (Ctrl+C, Ctrl+V, F2…) or assign it to a procedure of the `RibbonDESIGNER` module, through `Application.OnKey`. Keys are set when the workbook opens or comes back to the front, and given back to Excel when it loses focus or closes (`ThisWorkbook` snippet to paste). The right-click "No copy and paste" preset can also disable the six copy-and-paste shortcuts (Ctrl+C, Ctrl+X, Ctrl+V, Ctrl+Insert, Shift+Insert, Shift+Delete). Verified in Excel for Microsoft 365.
- **AI Assistant: "AI on this computer" mode** (desktop version, optional): the prompt shown is sent in one click to Ollama or LM Studio installed on your computer, and the answer is checked automatically before any change. Only addresses of your own computer are accepted: nothing is sent over the Internet. Copy and paste remains the default mode.

### Good to know

- Shortcuts: no effect while typing in a cell or in dialog boxes; this is not a protection. In an add-in, closing the add-in without quitting Excel leaves the keys taken until Excel restarts: call `RestaurerRaccourcis` first (example in the module).
- AI on this computer: install Ollama or LM Studio and a model yourself (both free); a model with at least 7 billion parameters is recommended. On the processor, a request can take several minutes.
- There is no version 1.19.0: 1.20.0 brings together phases 19 and 20.

## 1.18.0

Build a ribbon end to end from the application. Projects and workbooks from previous versions open and save as before.

### New

- **"Problems" panel in the Visual builder**, below the preview: ribbon errors and warnings while you build, with the same rules as when saving. A click selects the element; a problem without an element opens the XML editor at the right line. Markers in the tree and a "Problems with this element" box in the inspector.
- **Duplicate, copy and paste** a tab, a group or a control (inspector bar, Ctrl+D, Ctrl+C, Ctrl+V), also from one project to another. Free identifiers (`<id>_2`); VBA procedures derived from the identifier are renamed, shared procedures stay shared; the keyboard shortcut (`keytip`) is not copied.
- **"Apply to a new workbook…"**: a blank workbook (`.xlsm` if the ribbon calls macros, otherwise `.xlsx`) that receives the project's ribbon.
- **"Finish in Excel" guide** after every "Apply", reopened from the VBA Callbacks panel: only the steps the project needs. In the desktop application, VBA modules saved next to the workbook (never replaced without confirmation) and an "Open in Excel" button.

### Good to know

- CustomUI Studio still never changes a workbook's VBA code: modules are imported in Excel.
- In a new `.xlsm`, until the modules are imported, Excel reports on opening that it cannot run the `CustomUI_OnLoad` macro: click OK, then import the modules.

## 1.15.0

A field for every useful attribute of Microsoft's schema. Projects and workbooks from previous versions open and save as before.

### New

- **Tabs** (custom and contextual): visibility (`visible`), `getVisible`, `getLabel` and `getKeytip` callbacks, tag (`tag`). The label and the visibility can also be driven by a macro with `SetControlState`, as for a control.
- **Groups**: icon of the collapsed group (`imageMso` or a workbook image), automatic scaling of large buttons when the window narrows (`autoScale`), vertical centering (`centerVertically`), visibility and tag. Label and visibility can be driven by `SetControlState`.
- **Dialog box launcher button**: enabled state, callbacks (`getEnabled`, `getVisible`…), keyboard shortcut, display and tag.
- **Gallery and combo box**: refresh on each opening (`invalidateContentOnDrop`); gallery: `showInRibbon`; menu separator: title.
- **Repurposing of Office commands** (`<commands>`) on screen: new "Office commands (repurposing)" section; an Excel command can be grayed out (`enabled`, `getEnabled`) or replaced by a macro when it is started from the ribbon (`onAction`).
- **"Other schema attributes"**: any attribute allowed at that place without a dedicated field remains editable.

### Fixes

- Fields the schema does not allow at that place are no longer offered (input was lost until now): split button label (it is set on its button), idMso, visibility and image of the dialog box launcher, separators and containers.
- A custom tab is no longer shown struck through in the tree when `startFromScratch` is `true`.
- An Office control without `idMso` blocks saving: Excel ignored it and showed "Missing ID".

## 1.14.1

Backstage preview and open workbook format fix. Projects and workbooks do not change.

### Fixes

- **Backstage preview.** Columns no longer overlap: as in Excel, a column widens to fit its content, the second one starts after the first and the page scrolls horizontally when needed.
- **Open workbook format.** When a `.xlsm` or `.xlam` workbook is opened, its format becomes the project's: the "Format" list shows it (greyed out, "— open workbook") and, for a `.xlam` add-in, the RibbonDESIGNER module is generated in add-in mode. A `.xlsx` workbook does not change the project's format.

### Improvement

- **Collapsible Properties panel** to give the preview more room; its state is remembered. Smaller margins around the preview.

## 1.14.0

Workbook display settings generated in VBA. Projects and workbooks from previous versions open and save as before.

### New

- **Workbook display.** New section of the "VBA callbacks" panel: formula bar, status bar, whole ribbon, sheet tabs, scroll bars, headings, gridlines ("leave unchanged", "hide" or "show") and window title. Each setting shows its scope: all of Excel, this workbook or all sheets. Nothing is written to the ribbon XML.
- **Generated code.** The RibbonDESIGNER module applies the settings on opening (`AppliquerAffichage`) and gives back to Excel what affects all of Excel when another workbook becomes active or the workbook closes (`RestaurerAffichage`). A button copies the snippet to paste into ThisWorkbook; with the CommandBars option, a single snippet covers both.
- **Add-in (.xlam).** Only settings that affect all of Excel are offered (an add-in has no window).
- Behaviours surveyed in Excel for Microsoft 365, then checked on a workbook built by the app.

## 1.13.0

Every Excel group, and a preview that draws them as Excel does. Projects and workbooks from previous versions open and save as before.

### New

- **Every Excel group.** The "Copy an Excel group into this tab" list now offers the groups of every tab (Insert, Formulas, Data, View…), sorted by tab in Excel's order, followed by the other groups documented by Microsoft for the targeted Excel version. "Before / after" positions offer every group of the tab.
- **Faithful preview.** Excel groups are drawn with their real buttons, lists and checkboxes, in place and at size; in the desktop app, with Excel's icons. The "Alt" button also shows Excel's keytips.
- Ribbon surveyed in Excel for Microsoft 365; each group offered as verified was copied into a tab and shown by Excel.

## 1.12.0

Excel tabs and groups in the visual editor, keytips and display settings. Projects and workbooks from previous versions open and save as before.

### New

- **Excel tabs and groups.** Add an Excel tab (Home, Insert…) to your project to put your groups in it, hide it, or hide some of its groups with a checkbox. Copy an Excel group, for example Font, into your own tab. With `startFromScratch`, the added Excel tab is shown again (Excel requires it).
- **Position.** Place your tab before or after an Excel tab (for example right after Home), and your groups before or after an Excel group of the same tab.
- **Keytips.** `keytip` field on tabs, groups and controls, with a warning when a letter is already used by Excel (Excel would replace it silently).
- **Display settings.** New inspector section: greyed out, image only or label only, width and length of edit boxes, box layout, large menu items with description, gallery columns… Only the settings the schema allows for the chosen element are offered.
- **Preview.** Excel tabs in their place, their groups by name, your settings applied, and an "Alt" button that shows the keytips.
- Behaviours observed in Excel for Microsoft 365 and verified on a real workbook.

## 1.11.1

Quick Access Toolbar preview fix. Projects and workbooks do not change.

### Fixes

- **Quick Access Toolbar preview.** Buttons with a custom image now show it (only Office's built-in icons appeared). As in Excel, a button without a label shows only its image; its name remains visible as a tooltip.

### Improvement

- **Ribbon structure.** The "Quick Access Toolbar", "Context menus", "Contextual tabs" and "Backstage" sections collapse with a small "−" and expand with "+". "Collapse all" and "Expand all" include them.

## 1.11.0

Quick Access Toolbar in the visual editor. Projects and workbooks from previous versions open and save as before.

### New

- **Quick Access Toolbar.** New section of the Ribbon structure: add Excel commands (catalog with compatibility badge), your own buttons linked to a VBA procedure and separators, then reorder them. Two lists: "This workbook", shown while the workbook is active, and "Shared", shown in all workbooks for an `.xlam` add-in.
- **What Excel requires, explained.** Excel only applies the toolbar when `startFromScratch` is "true" (otherwise it ignores it and shows an error message), and the workbook's toolbar replaces the user's. On the first addition, the application offers to keep Undo and Redo and to set `startFromScratch` to "true", reminding you that Excel's tabs are then hidden. Nothing changes without your choice.
- **Preview.** The toolbar is shown below the ribbon, with its icons and labels, as in Excel.
- Behaviors recorded in Excel for Microsoft 365 and verified on a real workbook.

## 1.10.0

AI assistant by target: ribbon, contextual tab or File menu (Backstage). Projects and workbooks from previous versions open and save as before.

### New

- **Target choice.** Before describing what you want, choose "Ribbon", "Contextual tab" or "Backstage". The request prepared for your AI tool only contains what concerns that target (allowed elements, example, current structure), and an answer aimed at another target is refused.
- **Contextual tabs.** You choose the tab set (table, chart, PivotTable, picture…) from a list; the assistant adds tabs, groups and controls to it, without ever touching the rest of the ribbon.
- **File menu (Backstage).** The assistant adds pages, groups, buttons and quick commands, places them before or after an Excel command, hides, renames or disables an Excel command and sets `onShow` / `onHide`. Only the 13 File menu commands recorded in Excel are accepted. The proposal shows the File menu list as Excel will display it.
- The safeguards are unchanged: no connection, nothing is applied if an operation is refused, no VBA code from the AI, removals confirmed, a single "Undo" step.

### Compatibility

- Answers in the 1.8 format (without the `cible` field) are still accepted for the ribbon.
- No change to `.cuis` and `.cuic` files.

## 1.9.0

Backstage (File menu) in the library, and a File menu preview as close as possible to Excel. Projects and workbooks from previous versions open and save as before.

### What's new

- **File menu preview close to Excel 365.** The left-hand list follows Excel's (Home, New, Open, Info… Account, Options): hidden commands disappear, renamed commands show their new name, and each added item sits where Excel puts it (before or after an Excel command, otherwise at the bottom of the list). The page shows groups, the large primary button, check boxes, lists and option buttons with Excel's colors and layout. Order, colors and proportions were measured in Excel.
- **Backstage in the library.** A project's File menu can be saved as a `.cuic` component, with its VBA code, and inserted into another project. Insertion prefixes identifiers and callbacks, never touches Excel commands, and refuses anything that would conflict (two settings of the same Excel command, a second `onShow` callback).
- **Provided "Backstage – Administration" component.** An "Administration" page placed after "Info" (export options, approver name, output format, target server, workbook status) and a "Synchronize" quick command, checked in Excel with their VBA code.
- **File menu commands verified.** The identifiers of Excel's File menu commands (Home `PlaceTabHome`, Info `TabInfo`, Save As `TabSave`, Account `TabHelp`, Options `ApplicationOptionsDialog`…) now carry the "verified" badge.

### Fixes

- A project without any ribbon tab (Backstage only) produced XML that Excel rejected; it is now valid.
- The File menu preview shrinks to fit entirely in a narrow window.

### Simplification

- The target platform choice ("Windows / Other") is removed from the preview bar: compatibility is assessed for Excel on Windows, the only verified platform. Projects that specified another platform open as before.

### Compatibility

- A Backstage `.cuic` file, or a `.cuis` project containing one, does not open in versions earlier than 1.9.0: they report that it was created by a newer version. Other files remain readable by 1.8.

## 1.8.1

Display fix. Projects and workbooks are unchanged.

### Fixes

- When the window is narrow, the buttons of the header and toolbars (VBA module with several components, Properties panel, image libraries, idMso catalog) wrap to a new line instead of going out of frame or overlapping. The "Export XML / Save Office" button then shows only its icon, with its name as a tooltip.

## 1.8.0

Major update: projects, components, templates, advanced right-click menus, add-ins, batch processing and an AI assistant. Workbooks produced with 1.0.0 open and save as before.

### What's new

- **`.cuis` projects.** A ribbon can be saved independently of any workbook, reopened, then applied to the workbook of your choice ("Apply" button).
- **Undo / redo.** Every change can be undone (Ctrl+Z) and redone (Ctrl+Y).
- **Contextual tabs and Backstage.** The tabs Excel shows for a table, a chart or a picture can be extended in the editor, and the File menu (Backstage) is built element by element.
- **Compatibility.** Choose the targeted Excel version: each command, icon and menu has a "verified", "documented" or "unknown" badge, based on Microsoft's official lists and on checks made in Excel. A report sums up the project.
- **Components and library.** A group, a tab or a menu is stored with its VBA code in a `.cuic` component, reusable in any project. Components are provided.
- **Business templates.** Inventory, sales management and dashboard, with their ribbon, VBA code and the "Prepare the workbook" macro.
- **Advanced right-click menus.** The 89 menus of Microsoft's official list; entries shown only where they make sense (sheet, table, column, non-empty cell…), with generated VBA code; copying an entry to other menus; submenus.
- **`.xlam` add-ins.** Opening, editing and saving an add-in; generated VBA code adapted to add-ins; guided procedure to create one in Excel.
- **Apply to several workbooks** (desktop version). Each file is checked first, copies go to an output folder, with a CSV report; originals are never modified.
- **Controlled AI assistant.** Describe what you want: CustomUI Studio prepares a request for the AI tool of your choice, then checks its answer with its own rules before any change. The application connects to nothing and sends nothing.

### Fixes

- A native Excel tab or group without a label is no longer renamed on export.
- A line break in a tooltip is kept.
- An attribute a control does not accept is reported instead of being silently dropped.
- A file whose extension does not match its content (for example an `.xlsm` renamed to `.xlam`) is refused with an explanation, as Excel would.

## 1.0.0

First stable release. Includes the same features as version 0.8.0.

- **Microsoft Store.** An installable version is currently being published to the Microsoft Store. It will be signed by Microsoft and will install without triggering any Windows warnings.
- **Privacy Policy.** Available in [CONFIDENTIALITE.md](CONFIDENTIALITE.md): the application does not collect or transmit any data and functions without an Internet connection.
- The portable version remains available—free of charge and requiring no installation.
- Workbooks created with previous versions can be opened and saved just as before.

## 0.8.0

### New Features

- **Interface in French or English.** Upon first launch, the application asks for your preferred language; a dropdown menu at the top of the window allows you to change it at any time. Your selection is saved.
- **VBA-controlled properties.** In a control's properties, select which attributes should be dynamic (changeable while the workbook is in use)—such as the label, enabled/disabled (grayed out) state, visibility, tooltips, or image. In your macros, a single line is all it takes:

  ```vb
  SetControlState "btnExporter", "enabled", False
  ```

  The generated `RibbonDESIGNER` module tracks the current value and updates the ribbon accordingly. Until you call `SetControlState`, the value specified in the editor remains in effect.
- **Drag-and-drop in the Ribbon Structure.** Tabs, groups, and controls can be rearranged using the mouse—placed before, after, or inside another element. The application prevents actions that Excel would not allow and explains why. The "move up/down" arrows remain available.
- **Right-click menus: hiding additional Excel commands.** Certain commands (Paste Special, Insert, Clear Contents, Filter, Sort, etc.) cannot be hidden via the ribbon itself. These are labeled "VBA"; when you deselect them, the `RibbonDESIGNER` module includes the procedures `MasquerCommandesNatives` (Hide Native Commands) and `RestaurerCommandesNatives` (Restore Native Commands). This hiding effect applies across all of Excel while the workbook is active; three short procedures—provided at the top of the module—can be pasted into `ThisWorkbook` to restore these commands for other workbooks.
- **More descriptive VBA code.** The generated procedures include commented examples for common use cases. - **Right-click preview.** An 'X' allows you to close it and return to the ribbon.

### Corrections

- **Image deleted from the library:** It is now removed from the workbook upon saving, along with its associated link. Previously, it would remain in the archive.
- **Image re-imported under the same name:** The workbook now correctly receives the new image.
- **Image used by a control:** Deleting such an image now triggers a confirmation prompt that identifies the affected controls.
- **Format change (CustomUI 2006 to 2009, or vice versa):** The legacy ribbon file no longer remains in the workbook.
- **"Issues" panel:** Text has been clarified.

Properties managed via VBA, hiding elements via VBA, and workbooks from which images have been deleted or replaced have been tested in Excel for Microsoft 365 (version 16.0), in the French-language version.

## 0.6.0

### New Features

- **XML validation using Microsoft schemas.** The XML editor continuously validates your ribbon—offline—against the 2006 and 2009 CustomUI schemas.
  - "Issues" panel: each error is listed with an explanation and, where possible, a suggested fix. Clicking an item moves the cursor directly to the corresponding location.
  - Errors are highlighted within the editor.
  - Saving to a workbook or exporting as XML is blocked as long as any errors remain, preventing the common issue where a ribbon disappears in Excel without explanation.
- **Right-click menus: hide Excel commands with a single click**—for cells, row headers, column headers, and tables.
  - A list of Excel commands, each with its own checkbox.
  - Presets: "Minimal menu," "No copy/paste," "No insert/delete," and "Show all."
  - A preview displays the menu exactly as it will appear in Excel. Click any Excel command to hide it, or to insert a button immediately before or after it.

### Fixes

- **Generated VBA code:** The following issues previously prevented the ribbon from displaying or triggered a "Type mismatch" error in Excel:
  - `get…` procedures (e.g., `getLabel`, `getEnabled`, `getVisible`) and handlers for intercepted Excel commands were incorrectly declaring a return type.
  - `getVisible` and `getEnabled` defaulted to `False`, causing controls to disappear or appear grayed out.
  - `loadImage` was being generated as a function.
  - The `onAction` callback for drop-down lists or galleries failed to receive the index of the selected item.
  - The `onChange` callback for text boxes used an incorrect parameter name.
- **Toggle buttons and checkboxes:** Each control now maintains its own independent state (previously, they all shared the same state).
- **Attributes disallowed by Excel:** The editor no longer offers 26 attributes that are prohibited by the format (for example, assigning an image to a checkbox or a label to a split button)—previously, such attributes would cause Excel to reject the entire ribbon.
- **Redundant static attributes and callbacks** (e.g., `label` combined with `getLabel`, or `visible` with `getVisible`): The editor now writes only the callback and provides a notification to the user.
- **Right-click preview:** Custom images are now displayed.
- **Group or tab identifiers:** The Properties panel no longer loses the current selection while the user is typing.
- **VBA-loaded images:** Ribbons that utilize `loadImage` can now be saved without embedding the actual image files.

The generated VBA code, the handling of prohibited attributes, and the right-click menus have been tested in the French version of Excel for Microsoft 365 (version 16.0).

## 0.4.0

### New Features

- **Context Menus (Right-click menus):** A new "Context Menus" section has been added to the Ribbon Structure.
  - Select which Excel menu to extend: cell, row header, column header, structured table (list object), pivot table, shape, etc.
  - Add buttons, checkboxes, submenus, separators, etc.—only controls permitted within a right-click menu are available for selection.
  - Adjustable positioning: place items before or after an existing Excel command (Cut, Copy, Paste, etc.), or at the end of the menu.
  - Includes a preview of the right-click menu, automatic generation of VBA callbacks, and the ability to reload/read back from the XML editor.
  - CustomUI 2009 format (for Office 2010 and later): the application offers to upgrade your project if it is currently using the 2006 format.

### Fixes

- **No data loss during XML synchronization:** Previously, context menus, the Backstage view, and contextual tabs within a workbook were removed when clicking "Synchronize." They are now preserved and rewritten in the locations required by Office.
- **Control ID changes:** The Properties pane no longer loses the selected control while its ID is being edited.

## 0.3.0

Initial public release.

- Visual Ribbon builder: only elements permitted by Excel are offered.
- Supports CustomUI 2006 (Office 2007) and 2009 (Office 2010 and later) formats.
- Live preview, including lists, menus, and tooltips.
- Includes an `imageMso` library, support for custom images, and a catalog of native Excel commands.
- Generates a `RibbonDESIGNER` VBA module, which can be imported using the `ImporterModuleRibbonDESIGNER` macro.
- Synchronized XML editor with full validation before saving.
- Workbook structure remains unchanged: only the Ribbon is modified; your worksheets, data, and VBA macros remain intact.

# Excel Automation Tools

This repository contains a suite of Excel macros to streamline and automate your workbook workflows. The tools are designed for dynamic, evolving datasets where frequent updates and comparisons are necessary.

---

## Overview
This repository contains a suite of Excel macros to streamline and automate your workbook workflows. The tools are designed for dynamic, evolving datasets where frequent updates and comparisons are necessary.

- **ImportTSV**: Imports a TSV file into an open workbook, renaming the sheet to a user-specified date.
- **SheetSetup**: Prepares an Excel sheet by inserting columns, copying data, and applying formulas.
- **UpdateResults**: Updates formulas in the `Results` sheet to reference the latest data sheet.

## 🔧 Tools Included

### **SheetSetup**
Prepares a new data sheet for analysis.

**Features:**
- Inserts a "helper" column that combines columns B and C.
- Copies column I to AD (full copy).
- Copies column J to AF (values only).
- Adds a VLOOKUP-based p-code column from the "P-code" sheet.
- Performs a cross-sheet VLOOKUP from the most recent previous data sheet and places results in column J.
- Centers text in column J.
- Freezes the top row and applies filters.
- Hides columns S through AC.
- Applies conditional formatting:
  - Green highlight in H-K when column J contains "invoice".
  - Red highlight in H-K when column I ≠ column AD.
  - Red highlight in column O when Actual Material > Estimated Material.
  - Yellow highlight in column B when a value appears in the "rev rec" sheet.

### **ImportTSV**
Imports a `.tsv` file into a new sheet in the current workbook.

**Features:**
- Prompts user to select a `.tsv` file.
- Asks the user to input a date, then renames the sheet to `data - m-d-yyyy`.
- Automatically applies General formatting to imported columns.
- Inserts a helper column to the left of column A.

### **UpdateResults**
Smartly updates formulas in the `Results` sheet when a new data snapshot is added.

**Features:**
- Scans all sheet names to find the two most recent sheets named in the format `data - m-d-yyyy`.
- Prompts the user to confirm or adjust the old and new sheet names.
- Searches the `Results` sheet for formulas referencing the old sheet and replaces them with the new sheet.
- Reports how many formulas were updated.
- Locale-agnostic: date parsing is handled explicitly using `DateSerial`.

---

## 💡 How to Use

1. Open Excel and press `ALT + F11` to open the **VBA editor**.
2. Insert a new Module via `Insert > Module`.
3. Paste in the desired macro(s).
4. Save your workbook as a **Macro-Enabled Workbook (.xlsm)**.
5. Run the macros via `ALT + F8`.

### ➕ Adding a Ribbon Button for Easy Access

To add a macro to the **Excel Ribbon** for quick execution:

1. Open Excel and go to **File > Options**.
2. Choose **Customize Ribbon** from the sidebar.
3. On the right side, create a **New Tab** (or use an existing one).
4. Select **Macros** from the dropdown on the left.
5. Choose the desired macro (e.g., `SheetSetup`, `ImportTSV`, `UpdateResults`) and click **Add >>** to include it in your custom tab.
6. With the macro selected on the right, click **Rename** to give it a friendly name and pick an icon.
7. Click **OK**. Your macro will now appear in the Ribbon under your custom tab.

---

## 📁 Recommended Structure
- Store recurring lookup data in the `P-code` and `rev rec` sheets.
- Name all incoming data sheets using the `data - m-d-yyyy` convention.
- Always run `SheetSetup` after importing a new `.tsv` to prepare the sheet for downstream logic.

---

## ✅ Best Practices
- Avoid manually renaming data sheets — use the ImportTSV tool for consistency.
- Run `UpdateResults` whenever a new sheet is added to keep the `Results` sheet in sync.
- Lock or protect reference sheets (`P-code`, `rev rec`) if they're used by multiple tools.

---

For questions or to request enhancements, open an issue or submit a Pull Request! 💬

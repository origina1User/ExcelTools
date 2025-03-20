# Excel Macro Tools

## Overview
This repository contains VBA macros designed to automate common Excel data processing tasks. The tools included are:

- **ImportTSV**: Imports a TSV file into an open workbook, renaming the sheet to a user-specified date.
- **SheetSetup**: Prepares an Excel sheet by inserting columns, copying data, and applying formulas.

## Features of ImportTSV
- **Prompts user** to select a TSV file and imports its contents into a new sheet.
- **Renames the sheet** based on a user-specified date (format: "data - m-d-yyyy").
- **Copies data** from column I to AD using full copy (includes formatting and formulas).
- **Copies data** from column J to AF as values only.
- **Names column AE** as "p code".
- **Applies an optimized VLOOKUP formula** in column AE to retrieve data from a lookup sheet (`P-code`).
- **Ensures dynamic row handling**, adjusting automatically to the last row of data.

## Features of SheetSetup
- **Inserts a new column** to the left of column A and names it "helper".
- **Populates the new column** with a formula that concatenates values from columns B and C.
- **Copies data** from column I to AD using full copy (includes formatting and formulas).
- **Copies data** from column J to AF as values only.
- **Names column AE** as "p code".
- **Applies an optimized VLOOKUP formula** in column AE to retrieve data from a lookup sheet (`P-code`).
- **Ensures dynamic row handling**, adjusting automatically to the last row of data.

## Installation & Usage
### 1. Importing the Macro
1. Open Excel.
2. Press `ALT + F11` to open the **VBA Editor**.
3. Go to **Insert** > **Module**.
4. Copy and paste the `SheetSetup` macro into the module.
5. Save the workbook as a **Macro-Enabled Workbook (.xlsm)**.

### 2. Running the Macro
1. Press `ALT + F8`.
2. Select `SheetSetup` and click **Run**.

## Adding a Ribbon Button for Easy Access
To add the macro to the **Excel Ribbon** for quick execution:
1. **Open Excel Options**: Click **File** > **Options** > **Customize Ribbon**.
2. **Create a New Tab**: On the right side, click **New Tab**, then rename it (e.g., "Custom Tools").
3. **Add a Button**:
   - Select **Macros** from the left dropdown.
   - Choose `SheetSetup` and click **Add**.
4. **Modify the Button** (Optional):
   - Click **Modify**, choose an icon, and rename it (e.g., "Setup Sheet").
5. Click **OK**, and the macro will now be accessible from the Ribbon.
6. Repeat for any additional macros.

## Notes
- If the macro encounters an issue, check for missing data or incorrect column references.

## Contributing
Feel free to open issues or submit pull requests for improvements!

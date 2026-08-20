# Bulk File & Folder Renamer

A Python automation that combines an Excel-controlled rename workflow with validation, safety checks, and execution logging.

## Problem Statement

Renaming a large number of files and folders manually is repetitive and error-prone. Users may need to inspect the contents of a folder, decide new names, rename each item individually, and verify that no existing item is accidentally overwritten.

This project provides a controlled workflow:

```text
Select Folder
    ↓
Export Folder Contents to Excel
    ↓
Enter New Names in Excel
    ↓
Validate Rename Requests
    ↓
Review Summary / Warnings
    ↓
Confirm
    ↓
Bulk Rename Files & Folders
    ↓
Generate Execution Log
```

## Key Features

- Extracts direct files and folders from a selected directory into an Excel control file.
- Uses Excel as a user-friendly interface for defining new names.
- Preserves the original file extension when the new name is entered without an extension.
- Detects explicit extension changes and asks for confirmation.
- Validates Windows-invalid characters.
- Checks Windows reserved device names such as `CON`, `PRN`, `AUX`, and `NUL`.
- Detects blank names and unchanged names.
- Detects duplicate requested target names.
- Prevents overwriting an existing target.
- Re-checks source and target paths immediately before renaming.
- Handles permission, missing-file, and filesystem errors.
- Provides a validation summary before execution.
- Blocks the entire rename operation when validation failures exist.
- Generates an Excel execution log with status, messages, timestamps, source paths, and target paths.
- Generates a summary of renamed, skipped, failed, warning, and not-found items.
- Provides a Tkinter-based folder/file selection experience.

## Technologies

- Python
- `tkinter`
- `pathlib`
- `openpyxl`
- `datetime`
- `re`
- `sys`
- `os`

## Project Structure

```text
Bulk-File-Folder-Renamer/
│
├── Bulk_Rename.py
├── README.md
├── requirements.txt
│
├── sample_data/
│   ├── Rename_Control_Demo.xlsx
│   ├── Project_Report.xlsx
│   ├── Monthly_Sales.xlsx
│   ├── Customer_Data.xlsx
│   ├── Old_Presentation.pptx
│   ├── Notes.txt
│   ├── Archive_2025/
│   └── Pending_Review/
│
└── docs/
    └── workflow.md
```

## Installation

```bash
pip install -r requirements.txt
```

## Usage

Run:

```bash
python Bulk_Rename.py
```

The application provides three choices:

```text
1. Extract folder names to Excel
2. Rename using Excel control file
3. Exit
```

### Option 1 — Extract Folder Contents

Select the folder containing the files/folders to be renamed.

The program creates an Excel control file containing:

- Current Name
- Type
- Extension
- New Name

Enter the desired replacement names in the `New Name` column.

### Option 2 — Rename Using Excel Control File

Select:

1. The target folder.
2. The completed Excel control file.

The program validates every requested rename before making any changes.

## Safety Design

The automation deliberately uses a validation-first approach.

If any row has a validation failure, **no rename operation is executed**. A validation report is still generated so the user can correct the control file and run the process again.

The automation also does not overwrite an existing file or folder.

## Extension Handling

If the existing file is:

```text
Report.xlsx
```

and the Excel control file contains:

```text
Report_2026
```

the resulting name becomes:

```text
Report_2026.xlsx
```

If the user explicitly enters:

```text
Report_2026.pdf
```

the program warns that changing the extension only changes the filename; it does not convert the underlying file format.

## Validation

The project validates:

- Blank current names
- Blank new names
- Missing source items
- File/folder type mismatches
- Invalid Windows characters
- Control characters
- Trailing spaces or periods
- Windows reserved names
- Duplicate requested target names
- Existing target paths
- Unchanged names
- Extension changes

## Logging

After execution, the program creates an Excel log containing:

- Excel Row
- Current Name
- New Name
- Type
- Status
- Message
- Extension Change
- Timestamp
- Source Path
- Target Path

A summary sheet provides counts by status.

## Example

### Before

```text
Project_Report.xlsx
Monthly_Sales.xlsx
Customer_Data.xlsx
Archive_2025/
```

### Excel Control

```text
Project_Report.xlsx  → Project_Report_2026
Monthly_Sales.xlsx   → Monthly_Sales_August
Customer_Data.xlsx   → Customer_Master
Archive_2025         → Archive_2026
```

### After

```text
Project_Report_2026.xlsx
Monthly_Sales_August.xlsx
Customer_Master.xlsx
Archive_2026/
```

## Portfolio Value

This project demonstrates:

- Process automation
- Python scripting
- Excel integration
- GUI-based file selection
- Input validation
- Defensive programming
- Filesystem operations
- Error handling
- Audit logging
- User safety controls

## Limitations

- The automation operates on the selected folder's direct children; it does not recursively rename files inside nested folders.
- The Excel control file must contain the expected column names.
- Changing a file extension does not convert its format.
- The application is intended for Windows filesystem naming rules.

## Future Improvements

Potential future enhancements:

- Recursive folder processing
- Dry-run mode
- Undo/rollback capability
- Preview GUI
- Regex-based bulk renaming
- Search-and-replace rules
- Rename templates
- Advanced duplicate handling
- Packaged `.exe` distribution

## Disclaimer

The files in `sample_data/` are fictional demonstration files created for portfolio use. Do not use the automation on production data without reviewing the rename list and validating the target folder.

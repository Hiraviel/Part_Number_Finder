# Part Number Finder

Tkinter GUI to search for a specific text (e.g. a part number) inside Excel files.

Enter a part number and a folder, click **Buscar**, and the script walks through the folder (including subfolders), opens every `.xls` / `.xlsx` file, and reports which files contain an exact match. The results are shown in the GUI and also saved to a text file.

## Features

- **Recursive search** — scans the selected folder and all its subfolders.
- **Exact match** — compares each cell value with the searched text, case-insensitive.
- **Results in the GUI** — every file containing a match is listed in the results panel, with a running count.
- **Results file** — a `<PART_NUMBER>.txt` file is generated in the folder where the script was launched, containing the paths of all matching files.
- **Match count popup** — a dialog shows the total number of elements found.
- **Automatic cleanup** — before searching, temporary Excel lock files (names starting with `~` or `$`) are removed from the folder.
- **Folder picker** — a **Folder** button opens a directory chooser, so the path does not have to be typed.

## Requirements

- Python 3 with Tkinter (included with standard Python on Windows)
- The `xlrd` library:

```bash
pip install xlrd
```

> Note: `xlrd` 1.2.0 is the last version that reads `.xlsx` files. `xlrd` 2.x only supports the older `.xls` format, so use 1.2.0 if your files are `.xlsx`.

## Usage

```bash
python GUI.py
```

1. **Introduce el número de parte** — type the part number (or any text) to find.
2. **Folder donde buscar** — type the folder path, or click **Folder** to browse for it.
3. Click **Buscar**.

When the search finishes, a popup shows how many elements were found, the matching files are listed in the results panel, and a `<PART_NUMBER>.txt` file is created next to the script.

## Notes

- Only `.xlsx` and `.xls` files are searched, due to the limitations of the `xlrd` library.
- Close the Windows Preview Pane before running a search: an open preview can lock a file and crash the script.
- The search appears to do nothing while it runs — there is no progress indicator. Wait for the results popup.
- You can paste text into the GUI fields, but you cannot copy text out of them.
- When pasting a folder path, use `/` instead of `\`.

## Files

| File | Description |
| --- | --- |
| `GUI.py` | The application (Tkinter GUI + search logic) |
| `README.md` | This file |

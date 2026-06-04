# Shuttle Manifester

Shuttle Manifester is a Python desktop application built with Tkinter.

It helps format passenger shuttle data copied from Excel into a clean manifest list. The app is designed to make passenger manifesting faster, cleaner, and more consistent by reducing manual formatting and helping identify duplicate passenger entries.

---

## What It Does

Shuttle Manifester allows users to:

- Paste passenger data copied from Excel
- Select a formatting profile
- Automatically generate a formatted shuttle manifest
- Start seat numbering from a chosen number
- Copy the finished manifest to the clipboard
- Copy the shuttle service label with the total seat count
- Check for duplicate passengers
- Generate a cleaned manifest after resolving duplicates
- Manage different manifest profiles
- Switch between different visual themes

---

## What It Is For

This app is useful when passenger transport information is received in spreadsheet form and needs to be converted into a standard text format.

Instead of manually rewriting each passenger line, the user can paste the raw Excel data into the app and let the formatter produce a consistent manifest output.

Example output:

```text
Location A > Location B @ 1840 on 04/06/2026
Seat Number 01 John Smith 0400000000 QF123 @ 1840 1x Bags
Seat Number 02 Jane Citizen 0400000001 VA456 @ 1915 2x Bags
```

---

## Main Features

### Excel Paste Formatting

The app reads tab-separated data copied from Excel.

It uses the first row as the header row, then matches those headers against the selected profile.

Each passenger row is then converted into a formatted seat line.

---

### Profile-Based Formatting

Profiles control how the app reads and formats data.

Each profile contains:

- Expected Excel column headers
- A group label format
- A passenger output format

This means different spreadsheet layouts can be supported without rewriting the main code.

Example profiles include:

- Defence Travel
- Toll
- Toll 2

---

### Duplicate Checker

The duplicate checker compares new passenger entries against an existing manifest.

It can detect:

- Confirmed duplicates
- Possible duplicates
- Entries that may need merging

The checker mainly compares:

- Passenger names
- Phone numbers
- Flight details
- Missing or partial information

This helps prevent the same passenger from being added twice.

---

### Clean Manifest Generator

After checking for duplicates, the app can generate a clean final manifest.

The user can choose whether to:

- Keep the new entry
- Keep the existing entry
- Keep both
- Remove an entry
- Accept a merge suggestion

The final output is then renumbered automatically.

---

### Profile Manager

The Profile Manager allows users to manage formatting profiles inside the app.

Users can create, edit, delete, export, import, or reset profiles.

This makes the app flexible when spreadsheet formats change.

---

### Themes

The app includes multiple themes, such as:

- Default
- Default Blue
- Dark Mode
- OLED Black
- Rose Pine Dawn
- Solarized Light

Themes change the colours of the interface without changing the app logic.

---

## How It Works

At a high level, the app works in five steps:

1. The user copies passenger data from Excel.
2. The user pastes the data into the input box.
3. The app matches the Excel headers to the selected profile.
4. The app formats each row into a manifest seat line.
5. The user copies the final manifest or runs a duplicate check.

---

## Technical Overview

The app is written in Python and uses:

- `tkinter` for the graphical interface
- `json` for storing profile configuration
- `re` for pattern matching and duplicate detection
- `Pillow` for loading image assets such as logos and icons

The app also creates or reads a `config.json` file, which stores the formatting profiles and settings.

---

## Requirements

Install Python 3.

Install Pillow:

```bash
pip install pillow
```

Tkinter usually comes included with Python on Windows.

---

## Running the App

Run the Python file:

```bash
python ShuttleManifestor.py
```

---

## Example Workflow

1. Open the app.
2. Choose the correct profile.
3. Copy passenger data from Excel, including the header row.
4. Paste it into the input box.
5. Set the starting seat number.
6. Review the formatted output.
7. Copy the manifest.
8. Run the duplicate checker if comparing against an existing manifest.

---

## Why This Project Exists

This project was created to reduce repetitive manual formatting work.

It turns spreadsheet-based passenger data into a clean, consistent manifest format and adds extra tools to help with duplicate checking and final manifest cleanup.

---

## Future Improvements

Possible future improvements could include:

- Drag-and-drop file support
- Direct Excel file upload
- Export to `.txt` or `.docx`
- Better duplicate scoring
- Saved manifest history
- More advanced profile validation

---

## Author

Created by Thomas Brayovic.

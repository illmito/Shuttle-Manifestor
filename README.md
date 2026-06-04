# Shuttle Manifester

Shuttle Manifester is a Python desktop application built with Tkinter.  
It helps format passenger shuttle data copied from Excel into a clean manifest list.

The app is designed to make passenger manifesting faster, cleaner, and more consistent by reducing manual formatting and helping identify duplicate passenger entries.

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

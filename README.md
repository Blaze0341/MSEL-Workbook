# MSEL Inject Planning Tool

## Overview

This offline **MSEL (Master Scenario Event List) Inject Planning Tool** allows users to manage, edit, import, and export injects for exercise planning. It mimics the functionality of the Excel-based MSEL workbook but runs entirely in a web browser using local HTML and JavaScript.

The tool is designed for **offline use**, supports data persistence via **local storage** and file-based backups, and includes features like:

* Sorting and filtering
* PDF export
* Dynamic field interaction
* Line number override
* Script management
* Training data tagging

---

## Contents

1. [System Requirements](#system-requirements)
2. [Installation Instructions](#installation-instructions)
3. [Getting Started](#getting-started)
4. [Main Interface Overview](#main-interface-overview)
5. [How to Add or Edit Injects](#how-to-add-or-edit-injects)
6. [Saving and Backup Options](#saving-and-backup-options)
7. [Importing and Exporting Data](#importing-and-exporting-data)
8. [Script Management](#script-management)
9. [Unit Training Information](#unit-training-information)
10. [Special Features](#special-features)
11. [Table Sorting and Line Number Override](#table-sorting-and-line-number-override)
12. [Troubleshooting](#troubleshooting)

---

## System Requirements

* Any modern web browser (Chrome, Edge, Firefox)
* No internet connection required after download
* A local copy of the MSEL Tool folder

## Installation Instructions

### Step 1: Download the Tool

1. Open Microsoft Teams.
2. Click into **Bullpen Coordination** > **General** > **Files**.
3. Find the **MSEL Workbook** folder.
4. Click the ellipsis (`...`) and select **Download**.
5. Save the folder to a local location (e.g., `Documents/MSEL_Tool`).

### Step 2: Extract and Run

1. If zipped, right-click the folder and choose **Extract All**.
2. Open the extracted folder.
3. Double-click `msel.html` to launch the tool.

---

## Getting Started

Once `msel.html` is open, you'll see the main dashboard with:

* **Data Controls:** Save, backup, load, clear
* **Export Options:** Export to PDF/Excel
* **Planning Tools:** Add injects, create scripts, add training info
* **File Upload:** Import Excel or CSV inject lists

Below is a live-updating inject table. Click a row to edit.

---

## Main Interface Overview

### Key Features

* Interactive inject table with **sortable columns**
* Form-based **Inject Entry Panel**
* **Script Editor** and **Training Tool**
* **Persistent local storage**
* Override Auto Line Number feature
* **Script dropdown integration**
* Live PDF export for printing

---

## How to Add or Edit Injects

### Adding Injects

1. Click **Add New Inject**.
2. Fill in the form fields (Line#, Date, Time, etc).
3. Leave Line# blank to auto-generate based on Date/Time.
4. Check **Override Auto Line Number** to pin a specific Line#.
5. Click **Add and Close** or **Add and New**.

### Editing Injects

* Click a row to edit.
* Make changes and click **Save** or **Save and Close**.
* Use **Previous / Next** buttons to navigate.

---

## Saving and Backup Options

### Auto-Save

* Uses your browser's `localStorage`
* Works offline; data persists across sessions

### Manual Save

* Click **Save Now** to commit changes

### Backup (JSON)

* Click **Save Backup** to export data as `.json`
* Includes injects, scripts, overrides

### Restore from Backup

* Use **Load Save** to load a `.json` file

### Reset

* Click **Clear Saved Data** to reset everything

---

## Importing and Exporting Data

### Import

* Click **Import File** to load `.xlsx` or `.csv`
* Specify the row with headers when prompted

### Export

* **Export to Excel:** Full inject list as `.xlsx`
* **Export to MSEL PDF:** Printable injects
* **Export Scripts to PDF:** Print-ready script listing

---

## Script Management

* Open the script editor via **Create Scripts**
* Enter script title and body
* Scripts stored in browser memory
* Scripts are selectable via dropdowns when assigning to injects
* Export all scripts to PDF

---

## Unit Training Information

* Click **Add Unit Training Info** to begin
* Input unit details and training objectives
* Tag injects via notes or fields

---

## Special Features

* **Line# Override**
* **Hidden Columns** for internal state tracking
* **Auto Numbering** based on chronological order
* Clickable **column headers** for sorting
* **Prev/Next** navigation inside form
* **Multi-mode adding** (Add & Close / Add & New)
* **Script dropdowns auto-populate** from saved scripts

---

## Table Sorting and Line Number Override

* Click column headers (e.g., Date, Time) to sort
* Default sort: Date → Time
* When **Override Auto Line Number** is checked:

  * That inject keeps its Line#
  * All others auto-adjust around it

---

## Troubleshooting

| Issue                | Solution                                              |
| -------------------- | ----------------------------------------------------- |
| Changes not saving   | Ensure local storage is not blocked                   |
| Injects disappeared  | Re-import or load a backup JSON                       |
| Import failed        | Ensure clean column headers and no merged Excel cells |
| Override not working | Make sure override is checked before saving           |
| Missing checkboxes   | Refresh page to reinitialize UI                       |
| Reset completely     | Use **Clear Saved Data** and reload                   |

---

## Support

If you encounter issues or have feature requests, please contact the maintainer or use your team’s shared collaboration channel.

**Last Updated:** May 2025

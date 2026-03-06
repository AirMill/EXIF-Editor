EXIF Metadata Editor

A fast, desktop batch editor for EXIF, IPTC, and XMP metadata built with Python + PyQt6 and powered by ExifTool.

This application allows photographers, real estate agents, and content creators to easily inspect and edit metadata for large groups of images while maintaining safe workflows and backups.

Features
Batch Metadata Editing

Edit metadata for one or multiple files simultaneously.

Supported formats include:

JPEG / JPG

TIFF / TIF

PNG

HEIC

WEBP

RAW formats (NEF, CR2, CR3, ARW, DNG, RAF, ORF, RW2, SRW, etc.)

The application scans folders and optionally includes subfolders.

Editable Metadata Fields
Common Metadata

Creator / Author

Description

Title

Keywords

Copyright

Contact Email

Contact Phone

Website

Date Taken

Camera / EXIF

Camera Make

Camera Model

Lens Model

Serial Number

ISO

F-Number

Exposure Time

Focal Length

Exposure Compensation

Presets System

Create reusable metadata cards that contain contact information and copyright data.

Example preset:

Agent: Anna Kowalska
Email: anna@agency.com
Phone: +48 123 456 789
Website: https://agency.com
Copyright: © 2026 Agency Name

Presets can be applied to selected photos instantly.

Keyword Handling

Keywords entered as:

kitchen, living room, exterior

are automatically written as separate keywords in metadata.

Multi-File Editing Support

When multiple files are selected:

Fields that differ display (mixed values)

Changes are applied only to fields where Apply is checked

RAW File Protection

If RAW files are detected:

The application warns the user before writing metadata to prevent accidental modifications to RAW data.

Automatic Undo Backup

Before writing metadata the app automatically creates a backup.

Backup location:

_UNDO_METADATA_BACKUP/YYYYMMDD_HHMMSS/

This allows safe rollback of metadata changes.

Restore Backup

The Restore Backup feature restores files from the most recent backup.

This provides an easy Undo capability.

Batch Progress Indicator

Large operations display a progress bar showing:

number of processed files

current file being written

Professional UI

Dark green interface

Orange section titles

White text for readability

Modern tab layout

Large action buttons

Designed for comfortable long editing sessions.

Installation
1 Install Python

Python 3.10+ recommended

Download from:

https://python.org

2 Install Dependencies
pip install PyQt6
3 Install ExifTool

Download from:

https://exiftool.org

Place exiftool.exe in:

tools/exiftool.exe

inside the project folder.

Example structure:

metadata-editor/
│
├── Index_4.py
├── README.md
└── tools/
    └── exiftool.exe
Running the Application
python Index_4.py
Building a Windows Executable

Install PyInstaller:

pip install pyinstaller

Build the application:

pyinstaller --onefile --windowed Index_4.py

The executable will appear in:

dist/
Why ExifTool

ExifTool is one of the most reliable metadata libraries available and supports hundreds of formats.

This application uses ExifTool for:

reading metadata

writing metadata

maintaining compatibility across formats

Safety Notes

Editing metadata inside RAW files is technically safe with ExifTool, but some workflows prefer using XMP sidecar files.

The application warns users before editing RAW formats.

License

MIT License

You are free to use, modify, and distribute this software.

Nikolai Shimukovich

Created as a workflow tool for photography and real estate media management.
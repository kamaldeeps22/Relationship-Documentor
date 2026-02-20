RELATIONSHIP DOCUMENTOR - RESOURCES FOLDER
==========================================

This folder contains source image files for the plugin icons.

The actual icons used in the plugin are Base64-encoded strings
embedded directly in RelationshipDocumentorPlugin.cs file.

FILES:
------
Logo_32x32.png  - Small icon (32x32 pixels) - used in toolbar
Logo_80x80.png  - Large icon (80x80 pixels) - used in plugin tile

USAGE:
------
These PNG files are kept here for:
1. Reference - original source files
2. Updates - when you need to update the plugin icons
3. Documentation - for GitHub repository

TO UPDATE ICONS:
----------------
1. Edit the PNG files with your preferred image editor
2. Run: ConvertToBase64.ps1 -ImagePath "Logo_32x32.png"
3. Copy the output Base64 string
4. Paste into RelationshipDocumentorPlugin.cs ExportMetadata attribute

BUILD SETTINGS:
---------------
Build Action: None
Copy to Output Directory: Do not copy

These files are NOT deployed with the plugin - they are source files only.

---
Kamaldeep Singh
https://kamaldeepsingh.com
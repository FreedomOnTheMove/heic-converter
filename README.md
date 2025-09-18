# HEIC Converter Tool

Convert HEIC/HEIF images to JPG format with optional file renaming from Excel spreadsheets.

## Features

- Convert single or batch HEIC files to JPG
- Optional Excel/CSV file upload for automatic file renaming
- Drag & drop file support
- Folder upload support
- Download results as ZIP file
- Progress tracking for batch operations
- Browser-based - no server required

## Usage

1. **Optional**: Upload Excel/CSV file with filename mappings (columns Q and R)
2. Select HEIC files or folders containing HEIC files
3. Files are automatically converted and zipped for download
4. Non-HEIC image files are included as-is in the output

## Excel File Format

If using Excel renaming:

- Column Q: Original filename
- Column R: New filename
- First sheet is used by default

## Deployment

Build static version:

```
bash yarn build:static
``` 

Upload contents of `static-build/` folder to any static hosting service.

## Browser Support

Works in modern browsers that support ES modules and File API. No server-side processing required.


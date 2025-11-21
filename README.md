# PowerShell Utility Scripts

This repository contains PowerShell scripts designed to automate file compression and document conversion tasks.

## Prerequisites

To run these scripts successfully, you must have the following installed:
* **Windows PowerShell** (Pre-installed on most Windows systems)
* **7-Zip:** Required for the folder compression script. [Download 7-Zip here](https://www.7-zip.org/).
* **Microsoft PowerPoint:** Required for the PDF conversion script (uses COM objects). [Microsoft Office Information](https://www.microsoft.com/en-us/microsoft-365/powerpoint).

---

## 1. Zip Folders Individually (`Zip_Folders_Individually.ps1`)

### Description
This script scans a user-specified directory for immediate sub-folders and compresses each folder into its own distinct `.zip` archive using 7-Zip.

### Features
* **Interactive Interface:** Color-coded prompts guide you through the process.
* **Batch Processing:** Automates the archiving of multiple folders at once.
* **Location:** The resulting `.zip` files are saved in the same root directory as the original folders.

### Configuration
**Important:** The script assumes 7-Zip is installed at the following default path:
`C:\Program Files\7-Zip\7z.exe`

If your installation is in a different location, open the script in a text editor and modify the `$7zipPath` variable at the top of the file.

### Usage
1.  Run the script.
2.  Paste the full path of the directory containing the folders you wish to zip.
3.  Review the list of folders found.
4.  Select **Yes** to begin compression.

---

## 2. PowerPoint to PDF Converter (`pptx_to_pdf.ps1`)

### Description
This script recursively searches through a source directory for PowerPoint files (`.ppt` and `.pptx`) and converts them to PDF format. It maintains the original folder structure in the destination directory.

### Features
* **Recursive Scan:** Finds files in the main folder and all sub-folders.
* **Structure Mirroring:** Recreates the source folder hierarchy in the destination folder so files remain organized.
* **Auto-Creation:** If the destination folder does not exist, the script offers to create it for you.

### Usage
1.  Run the script.
2.  Enter the **Source Directory** path (where your PowerPoint files are located).
3.  Enter the **Destination Directory** path (where you want the PDFs saved).
4.  Select **Yes** to begin the batch conversion.

---

## Troubleshooting
* **Execution Policy:** If the scripts fail to run, you may need to adjust your PowerShell execution policy. Run `Set-ExecutionPolicy RemoteSigned` in an Administrator PowerShell window.
* **7-Zip Error:** If `Zip_Folders_Individually.ps1` closes immediately or shows an error, verify the path to `7z.exe` in the script configuration matches your actual installation.

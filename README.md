# PowerShell Utility Scripts

This repository contains PowerShell scripts designed to automate file compression and document conversion tasks.

---

## 1. Zip Folders Individually (`Zip_Folders_Individually.ps1`)

### Description
This script scans a user-specified directory for immediate sub-folders and compresses each folder into its own distinct `.zip` archive. The resulting files are saved in the same root directory as the original folders.

### Prerequisites
To run this script, you need:
* **7-Zip:** The script relies on the `7z.exe` executable to perform compression.
    * *Download:* [https://www.7-zip.org/](https://www.7-zip.org/)
* **Windows PowerShell:** Built into Windows.

### Configuration
**Important:** The script is pre-configured to look for 7-Zip at this specific path:
`C:\Program Files\7-Zip\7z.exe`

If your 7-Zip installation is in a custom location, you must edit the `$7zipPath` variable at the very top of the script to match your system.

### Usage
1.  Run the script.
2.  When prompted, paste the full path of the directory containing the folders you wish to zip.
3.  Review the list of folders found.
4.  Select **Yes** to begin compression.

---

## 2. PowerPoint to PDF Converter (`pptx_to_pdf.ps1`)

### Description
This script recursively searches through a source directory for PowerPoint files (`.ppt` and `.pptx`) and converts them to PDF format. It automatically recreates the source folder structure in the destination directory to keep files organized.

### Prerequisites
To run this script, you need:
* **Microsoft PowerPoint:** The script creates a `PowerPoint.Application` COM object, meaning actual PowerPoint software must be installed on the machine for the conversion to work.
    * *Reference:* [Microsoft Office PowerPoint](https://www.microsoft.com/en-us/microsoft-365/powerpoint)
* **Windows PowerShell:** Built into Windows.

### Usage
1.  Run the script.
2.  Enter the **Source Directory** path (where your PowerPoint files are currently located).
3.  Enter the **Destination Directory** path (where you want the PDFs saved).
    * *Note:* If the destination does not exist, the script will offer to create it for you.
4.  Select **Yes** to begin the batch conversion.

---

## Troubleshooting

* **Execution Policy Errors:** If you are unable to run the scripts due to security policies, you may need to run the following command in an Administrator PowerShell window:
    `Set-ExecutionPolicy RemoteSigned`
* **Path Errors:** If the `Zip_Folders_Individually.ps1` script fails immediately, double-check that the `$7zipPath` variable inside the script file points to the correct location of `7z.exe` on your computer.

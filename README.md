# FileFinder 1.1

**FileFinder** is a lightweight, GUI-based tool written in Python designed to help users quickly scan directories for specific file types and perform bulk actions like copying, deleting, or exporting file lists to Excel.

![Python](https://img.shields.io/badge/Python-3.x-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Version](https://img.shields.io/badge/Version-1.1-orange.svg)

## Inspiration

As an editor and filmmaker, I work with a huge number of files every day — images, audio, video, Excel sheets, Word documents, notes, and more. While Windows Explorer works for basic file selection, it quickly becomes a nightmare on large projects with nested folders and diverse file types, especially when clients send massive amounts of data.

This is what inspired me to create **FileFinder**. The tool allows me to quickly scan, filter, and organize files in any directory. It also makes project cleanup safe and efficient: I can remove downloaded templates or reference materials without losing track of what I actually need, because I always know exactly what resources I’ve gathered during my work.

## Features

*   **Deep Scanning:** Recursively scans folders for specific file extensions.
*   **Format Support:**
    *   **Images:** PNG, JPG, JPEG, WEBP
    *   **Videos:** MP4, MOV, AVI, MKV
    *   **Archives:** RAR, ZIP, 7Z, TAR
    *   **Audio:** MP3, WAV, FLAC, AAC
    *   **Documents:** DOC, DOCX, XLS, XLSX, PDF, PSD
*   **Custom Extensions:** Ability to input custom file extensions (e.g., `.py`, `.cpp`).
*   **Bulk Actions:**
    *   **Copy:** Organized copying (creates subfolders based on file type). Includes duplicate name handling (e.g., `image_1.png`).
    *   **Delete:** Permanently remove found files (with confirmation safety).
    *   **Excel Report:** Export the list of found files and paths to an `.xlsx` file.
*   **Responsive UI:** Multi-threaded architecture ensures the app doesn't freeze during large operations.

## Installation (Source Code)

If you want to run the script from the source:

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/e6pure/FileFinder.git
    cd FileFinder
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the application:**
    ```bash
    python FileFinder.py
    ```

## Download Executable (.exe)

If you don't have Python installed, you can download the latest standalone executable from the **[Releases](https://github.com/e6pure/FileFinder/releases)** page.

## Usage Guide

1.  **Select Source:** Click "Browse" to choose the folder you want to scan.
2.  **Select File Types:** Check the boxes for the file formats you are looking for, or add custom ones.
3.  **Run (Scan):** Click the button to search the directory. A summary will pop up showing what was found.
4.  **Select Action:**
    *   *Copy:* Choose a destination folder.
    *   *Delete:* **Warning:** This cannot be undone.
    *   *Excel:* Save a report.
5.  **Execute:** Click "Action" to start processing.

## Disclaimer

This software includes a **Delete** function. While safeguards (confirmation dialogs) are in place, the author is not responsible for any accidental data loss. Please verify your file selection before executing delete actions.

## Contributing

Contributions are welcome! Please submit issues or pull requests for bug fixes, improvements, or new features.

## Contact / Support

GitHub: https://github.com/e6pure

## 📝 Changelog

### v1.1
*   Changed UI language to English.
*   Implemented Multi-threading (GUI no longer freezes during heavy tasks).
*   Fixed variable conflicts regarding Excel selection.
*   Added automatic file renaming to prevent overwriting during Copy operations.
*   Improved error handling for individual files.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

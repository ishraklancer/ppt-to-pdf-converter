# PPT to PDF Batch Converter

A multithreaded tool to batch convert PowerPoint files (`.ppt`, `.pptx`) to PDF format. It uses Python and the Windows COM interface to reliably automate PowerPoint and export your presentations.

## Features
- **Multithreaded Processing**: Intelligently uses up to 50% of available CPU threads to process files rapidly while preventing your computer from freezing.
- **Safe COM Handling**: Includes global synchronization to prevent PowerPoint from rejecting concurrent requests.
- **Recursive Directory Search**: Optionally scans all nested subfolders for presentations.
- **Graphical Interface**: Features a clean UI with real-time logging and a visual progress bar.
- **CLI Support**: Fully automatable via the command line.

## Prerequisites
- Windows OS
- Microsoft PowerPoint must be installed on your machine.
- Python 3 installed.

## How to use

### The Easy Way (GUI)
Double-click `run.bat`. It will automatically install the required Python packages and launch the graphical interface.
1. Browse and select your input folder.
2. Check **Include nested folders** if you want to scan subdirectories.
3. Check **Delete source files** if you want to automatically clean up `.ppt`/`.pptx` files after a successful conversion.
4. Click **Start Conversion**. You'll see real-time progress and logs.

### Command Line Interface (CLI)
You can also use this tool directly from your terminal.

```cmd
# Basic usage
python converter.py -f "C:\path\to\your\folder" --cli

# Include all nested subfolders
python converter.py -f "C:\path\to\your\folder" --cli --recursive

# Delete source files after conversion
python converter.py -f "C:\path\to\your\folder" --cli --delete

# Combine options
python converter.py -f "C:\path\to\your\folder" --cli --recursive --delete
```

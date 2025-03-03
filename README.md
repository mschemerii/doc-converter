# Doc to Docx Converter

A Python solution to convert Microsoft Word .doc files to .docx format on both Windows and macOS.

## Requirements

- Python 3.6 or higher
- Microsoft Word installed on your system
- On Windows: `pywin32` package (`pip install pywin32`)
- Tkinter (included with most Python installations)

## Installation

1. Clone or download this repository
2. Ensure you have the required dependencies installed

## Usage

This converter can be used in two ways:

### 1. Command Line Interface

Convert a .doc file to .docx (saves in the same directory):

```bash
python convertDoc2Docx.py path/to/document.doc
```

Specify output path:

```bash
python convertDoc2Docx.py path/to/document.doc path/to/output.docx
```

### 2. Graphical User Interface

Launch the GUI application:

```bash
python doc_converter_gui.py
```

The GUI provides an easy-to-use interface with the following features:
- File browser to select input .doc files
- Optional output path selection
- Progress indication during conversion
- Status updates and error messages
- Simple one-click conversion

## How It Works

The converter uses different methods depending on your operating system:

- **Windows**: Uses the `win32com.client` library to automate Microsoft Word
- **macOS**: Uses AppleScript via the `subprocess` module to automate Microsoft Word

## Project Structure

- `convertDoc2Docx.py` - Core conversion functionality and command-line interface
- `doc_converter_gui.py` - Graphical user interface for the converter
- `README.md` - Documentation

## Limitations

- Requires Microsoft Word to be installed on your system
- Only supports Windows and macOS (not Linux)
- The input file must have a .doc extension

## Troubleshooting

### Windows

If you encounter an error about missing `win32com.client`, install the pywin32 package:

```bash
pip install pywin32
```

### macOS

Make sure Microsoft Word is installed and accessible. The script uses AppleScript to control Word, so Word must be able to open via AppleScript.

### GUI Issues

If the GUI doesn't appear or shows errors:
- Ensure Tkinter is properly installed with your Python installation
- Try running the command-line version to check if the core conversion works
- Check console output for any error messages

## Changelog

### March 3, 2025
- **Initial Release**: Created convertDoc2Docx.py script with support for Windows and macOS
- **Enhancement**: Improved error handling, added input validation and custom output paths
- **Feature**: Added separate TKinter GUI application (doc_converter_gui.py)
- **Documentation**: Created comprehensive README with usage instructions for both CLI and GUI
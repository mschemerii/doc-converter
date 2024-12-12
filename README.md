# Document Conversion and Processing Utility

A comprehensive Python-based utility for converting, modifying, and preparing document files across multiple processing stages. This tool is designed to handle .doc and .docx files, with specific focus on deployment and evidence documentation.

## Author

Michael Schemer

## Last Updated

December 9, 2024

## Features

- Convert .doc files to .docx format
- Modify table properties in documents
- Add empty rows after content rows in tables
- Create multiple document copies with specific modifications
- Platform-specific support for Windows
- Simple, intuitive GUI interface
- Automatic output window display
- Real-time conversion progress
- Copy-able conversion logs
- Cross-platform compatibility
- Robust error handling

## System Requirements

- Operating System: Windows - binaries are provided for Windows
- Python 3.11 or newer - All systems that can install Python 3.11+ are supported, including Linux
- Dependencies: `pip install -r requirements.txt`
- Tested on Python 3.9.13, 3.10.8, 3.11.4, 3.12.0, 3.13.0

## Installation and Usage

### Windows Installation

1. Download `doc_converter-windows.exe` from the [latest release](https://github.com/mschemer/doc-converter/releases)
2. Double-click the `.exe` file to launch the application
   - If Windows Defender or antivirus warns you, click "More info" and then "Run anyway"
3. The application will start with the Document Converter interface

### Python Environment Setup

1. Clone or download this repository
2. Open a terminal/command prompt
3. Navigate to the project directory
4. Install required packages:

   ```bash
   pip install -r requirements.txt
   ```

### Command Line Usage

#### Windows Command Line

Start -> cmd.exe:

```cmd
python doc_converter_gui.py
```

#### macOS Terminal

/Applications/Utilities/Terminal.app:

```bash
python3 doc_converter_gui.py
```

#### Linux Terminal

Open a terminal: (Oracle Linux default Gnome Terminal is named "gnome-terminal")

```bash
python3 doc_converter_gui.py
```

### GUI Usage

1. Click 'Browse' to select a .doc file
2. Click 'Convert' to start the conversion process
3. The output window will automatically appear showing conversion progress
4. Use the 'Copy Output' button to copy conversion logs if needed
5. Click 'Exit' when finished

[Rest of the README content...]

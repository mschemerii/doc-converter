# Document Conversion and Processing Utility

A comprehensive Python-based utility for converting, modifying, and preparing document files across multiple processing stages. This tool is designed to handle .doc and .docx files, with specific focus on deployment and evidence documentation.

## Features

- Convert .doc files to .docx format
- Modify table properties in documents
- Add empty rows after content rows in tables
- Create multiple document copies with specific modifications
- Platform-specific support for Windows, macOS, and Linux
- Automatic virtual environment management
- **New: Python Version Compatibility Check**
  - Automatically detects and recommends Python 3.12
  - Cross-platform download guidance for Python installation

## Prerequisites

- **Recommended Python Version: 3.12**
- Microsoft Word (for .doc to .docx conversion on Windows/macOS)
- Pandoc (for .doc to .docx conversion on Linux)
- Operating System: Windows, macOS, or Linux (Oracle Linux 8/9, Ubuntu LTS)

## Python Version Management

The application now includes an intelligent Python version check:
- Automatically detects your current Python version
- Recommends upgrading to Python 3.12
- Provides OS-specific download links
- Prevents application launch with incompatible Python versions

### Version Check Behavior
- If Python version is below 3.12, a dialog will appear
- Users can choose to download the recommended version
- Application will not start with unsupported Python versions

## Installation and Usage

### Windows
1. Download `doc_converter-windows.exe` from the latest release
2. Double-click the `.exe` file to launch the application
   - If Windows Defender or antivirus warns you, click "More info" and then "Run anyway"
3. The application will start with the Document Converter interface

### macOS
1. Download `Doc-Converter-Intel.app` (for Intel Macs) or `Doc-Converter-Silicon.app` (for Apple Silicon Macs) from the latest release
2. Running the App for the First Time:
   - Right-click (or Control-click) the `.app` and select "Open"
   - If macOS displays a security warning saying the app is from an unidentified developer:
     * Click "Open" to bypass the initial warning
     * You may need to go to System Preferences > Security & Privacy > General
     * Click "Open Anyway" for the Doc Converter application

#### Alternative Methods to Run the App
- **Terminal Method**: If you're comfortable with the command line, you can remove the quarantine attribute:
  ```bash
  # For Intel Macs
  xattr -d com.apple.quarantine /path/to/Doc-Converter-Intel.app
  
  # For Apple Silicon Macs
  xattr -d com.apple.quarantine /path/to/Doc-Converter-Silicon.app
  ```

#### Troubleshooting
- If the app doesn't open, ensure you have Python 3.12 installed
- Check System Preferences > Security & Privacy settings
- Verify that you have the necessary permissions to run applications

### Linux (Ubuntu/Debian)
1. Download the appropriate `.deb` file for your Ubuntu version
2. Install the package using the terminal:
```bash
sudo dpkg -i doc_converter-ubuntu-*.deb
```
3. Launch the application from your applications menu or via terminal:
```bash
doc_converter
```

### Linux (Oracle Linux/RPM-based)
1. Download the appropriate `.rpm` file for your Oracle Linux version
2. Install the package using the terminal:
```bash
sudo rpm -i doc_converter-ol*.rpm
```
3. Launch the application from your applications menu or via terminal:
```bash
doc_converter
```

## Running the Application

### Launching the GUI Executable

#### Standalone Executable
- Use the platform-specific executable from the latest release
- Double-click the application icon to launch

### Running Directly with Python

#### Prerequisites
- Python 3.12 installed
- Project dependencies installed

#### Setup Virtual Environment (Recommended)
```bash
# Create a virtual environment
python3 -m venv doc-converter-env

# Activate the virtual environment
# On Windows
doc-converter-env\Scripts\activate

# On macOS/Linux
source doc-converter-env/bin/activate

# Install dependencies
pip install -r requirements.txt
```

#### Launch GUI Directly
```bash
# Ensure you're in the project directory
python doc_converter_gui.py
```

#### Command-Line Usage
Some utility scripts can be run directly from the command line:

```bash
# Convert documents
python convert_docx.py input.doc output.docx

# Modify document properties
python modify_docx.py input.docx output.docx
```

### Development Mode

#### Running Tests
```bash
# Install test dependencies
pip install pytest

# Run tests
pytest tests/
```

#### Debugging
- Use `python -m pdb doc_converter_gui.py` for interactive debugging
- Set `logging.DEBUG` for verbose output in scripts

## Troubleshooting

### Permission Issues
- On Windows and macOS, right-click and choose "Run as administrator"
- On Linux, use `sudo` to run the application if needed

### Python Version Compatibility
- Ensure you have Python 3.12 installed
- The application will guide you through version-specific requirements
- Visit [python.org](https://www.python.org/downloads/) for the latest Python version

## Continuous Integration

- Automated builds for Windows, macOS, Linux
- Supports multiple Linux distributions
- Uses GitHub Actions for consistent deployment

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

[Specify your license here]

## Contact

On Oracle Slack @mschemer

## Output Files

The script generates three variants of the processed document:
1. **Stage-Evidence**: Deploy to Stage header, no rollback section
2. **StageDR-Evidence**: Deploy to StageDR header, no rollback section
3. **Rollback-Evidence**: Rollback header, no pre-deploy section

## Dependencies

Platform-independent:
- python-docx >= 0.8.11

Windows-specific:
- pywin32 >= 306

macOS-specific:
- pyobjc-framework-Cocoa >= 9.2
- pyobjc-framework-CoreServices >= 9.2
- pyobjc-framework-ScriptingBridge >= 9.2

Linux-specific:
- pandoc

## Usage Examples

1. Basic document processing via CLI:
```bash
python process_document.py "document.doc"
```

2. GUI Document Conversion:
```bash
python doc_converter_gui.py
```

### GUI Features
- Browse button to select .doc file
- Convert button to start conversion
- Real-time output window
- Exit button with safety checks

## Error Handling

The utility includes comprehensive error handling for:
- Missing dependencies
- File access issues
- Conversion failures
- Invalid document formats
- Platform-specific operations

## Limitations

- Requires Microsoft Word installation (Windows/macOS)
- Requires Pandoc installation (Linux)
- Platform-specific implementation (Windows/macOS/Linux)
- Assumes specific document structure for section removal
- Relies on exact text matching for section identification

## Known Linux Issues and Solutions

1. **Pandoc Conversion Issues**
   If you encounter issues with Pandoc conversion, ensure the latest version is installed:
   ```bash
   # Update Pandoc to the latest version
   sudo apt-get update
   sudo apt-get install pandoc
   ```

2. **Permission Issues**
   Ensure proper file permissions:
   ```bash
   chmod +x *.py
   ```

## Components

### 1. doc_to_docx_converter.py
Converts .doc files to .docx format:
- Uses Microsoft Word on Windows/macOS
- Uses Pandoc on Linux
- Maintains document formatting
- Handles file path conversion

### 2. modify_docx_tables.py
Modifies table properties in .docx files:
- Removes "Automatically resize to fit contents" setting
- Removes fixed column widths
- Sets table width to 100%
- Uses python-docx for XML manipulation

### 3. add_table_rows.py
Adds empty rows after content rows in tables:
- Preserves table formatting
- Copies row properties
- Maintains cell and paragraph formatting
- Handles complex table structures

### 4. rename_docx.py
Creates multiple document copies with modifications:
- Renames files (replaces "+-+" with "_")
- Removes "+" characters
- Adds custom suffixes
- Adds centered headers
- Selectively removes document sections:
  * Removes content from "Rollback" heading
  * Removes content from "Pre-Deploy Steps" to "Rollback"

### 5. process_document.py
Orchestrates the entire document processing workflow:
1. Manages Python virtual environment
2. Converts .doc to .docx
3. Modifies table properties
4. Adds empty rows to tables
5. Creates renamed copies with headers

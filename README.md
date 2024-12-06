# Document Conversion and Processing Utility

A comprehensive Python-based utility for converting, modifying, and preparing document files across multiple processing stages. This tool is designed to handle .doc and .docx files, with specific focus on deployment and evidence documentation.

## Features

- Convert .doc files to .docx format
- Modify table properties in documents
- Add empty rows after content rows in tables
- Create multiple document copies with specific modifications
- Platform-specific support for Windows and macOS
- Automatic virtual environment management

## Prerequisites

- Operating System: Windows or macOS - binaries are provided for Windows and macOS
- Python 3.11 or newer - All systems that can install Python 3.11+ are supported, including Linux

## Installation and Usage

### Windows
1. Download `doc_converter-windows.exe` from the [latest release](https://github.com/mschemer/doc-converter/releases)
2. Double-click the `.exe` file to launch the application
   - If Windows Defender or antivirus warns you, click "More info" and then "Run anyway"
3. The application will start with the Document Converter interface

### macOS
1. Download `Doc-Converter-Intel.app` (for Intel Macs) or `Doc-Converter-Silicon.app` (for Apple Silicon Macs) from the [latest release](https://github.com/mschemer/doc-converter/releases)
2. Running the App for the First Time:
   - Right-click (or Control-click) the `.app` and select "Open"
   - If macOS displays a security warning saying the app is from an unidentified developer:
     * Click "Open" to bypass the initial warning
     * You may need to go to System Preferences > Security & Privacy > General

### Running from Source

#### Prerequisites
- pip
- virtualenv (recommended)
- Python 3.11 or newer

#### Setup Steps
1. Clone the repository:
   ```bash
   git clone https://github.com/mschemer/doc-converter.git
   cd doc-converter
   ```
2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the application:
   ```bash
   python doc_converter_gui.py
   ```

## Troubleshooting

### Permission Issues
- On Windows, right-click and choose "Run as administrator"
- On macOS, ensure you have the necessary permissions to run applications

### General Troubleshooting
- Ensure all dependencies are installed
- Check the conversion output window for detailed error messages
- Ensure your .doc file is not corrupted or password-protected

## Continuous Integration

### Build Workflow
- Supports macOS (Intel and Apple Silicon) and Windows
- Uses GitHub Actions for automated builds
- Builds executable for each platform
- Automatic artifact and release generation

### Supported Platforms
- macOS 13 (Intel)
- macOS 14 (Apple Silicon)
- Windows 10/11 (64-bit)

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

[Specify your license here]

## Contact

[Provide contact information or support channels]

## Future Roadmap

- Potential Linux support
- Enhanced cross-platform compatibility
- Additional document processing features

## Latest Changes (March 2024)

### GUI Improvements
- Added clear instructions panel above the file selection area
- Automatic output window display during conversion
- Added copy button for conversion output
- Main window now appears on top at launch
- Improved window management and exit behavior
- Enhanced logging display in GUI

### System Changes
- Removed virtual environment dependencies
- Streamlined package management
- Improved cross-platform compatibility
- Enhanced error handling and user feedback
- Simplified installation process
- Removed Python version compatibility check

## System Requirements

- Python (with pip package installer)
- Operating System: macOS, Windows, or Linux

## Installation

1. Clone or download this repository
2. Open a terminal/command prompt
3. Navigate to the project directory
4. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

### From Command Line

#### macOS/Linux
```bash
python3 doc_converter_gui.py
```

#### Windows
```cmd
python doc_converter_gui.py
```

### From File Explorer/Finder

- **macOS**: Double-click `doc_converter_gui.py` (if Python is properly associated with .py files)
- **Windows**: Double-click `doc_converter_gui.py` (if Python is properly associated with .py files)
- **Linux**: Right-click `doc_converter_gui.py` and select "Run with Python" (if available in your desktop environment)

## Using the Application

1. Click 'Browse' to select a .doc file
2. Click 'Convert' to start the conversion process
3. The output window will automatically appear showing conversion progress
4. Use the 'Copy Output' button to copy conversion logs if needed
5. Click 'Exit' when finished

## Features

- Simple, intuitive GUI interface
- Automatic output window display
- Real-time conversion progress
- Copy-able conversion logs
- Cross-platform compatibility
- Robust error handling

## Troubleshooting

If you encounter any issues:

1. Verify Python is installed:
   ```bash
   python --version
   # or
   python3 --version
   ```

2. Verify all dependencies are installed:
   ```bash
   pip list
   ```

3. Check the conversion output window for detailed error messages

4. Ensure your .doc file is not corrupted or password-protected

## Support

For issues, questions, or contributions, please:
1. Check the existing issues in the repository
2. Create a new issue with detailed information about your problem
3. Include the conversion output log when reporting issues

## License

[Insert your license information here]

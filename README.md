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

## Prerequisites
- Operating System: Windows - binaries are provided for Windows
- Python 3.11 or newer - All systems that can install Python 3.11+ are supported, including Linux

## Installation and Usage

### Windows
1. Download `doc_converter-windows.exe` from the [latest release](https://github.com/mschemer/doc-converter/releases)
2. Double-click the `.exe` file to launch the application
   - If Windows Defender or antivirus warns you, click "More info" and then "Run anyway"
3. The application will start with the Document Converter interface

### Any Python3 Environment
#### Prerequisites
- pip
- Python 3.11 or newer

#### Installation
1. Clone or download this repository
2. Open a terminal/command prompt
3. Navigate to the project directory
4. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Running the Application

### From Command Line

#### Windows
Start -> cmd.exe:
```cmd
python doc_converter_gui.py
```
#### macOS
/Applications/Utilities/Terminal.app:
```bash
python3 doc_converter_gui.py
```
### Linux
Open a terminal: (Oracle Linux default Gnome Terminal is named "gnome-terminal")
```bash
python3 doc_converter_gui.py
```

## Using the Application

1. Click 'Browse' to select a .doc file
2. Click 'Convert' to start the conversion process
3. The output window will automatically appear showing conversion progress
4. Use the 'Copy Output' button to copy conversion logs if needed
5. Click 'Exit' when finished

## Troubleshooting

### Permission Issues
- On Windows, right-click and choose "Run as administrator"
- On macOS, ensure you have the necessary permissions to run applications

### General Troubleshooting
- Ensure all dependencies are installed
- Check the conversion output window for detailed error messages
- Ensure your .doc file is not corrupted or password-protected

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## Contact

Email: michael.schemer@oracle.com
Slack: @mschemer

## Latest Changes (December 9th 2024)

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



## Features

- Simple, intuitive GUI interface
- Automatic output window display
- Real-time conversion progress
- Copy-able conversion logs
- Cross-platform compatibility
- Robust error handling

## Self-Contained Binaries

The application is packaged as a self-contained binary. This means that it includes the Python interpreter and all necessary dependencies, allowing you to run it without needing to install Python separately.

### Building the Application

To build the application, use the following command:
```bash
python3 -m PyInstaller --onefile --add-data "assets/file_conversion_icon.icns;assets" doc_converter_gui.py
```

### Running the Application

Simply download the executable and run it directly. No additional installations are required.

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

# Document Conversion and Processing Utility

A comprehensive Python-based utility for converting, modifying, and preparing document files across multiple processing stages. This tool is designed to handle .doc and .docx files, with specific focus on deployment and evidence documentation.

## Features

- Convert .doc files to .docx format
- Modify table properties in documents
- Add empty rows after content rows in tables
- Create multiple document copies with specific modifications
- Platform-specific support for Windows and macOS
- Automatic virtual environment management
- **New: Python Version Compatibility Check**
  - Automatically detects and recommends Python 3.10
  - Cross-platform download guidance for Python installation

## Prerequisites

- **Recommended Python Version: 3.10**
- Microsoft Word (for .doc to .docx conversion)
- Operating System: Windows or macOS

## Python Version Management

The application now includes an intelligent Python version check:

- Automatically detects your current Python version
- Recommends upgrading to Python 3.10
- Provides OS-specific download links
- Prevents application launch with incompatible Python versions

### Version Check Behavior

- If Python version is below 3.10, a dialog will appear
- Users can choose to download the recommended version
- Application will not start with unsupported Python versions

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

- Python 3.10
- pip
- virtualenv (recommended)

#### Setup Steps

1. Clone the repository:

```bash
git clone https://github.com/mschemer/doc-converter.git
cd doc-converter
```

2. Create a virtual environment:

```bash
python3.10 -m venv venv
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

### Python Version Compatibility

- Ensure you have Python 3.10 installed
- The application will guide you through version-specific requirements
- Visit [python.org](https://www.python.org/downloads/) for the latest Python version

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

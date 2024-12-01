# Document Conversion and Processing Utility

A comprehensive Python-based utility for converting, modifying, and preparing document files across multiple processing stages. This tool is designed to handle .doc and .docx files, with specific focus on deployment and evidence documentation.

## Features

- Convert .doc files to .docx format
- Modify table properties in documents
- Add empty rows after content rows in tables
- Create multiple document copies with specific modifications
- Platform-specific support for Windows, macOS, and Linux
- Automatic virtual environment management

## Prerequisites

- Python 3.x
- Microsoft Word (for .doc to .docx conversion on Windows/macOS)
- Pandoc (for .doc to .docx conversion on Linux)
- Operating System: Windows, macOS, or Linux (Oracle Linux 7/8/9, Ubuntu LTS)

## Installation

### Windows and macOS

1. Create a new directory for the utility:
```bash
mkdir doc-converter
cd doc-converter
```

2. Download or copy these required files into the directory:
- `process_document.py`: Main orchestration script
- `doc_to_docx_converter.py`: Handles .doc to .docx conversion
- `modify_docx_tables.py`: Modifies table properties
- `add_table_rows.py`: Adds empty rows to tables
- `rename_docx.py`: Creates document copies with modifications
- `requirements.txt`: Lists all dependencies

3. Run the processing script:
```bash
python process_document.py <path-to-doc-file>
```

The script will automatically:
- Create a virtual environment (if not exists)
- Install required dependencies
- Process your document
- Deactivate the virtual environment when done

### Linux

#### Ubuntu/Debian
```bash
# Install system dependencies
sudo apt-get update
sudo apt-get install pandoc

# Set up the project
git clone <repository-url>
cd doc-converter
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

#### Oracle Linux 7/8/9
```bash
# Install system dependencies
sudo yum update
sudo yum install pandoc

# Set up the project
git clone <repository-url>
cd doc-converter
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Linux Conversion Notes

On Linux systems, the document conversion relies on Pandoc:
- Ensure Pandoc is installed before running the script
- Conversion may have limitations compared to Microsoft Word
- Some complex formatting might not be perfectly preserved

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

1. Basic document processing:
```bash
python process_document.py "document.doc"
```

2. Process multiple documents:
```bash
for file in *.doc; do
    python process_document.py "$file"
done
```

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

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## License

[Specify your license here]

## Support

For issues, questions, or contributions, please:
1. Check existing issues
2. Create a new issue with:
   - OS version
   - Python version
   - Document sample (if possible)
   - Error message/logs

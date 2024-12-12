#!/usr/bin/env python3
import os
import sys
import platform
import subprocess
from pathlib import Path
import tempfile
import time
import logging
import traceback

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler('doc_converter.log', mode='a'),
        logging.StreamHandler(sys.stdout)
    ]
)

def check_macos_requirements():
    """Check if Microsoft Word is installed on macOS"""
    try:
        result = subprocess.run(
            ['osascript', '-e', 'tell application "System Events" to exists application "Microsoft Word"'],
            capture_output=True,
            text=True,
            check=True
        )
        return result.stdout.strip().lower() == 'true'
    except Exception:
        return False

def convert_using_windows_com(doc_path, output_path):
    """Convert a .doc file to .docx using Microsoft Word COM interface on Windows"""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        logging.error("pywin32 not installed. Cannot convert using Windows COM.")
        logging.error("Please install with: pip install pywin32")
        raise

    # Initialize COM in the current thread
    pythoncom.CoInitialize()
    
    try:
        # Create Word application instance
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Run in background
        
        try:
            # Convert paths to absolute and normalize for Windows
            doc_path = os.path.abspath(doc_path).replace('/', '\\')
            output_path = os.path.abspath(output_path).replace('/', '\\')
            
            logging.info(f"Opening document: {doc_path}")
            
            # Open the document
            doc = word.Documents.Open(doc_path)
            
            # Constants for Word file formats
            wdFormatDocumentDefault = 16  # .docx
            
            logging.info(f"Saving as .docx: {output_path}")
            
            # Save as DOCX
            doc.SaveAs2(
                FileName=output_path,
                FileFormat=wdFormatDocumentDefault,
                AddToRecentFiles=False,
                ReadOnlyRecommended=False
            )
            
            # Close the document
            doc.Close(SaveChanges=False)
            
            logging.info(f"Successfully converted {doc_path} to {output_path}")
            return True
            
        except Exception as e:
            logging.error(f"Error during Word COM conversion: {str(e)}")
            logging.error(traceback.format_exc())
            raise
            
        finally:
            # Always quit Word and release COM objects
            try:
                word.Quit()
                del word
            except:
                pass
            
    finally:
        # Clean up COM
        pythoncom.CoUninitialize()

def check_windows_requirements():
    """Check if Microsoft Word is available on Windows"""
    try:
        import win32com.client
        
        # Try to create Word application instance
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        return True
    except Exception as e:
        logging.error(f"Microsoft Word not available on Windows: {str(e)}")
        return False

def convert_using_pandoc(doc_path, output_path):
    """Convert a .doc file to .docx using Pandoc on Linux"""
    try:
        import pypandoc
        
        # Ensure output is .docx
        if not output_path.lower().endswith('.docx'):
            output_path = os.path.splitext(output_path)[0] + '.docx'
        
        # Convert using Pandoc
        pypandoc.convert_file(
            doc_path, 
            'docx', 
            outputfile=output_path
        )
        
        logging.info(f"Successfully converted {doc_path} to {output_path} using Pandoc")
        
    except ImportError:
        logging.error("Pandoc or pypandoc not installed. Cannot convert using Pandoc.")
        raise
    except Exception as e:
        logging.error(f"Error converting document using Pandoc: {e}")
        logging.error(traceback.format_exc())
        raise

def convert_using_macos_word(doc_path, output_path):
    """Convert a .doc file to .docx using Microsoft Word via AppleScript on macOS"""
    try:
        # Create AppleScript command
        applescript = f'''
            tell application "Microsoft Word"
                set wordDoc to open "{doc_path}"
                save as wordDoc file name "{output_path}" file format format document
                close wordDoc saving no
                quit
            end tell
        '''
        
        # Run AppleScript command
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True,
            text=True,
            check=True
        )
        
        logging.info(f"Successfully converted {doc_path} to {output_path}")
        return True
        
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running AppleScript: {e.stderr}")
        raise
    except Exception as e:
        logging.error(f"Error converting document using macOS Word: {e}")
        logging.error(traceback.format_exc())
        raise

def convert_doc_to_docx(doc_path):
    """Convert a .doc file to .docx using available methods"""
    # Validate input file
    if not os.path.exists(doc_path):
        logging.error(f"Input file not found: {doc_path}")
        raise FileNotFoundError(f"Input file not found: {doc_path}")
    
    # Get absolute paths
    doc_path = os.path.abspath(doc_path)
    output_path = os.path.splitext(doc_path)[0] + '.docx'
    
    # Determine conversion method based on platform
    os_name = platform.system().lower()
    
    try:
        if os_name == 'windows':
            if not check_windows_requirements():
                raise RuntimeError("Microsoft Word not available on Windows")
            
            logging.info("Using Microsoft Word COM interface for Windows conversion")
            convert_using_windows_com(doc_path, output_path)
        elif os_name == 'darwin':  # macOS
            if not check_macos_requirements():
                raise RuntimeError("Microsoft Word not found on macOS")
            
            logging.info("Using Microsoft Word via AppleScript for macOS conversion")
            convert_using_macos_word(doc_path, output_path)
        elif os_name == 'linux':
            convert_using_pandoc(doc_path, output_path)
        else:
            logging.error(f"Unsupported operating system: {os_name}")
            raise OSError(f"Conversion not supported on {os_name}")
        
        # Verify the output file was created
        if not os.path.exists(output_path):
            raise RuntimeError(f"Conversion failed: Output file not created at {output_path}")
        
        return output_path
    
    except Exception as e:
        logging.critical(f"Conversion failed for {doc_path}: {e}")
        logging.critical(traceback.format_exc())
        raise

def main():
    """Command-line interface for document conversion"""
    if len(sys.argv) < 2:
        logging.error("Usage: python doc_to_docx_converter.py <input_file.doc>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    try:
        output_file = convert_doc_to_docx(input_file)
        print(f"Successfully converted {input_file} to {output_file}")
    except Exception as e:
        print(f"Conversion failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()

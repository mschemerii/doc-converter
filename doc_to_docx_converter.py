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

def convert_using_windows_com(doc_path, output_path):
    """Convert a .doc file to .docx using Microsoft Word COM interface on Windows"""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        logging.error("pywin32 not installed. Cannot convert using Windows COM.")
        raise

    # Initialize COM in the current thread
    pythoncom.CoInitialize()
    
    try:
        # Create Word application instance
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        try:
            # Open the document
            doc = word.Documents.Open(doc_path)
            
            # Save as DOCX
            doc.SaveAs2(output_path, FileFormat=16)  # wdFormatDocumentDefault = 16
            
            # Close the document
            doc.Close()
            
            logging.info(f"Successfully converted {doc_path} to {output_path}")
            
        except Exception as e:
            logging.error(f"Error converting document using Windows COM: {e}")
            logging.error(traceback.format_exc())
            raise
            
        finally:
            # Quit Word application
            word.Quit()
            
    finally:
        # Clean up COM
        pythoncom.CoUninitialize()

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

def convert_doc_to_docx(doc_path):
    """
    Convert a .doc file to .docx using available methods
    Supports Windows, macOS, and Linux
    """
    # Validate input file
    if not os.path.exists(doc_path):
        logging.error(f"Input file not found: {doc_path}")
        raise FileNotFoundError(f"Input file not found: {doc_path}")
    
    # Determine output path
    output_path = os.path.splitext(doc_path)[0] + '.docx'
    
    # Determine conversion method based on platform
    os_name = platform.system().lower()
    
    try:
        if os_name == 'windows':
            convert_using_windows_com(doc_path, output_path)
        elif os_name == 'darwin':  # macOS
            # Implement macOS-specific conversion (e.g., AppleScript)
            logging.warning("macOS conversion not fully implemented")
            raise NotImplementedError("macOS conversion not yet supported")
        elif os_name == 'linux':
            convert_using_pandoc(doc_path, output_path)
        else:
            logging.error(f"Unsupported operating system: {os_name}")
            raise OSError(f"Conversion not supported on {os_name}")
        
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

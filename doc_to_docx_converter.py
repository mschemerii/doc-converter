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

def convert_using_windows_com(doc_path: str, output_path: str) -> None:
    """Convert a .doc file to .docx using Microsoft Word COM interface on Windows"""
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
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

def convert_using_pandoc(doc_path: str, output_path: str) -> None:
    """Convert a .doc file to .docx using Pandoc on Linux"""
    try:
        import pypandoc  # type: ignore
        
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
        logging.info(f"Final output path: {output_path}")
        
    except ImportError:
        logging.error("Pandoc or pypandoc not installed. Cannot convert using Pandoc.")
        raise
    except Exception as e:
        logging.error(f"Error converting document using Pandoc: {e}")
        logging.error(traceback.format_exc())
        raise

#!/usr/bin/env python3
import os
import sys
import platform
import subprocess
from pathlib import Path
import tempfile
import time

def convert_using_windows_com(doc_path, output_path):
    """Convert a .doc file to .docx using Microsoft Word COM interface on Windows"""
    import win32com.client
    import pythoncom
    
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
            
        finally:
            # Quit Word application
            word.Quit()
            
    finally:
        # Clean up COM
        pythoncom.CoUninitialize()

def write_applescript(doc_path, output_path):
    """Create the AppleScript file for Word conversion"""
    script = f'''
    tell application "Microsoft Word"
        activate
        set input_path to "{doc_path}"
        set output_path to "{output_path}"
        
        open input_path
        
        tell active document
            save as file name output_path file format format document
            close saving no
        end tell
    end tell
    '''
    
    # Create a temporary file for the AppleScript
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.scpt', mode='w')
    temp_file.write(script)
    temp_file.close()
    return temp_file.name

def convert_doc_to_docx(doc_path):
    """
    Convert a .doc file to .docx using Microsoft Word
    Supports both Windows and macOS
    """
    # Convert path to absolute path
    doc_path = os.path.abspath(doc_path)
    
    # Check if file exists
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"File not found: {doc_path}")
    
    # Check if it's a .doc file
    if not doc_path.lower().endswith('.doc'):
        raise ValueError("Input file must be a .doc file")
    
    # Create output path
    output_path = str(Path(doc_path).with_suffix('.docx'))
    
    system = platform.system()
    
    try:
        if system == 'Windows':
            convert_using_windows_com(doc_path, output_path)
        elif system == 'Darwin':  # macOS
            # Create AppleScript file
            script_path = write_applescript(doc_path, output_path)
            
            try:
                result = subprocess.run(
                    ['osascript', script_path],
                    capture_output=True,
                    text=True,
                    check=True
                )
            finally:
                # Clean up the temporary script file
                os.unlink(script_path)
        else:
            raise Exception(f"Unsupported operating system: {system}")
        
        # Wait briefly and check if the output file was created
        time.sleep(1)
        if not os.path.exists(output_path):
            raise Exception("Conversion failed: Output file was not created")
        
        return output_path
            
    except subprocess.CalledProcessError as e:
        raise Exception(f"Word conversion failed: {e.stderr}")
    except Exception as e:
        raise Exception(f"Conversion failed: {str(e)}")

def main():
    if len(sys.argv) != 2:
        print("Usage: python doc_to_docx_converter.py <path_to_doc_file>")
        sys.exit(1)
    
    try:
        input_file = sys.argv[1]
        output_file = convert_doc_to_docx(input_file)
        print(f"Successfully converted file to: {output_file}")
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()

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
        # Check the known Word location first
        word_path = '/Applications/Microsoft Word.app'
        if os.path.exists(word_path):
            logging.info(f"Found Word at: {word_path}")
            
            # Verify we can communicate with Word
            check_script = '''
                try
                    tell application "Microsoft Word"
                        if not running then
                            launch
                            delay 1
                        end if
                        quit
                        return true
                    end tell
                on error errMsg
                    log errMsg
                    return false
                end try
            '''
            
            result = subprocess.run(
                ['osascript', '-e', check_script],
                capture_output=True,
                text=True,
                timeout=10
            )
            
            if result.stdout.strip().lower() == 'true':
                logging.info("Successfully verified Word installation")
                return True
            else:
                logging.warning(f"Word found at {word_path} but communication failed")
        
        logging.error("Microsoft Word not found or not accessible")
        return False
        
    except Exception as e:
        logging.error(f"Error checking Word availability: {str(e)}")
        logging.error(traceback.format_exc())
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
        # Ensure paths are absolute and properly escaped for AppleScript
        doc_path = os.path.abspath(doc_path).replace('"', '\\"')
        output_path = os.path.abspath(output_path).replace('"', '\\"')
        
        logging.info(f"Attempting to convert: {doc_path} to {output_path}")
        
        # Create AppleScript command with correct Word syntax
        applescript = f'''
            try
                tell application "Microsoft Word"
                    set isRunning to running
                    if not isRunning then
                        launch
                        delay 2
                    end if
                    
                    -- Open the document
                    set docPath to POSIX file "{doc_path}"
                    set doc to open docPath
                    
                    -- Wait for document to load
                    delay 2
                    
                    -- Save as docx
                    set outputPath to POSIX file "{output_path}"
                    save as active document file name outputPath file format document format
                    
                    -- Close document
                    close active document saving no
                    
                    -- Quit if we launched it
                    if not isRunning then
                        quit
                    end if
                    
                    return "success"
                end tell
            on error errMsg
                log errMsg
                error errMsg
            end try
        '''
        
        # Run AppleScript command
        logging.info("Executing AppleScript...")
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True,
            text=True,
            check=True
        )
        
        if result.stderr:
            logging.warning(f"AppleScript warnings: {result.stderr}")
        
        logging.info(f"AppleScript output: {result.stdout}")
        
        # Verify the file was created
        if not os.path.exists(output_path):
            raise RuntimeError(f"Output file not created at {output_path}")
        
        logging.info(f"Successfully converted {doc_path} to {output_path}")
        return True
        
    except subprocess.CalledProcessError as e:
        logging.error(f"AppleScript error: {e.stderr}")
        logging.error(traceback.format_exc())
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

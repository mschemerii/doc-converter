#!/usr/bin/env python3
"""
Doc to Docx Converter

This script converts Microsoft Word .doc files to .docx format on both Windows and macOS.
It requires Microsoft Word to be installed on the system.

Usage:
    python convertDoc2Docx.py <input_file> [<output_file>]

Arguments:
    input_file  - Path to the .doc file to convert
    output_file - (Optional) Path where the .docx file should be saved
                  If not provided, the file will be saved in the same directory
                  with the same name but .docx extension
"""

import os  # For file path operations and environment variables
import sys  # For command-line arguments and exit codes
import platform  # For detecting the operating system
import subprocess  # For running external commands (AppleScript on macOS)
from pathlib import Path  # For object-oriented file path manipulation


def doc_to_docx_windows(input_path, output_path=None):
    """
    Convert .doc file to .docx file using Microsoft Word on Windows.
    
    Args:
        input_path (str): Path to the .doc file.
        output_path (str, optional): Path where the .docx file should be saved.
            If not provided, the file will be saved in the same directory with
            the same name but .docx extension.
        
    Returns:
        str: Path to the converted .docx file, or None if conversion failed.
    """
    try:
        # Import win32com.client only on Windows
        import win32com.client # type: ignore
        import pythoncom # type: ignore
        
        # Check if Microsoft Word is installed
        try:
            # Try to create a Word application instance just to check if Word is installed
            word_check = win32com.client.Dispatch("Word.Application")
            # If we get here, Word is installed, so close the application
            word_check.Quit()
            print("Microsoft Word is installed. Proceeding with conversion...")
        except pythoncom.com_error:
            print("Error: Microsoft Word is not installed on this system.")
            return None
        except Exception as e:
            print(f"Warning: Could not verify if Microsoft Word is installed: {e}")
        
        # Get absolute paths
        input_abs_path = os.path.abspath(input_path)
        
        if output_path:
            output_abs_path = os.path.abspath(output_path)
        else:
            # Get the base filename without extension and add .docx
            base_path = os.path.splitext(input_abs_path)[0]
            output_abs_path = f"{base_path}.docx"
        
        print(f"Converting {input_abs_path} to {output_abs_path}...")
        
        # Create a new instance of Microsoft Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Run in background
        
        try:
            # Open the document
            doc = word.Documents.Open(input_abs_path)
            
            # Save the document as .docx (16 is the format code for .docx)
            doc.SaveAs(output_abs_path, FileFormat=16)
            
            # Close the document
            doc.Close()
            
            print("Conversion completed successfully.")
            return output_abs_path
        except Exception as e:
            print(f"Error during Word automation: {e}")
            return None
        finally:
            # Ensure Word is closed even if an error occurs
            word.Quit()
    except ImportError:
        print("Error: win32com.client module is required for Windows conversion.")
        print("Install it using: pip install pywin32")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None


def doc_to_docx_macos(input_path, output_path=None):
    """
    Convert .doc file to .docx file using Microsoft Word on macOS.
    
    Args:
        input_path (str): Path to the .doc file.
        output_path (str, optional): Path where the .docx file should be saved.
            If not provided, the file will be saved in the same directory with
            the same name but .docx extension.
        
    Returns:
        str: Path to the converted .docx file, or None if conversion failed.
    """
    try:
        # Get absolute paths
        input_abs_path = os.path.abspath(input_path)
        
        if output_path:
            output_abs_path = os.path.abspath(output_path)
        else:
            # Get the base filename without extension and add .docx
            base_path = os.path.splitext(input_abs_path)[0]
            output_abs_path = f"{base_path}.docx"
        
        print(f"Converting {input_abs_path} to {output_abs_path}...")
        
        # Check if Microsoft Word is installed
        try:
            result = subprocess.run(
                ['osascript', '-e', 'application "Microsoft Word" exists'],
                capture_output=True,
                text=True,
                check=True
            )
            if "false" in result.stdout.lower():
                print("Error: Microsoft Word is not installed on this system.")
                return None
        except subprocess.SubprocessError:
            print("Warning: Could not verify if Microsoft Word is installed.")
        
        # Use AppleScript to open the document in Microsoft Word and save it as .docx
        applescript_code = f'''
        tell application "Microsoft Word"
            open POSIX file "{input_abs_path}"
            set doc to active document
            save as doc in POSIX file "{output_abs_path}" file format format docx
            close doc
            quit
        end tell
        '''
        
        # Run the AppleScript code
        subprocess.run(['osascript', '-e', applescript_code], check=True)
        
        # Verify the file was created
        if os.path.exists(output_abs_path):
            print("Conversion completed successfully.")
            return output_abs_path
        else:
            print("Error: Conversion failed. Output file was not created.")
            return None
    except subprocess.SubprocessError as e:
        print(f"Error during AppleScript execution: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None


def doc_to_docx(input_path, output_path=None):
    """
    Convert .doc file to .docx file based on the current operating system.
    
    Args:
        input_path (str): Path to the .doc file.
        output_path (str, optional): Path where the .docx file should be saved.
            If not provided, the file will be saved in the same directory with
            the same name but .docx extension.
        
    Returns:
        str: Path to the converted .docx file, or None if conversion failed.
    """
    # Validate input file
    if not os.path.exists(input_path):
        print(f"Error: Input file '{input_path}' does not exist.")
        return None
    
    # Check if the file is a .doc file
    if not input_path.lower().endswith('.doc'):
        print(f"Error: Input file '{input_path}' is not a .doc file.")
        return None
    
    # Convert based on operating system
    system = platform.system()
    if system == 'Windows':
        return doc_to_docx_windows(input_path, output_path)
    elif system == 'Darwin':  # macOS
        return doc_to_docx_macos(input_path, output_path)
    else:
        print(f"Error: Unsupported operating system '{system}'.")
        print("This script only supports Windows and macOS.")
        return None


def main():
    """
    Main function to handle command line arguments and execute the conversion.
    """
    # Check command line arguments
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print(f"Usage: python {os.path.basename(__file__)} <input_file> [<output_file>]")
        sys.exit(1)
    
    # Get input and output paths
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) == 3 else None
    
    # Convert the file
    result = doc_to_docx(input_path, output_path)
    
    # Check the result
    if result:
        print(f"Converted file saved at: {result}")
        sys.exit(0)
    else:
        print("Conversion failed.")
        sys.exit(1)


if __name__ == "__main__":
    main()
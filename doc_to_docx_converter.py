import os
import sys
import platform
import subprocess
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)


def convert_using_windows_com(doc_path, output_path):
    """Convert a .doc file to .docx using Microsoft Word COM interface on Windows"""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        logging.error("pywin32 not installed. Cannot convert using Windows COM.")
        logging.error("Please install with: pip install pywin32")
        return False

    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(output_path, FileFormat=16)  # 16 is the format for .docx
        doc.Close()
        word.Quit()
        logging.info(f"Successfully converted {doc_path} to {output_path}")
        return True
    except Exception as e:
        logging.error(f"Failed to convert {doc_path} using Windows COM: {e}")
        return False


def convert_using_macos_word(doc_path, output_path):
    """Convert a .doc file to .docx using Microsoft Word via AppleScript on macOS"""
    applescript = f"""
    tell application "Microsoft Word"
        set isRunning to running
        if not isRunning then
            launch
            delay 2
        end if
        
        open POSIX file "{doc_path}"
        
        delay 5
        
        save as active document file name POSIX file "{output_path}"
        
        close active document saving no
        
        if not isRunning then
            quit saving no
        end if
    end tell
    """

    try:
        subprocess.run(['osascript', '-e', applescript], check=True)
        logging.info(f"Successfully converted {doc_path} to {output_path}")
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to convert {doc_path} using macOS Word: {e}")
        return False


def convert_using_pandoc(doc_path, output_path):
    """Convert a .doc file to .docx using Pandoc on Linux"""
    try:
        subprocess.run(['pandoc', doc_path, '-o', output_path], check=True)
        logging.info(f"Successfully converted {doc_path} to {output_path}")
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to convert {doc_path} using Pandoc: {e}")
        return False


def convert_doc_to_docx(doc_path, output_path):
    """Convert a .doc file to .docx based on the operating system."""
    if platform.system() == "Windows":
        return convert_using_windows_com(doc_path, output_path)
    elif platform.system() == "Darwin":  # macOS
        return convert_using_macos_word(doc_path, output_path)
    elif platform.system() == "Linux":
        return convert_using_pandoc(doc_path, output_path)
    else:
        logging.error("Unsupported operating system.")
        return False


def main():
    if len(sys.argv) != 3:
        logging.error("Usage: python doc_converter.py <input_file> <output_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    convert_doc_to_docx(input_file, output_file)


if __name__ == "__main__":
    main()
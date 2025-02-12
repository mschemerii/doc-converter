import os
import sys
import platform
import subprocess
import win32com.client

def doc_to_docx(file_path):
    """
    Convert .doc file to .docx file.
    
    Args:
        file_path (str): Path to the .doc file.
        
    Returns:
        str: Path to the converted .docx file.
    """
    # Check if the operating system is Windows
    if platform.system() == 'Windows':
        try:
            # Create a new instance of Microsoft Word
            word = win32com.client.Dispatch("Word.Application")
            
            # Open the document
            doc = word.Documents.Open(file_path)
            
            # Get the base filename without extension
            base_filename = os.path.splitext(file_path)[0]
            
            # Save the document as .docx
            doc.SaveAs(base_filename + ".docx", FileFormat=16)  # 16 is the format code for .docx
            
            # Close the document and quit Word
            doc.Close()
            word.Quit()
            
            return base_filename + ".docx"
        except Exception as e:
            print(f"An error occurred while converting the file: {e}")
            return None
    elif platform.system() == 'Darwin':  # macOS
        try:
            # Get the base filename without extension
            base_filename = os.path.splitext(file_path)[0]
            
            # Use AppleScript to open the document in Microsoft Word and save it as .docx
            applescript_code = f'''
            tell application "Microsoft Word"
                open POSIX file "{file_path}"
                set doc to active document
                save as doc in POSIX file "{base_filename}.docx" file format docx
                close doc
            end tell
            '''
            
            # Run the AppleScript code
            subprocess.call(['osascript', '-e', applescript_code])
            
            return base_filename + ".docx"
        except Exception as e:
            print(f"An error occurred while converting the file: {e}")
            return None
    else:
        print("Unsupported operating system.")
        return None

# Example usage
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python doc_to_docx.py <file_path>")
    else:
        file_path = sys.argv[1]
        converted_file_path = doc_to_docx(file_path)
        if converted_file_path:
            print(f"Converted file saved at {converted_file_path}")
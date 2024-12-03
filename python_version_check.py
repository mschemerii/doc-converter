import sys
import tkinter as tk
from tkinter import messagebox
import webbrowser
import platform

def check_python_version():
    # Minimum supported Python version
    MIN_MAJOR = 3
    MIN_MINOR = 10
    
    # Current Python version
    current_major = sys.version_info.major
    current_minor = sys.version_info.minor
    
    # Check if Python version is less than minimum supported
    if (current_major < MIN_MAJOR or 
        (current_major == MIN_MAJOR and current_minor < MIN_MINOR)):
        
        # Create a root window (hidden)
        root = tk.Tk()
        root.withdraw()
        
        # Construct version message
        current_version = f"{current_major}.{current_minor}"
        recommended_version = f"{MIN_MAJOR}.{MIN_MINOR}+"
        
        # Determine download URL based on OS
        os_name = platform.system().lower()
        if os_name == 'darwin':
            download_url = "https://www.python.org/downloads/macos/"
        elif os_name == 'windows':
            download_url = "https://www.python.org/downloads/windows/"
        else:
            download_url = "https://www.python.org/downloads/"
        
        # Construct message
        message = (
            f"Python {recommended_version} is recommended.\n"
            f"Current version: {current_version}\n\n"
            "Would you like to download the latest Python version?"
        )
        
        # Show message box
        response = messagebox.askyesno(
            "Python Version Warning", 
            message, 
            icon='warning'
        )
        
        # Open download page if user agrees
        if response:
            webbrowser.open(download_url)
        
        return False
    
    return True

# Optional: Run check when imported
if __name__ == '__main__':
    if not check_python_version():
        sys.exit(1)  # Exit with error code if version check fails

import sys
import tkinter as tk
from tkinter import messagebox
import webbrowser
import platform

def check_python_version():
    # Recommended Python version
    RECOMMENDED_MAJOR = 3
    RECOMMENDED_MINOR = 12
    
    # Current Python version
    current_major = sys.version_info.major
    current_minor = sys.version_info.minor
    
    # Check if Python version is less than recommended
    if (current_major < RECOMMENDED_MAJOR or 
        (current_major == RECOMMENDED_MAJOR and current_minor < RECOMMENDED_MINOR)):
        
        # Create a root window (hidden)
        root = tk.Tk()
        root.withdraw()
        
        # Construct version message
        current_version = f"{current_major}.{current_minor}"
        recommended_version = f"{RECOMMENDED_MAJOR}.{RECOMMENDED_MINOR}"
        
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
        
        # Close hidden root window
        root.destroy()
        
        return False
    
    return True

# Optional: Run check when imported
if __name__ == '__main__':
    check_python_version()
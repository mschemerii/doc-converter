#!/usr/bin/env python3
# Executable GUI for Doc Converter
# Last updated: 2024-03-14

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import logging
import threading
import traceback

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    handlers=[
        logging.StreamHandler()
    ]
)

# Import Python version check
from python_version_check import check_python_version

# Run version check immediately
if not check_python_version():
    logging.error("Python version check failed. Exiting.")
    sys.exit(1)

# Import the existing conversion script
from doc_to_docx_converter import convert_doc_to_docx

class RedirectText:
    """Redirect print statements to a tkinter Text widget"""
    def __init__(self, text_widget):
        self.output = text_widget
    
    def write(self, string):
        try:
            self.output.insert(tk.END, string)
            self.output.see(tk.END)
        except Exception as e:
            logging.error(f"Error writing to text widget: {e}")
    
    def flush(self):
        pass

class DocConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("Doc Converter")
        master.geometry("500x400")
        
        # Configure grid weights for centering
        master.grid_rowconfigure(0, weight=0)  # Instructions
        master.grid_rowconfigure(1, weight=0)  # Status text
        master.grid_rowconfigure(2, weight=0)  # Progress bar
        master.grid_rowconfigure(3, weight=1)  # Main frame
        master.grid_rowconfigure(6, weight=1)
        master.grid_columnconfigure(0, weight=1)
        
        # Instructions label
        instructions = (
            "1. Click on Browse to select your file.\n"
            "2. Click on Convert to start the conversion process."
        )
        self.instructions_label = tk.Label(
            master,
            text=instructions,
            justify=tk.LEFT,
            anchor="w",
            padx=20,
            pady=10
        )
        self.instructions_label.grid(row=0, column=0, sticky="w")
        
        # Status text
        self.status_var = tk.StringVar(value="Ready to convert...")
        self.status_label = tk.Label(
            master,
            textvariable=self.status_var,
            justify=tk.LEFT,
            anchor="w",
            padx=20,
            pady=5
        )
        self.status_label.grid(row=1, column=0, sticky="w")
        
        # Progress bar
        self.progress = ttk.Progressbar(
            master,
            orient="horizontal",
            length=460,
            mode="determinate"
        )
        self.progress.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        
        # Main container frame
        self.main_frame = tk.Frame(master)
        self.main_frame.grid(row=3, column=0, sticky="nsew")
        
        # Configure main frame grid
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # File selection frame
        self.file_frame = tk.Frame(self.main_frame)
        self.file_frame.grid(row=0, column=0, pady=10, padx=20, sticky="ew")
        self.file_frame.grid_columnconfigure(0, weight=1)
        
        # File entry and browse button in file frame
        self.file_path = tk.StringVar()
        self.file_entry = tk.Entry(self.file_frame, textvariable=self.file_path)
        self.file_entry.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        
        self.browse_button = tk.Button(self.file_frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=1)
        
        # Convert button
        self.convert_button = tk.Button(self.main_frame, text="Convert", command=self.start_conversion, 
                                    state=tk.NORMAL, width=15)
        self.convert_button.grid(row=1, column=0, pady=10)
        
        # Exit button (always enabled)
        self.exit_button = tk.Button(self.main_frame, text="Exit", command=self.exit_app, 
                                   state=tk.NORMAL, width=15)
        self.exit_button.grid(row=2, column=0, pady=10)
        
        # Store last output
        self.last_output = []
    
    def browse_file(self):
        """Open file browser to select .doc file"""
        filename = filedialog.askopenfilename(
            title="Select .doc file",
            filetypes=[("Word Document", "*.doc")]
        )
        if filename:
            self.file_path.set(filename)
    
    def start_conversion(self):
        """Start conversion process in a separate thread"""
        input_file = self.file_path.get()
        
        # Validate input
        if not input_file:
            messagebox.showerror("Error", "Please select a .doc file to convert")
            return
        
        if not input_file.lower().endswith('.doc'):
            messagebox.showerror("Error", "Please select a .doc file")
            return
        
        # Clear previous output
        self.last_output = []
        
        # Disable convert button during conversion
        self.convert_button.config(state=tk.DISABLED)
        
        # Start conversion in a separate thread
        conversion_thread = threading.Thread(
            target=self.run_conversion, 
            args=(input_file,), 
            daemon=True
        )
        conversion_thread.start()
    
    def run_conversion(self, input_file):
        """Perform the actual conversion"""
        try:
            # Reset progress bar
            self.progress["value"] = 0
            self.status_var.set("Starting conversion process...")
            self.master.update_idletasks()
            
            # Import the full document processing function
            from process_document import DocumentProcessor
            
            # Create an instance of DocumentProcessor
            processor = DocumentProcessor()
            
            # Update progress for initialization
            self.progress["value"] = 20
            self.status_var.set("Converting .doc to .docx...")
            self.master.update_idletasks()
            
            # Call the process_document method with the input file
            success = processor.process_document(input_file)
            
            # Update progress for main conversion
            self.progress["value"] = 60
            self.status_var.set("Processing document structure...")
            self.master.update_idletasks()
            
            if success:
                # Update progress for completion
                self.progress["value"] = 100
                self.status_var.set("Conversion completed successfully!")
                self.master.update_idletasks()
                
                # Show success message
                message = self.format_success_message(input_file)
                self.master.after(0, lambda: messagebox.showinfo("Success", message))
            else:
                self.progress["value"] = 0
                self.status_var.set("Conversion failed")
                self.master.update_idletasks()
                # Show error message
                self.master.after(0, lambda: messagebox.showerror("Error", "Document processing failed"))
        except Exception as e:
            self.progress["value"] = 0
            self.status_var.set("Error occurred during conversion")
            self.master.update_idletasks()
            error_message = f"Error during conversion: {str(e)}"
            self.master.after(0, lambda: messagebox.showerror("Error", error_message))
        finally:
            self.convert_button.config(state=tk.NORMAL)
            # Reset progress after a delay if there was an error
            if self.progress["value"] != 100:
                self.master.after(2000, self.reset_progress)

    def format_success_message(self, input_file):
        """Format the success message with file details"""
        directory = os.path.dirname(input_file) or '.'
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        
        message = (
            f"Document Processing Complete!\n\n"
            f"Original File: {input_file}\n\n"
            f"Processed Files Location: {directory}\n\n"
            f"Files Created:\n"
            f"- {base_name}.docx\n"
            f"- {base_name}_with_headers.docx\n"
            f"- {base_name}_final.docx"
        )
        return message
    
    def reset_progress(self):
        """Reset progress bar and status text"""
        self.progress["value"] = 0
        self.status_var.set("Ready to convert...")
        self.master.update_idletasks()
    
    def exit_app(self):
        """Handle application exit"""
        # Just close any output window and quit
        self.master.quit()

def main():
    root = tk.Tk()
    app = DocConverterApp(root)
    
    # Center the window on the screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'+{x}+{y}')
    
    root.attributes('-topmost', True)  # Ensure the main application window is the topmost window
    root.attributes('-topmost', False)  # Reset the topmost attribute
    
    root.lift()  # Bring the main window to the front
    root.focus_force()  # Focus on the main window
    
    root.mainloop()

if __name__ == "__main__":
    main()

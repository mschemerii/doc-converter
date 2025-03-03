#!/usr/bin/env python3
"""
Doc Converter GUI

A graphical user interface for the Doc to Docx Converter script.
This GUI allows users to select a .doc file, specify an output path (optional),
and convert the file to .docx format using the convertDoc2Docx.py script.

Usage:
    python doc_converter_gui.py
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from convertDoc2Docx import doc_to_docx


class DocConverterGUI:
    """
    GUI application for converting .doc files to .docx format.
    """
    
    def __init__(self, root):
        """
        Initialize the GUI application.
        
        Args:
            root: The tkinter root window.
        """
        self.root = root
        self.root.title("Doc to Docx Converter")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # Set minimum window size
        self.root.minsize(500, 300)
        
        # Variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.status = tk.StringVar()
        self.status.set("Ready")
        
        # Create the main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create the input file selection frame
        input_frame = ttk.LabelFrame(main_frame, text="Input .doc File", padding="10")
        input_frame.pack(fill=tk.X, pady=10)
        
        ttk.Entry(input_frame, textvariable=self.input_path, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(input_frame, text="Browse...", command=self.browse_input).pack(side=tk.RIGHT)
        
        # Create the output file selection frame
        output_frame = ttk.LabelFrame(main_frame, text="Output .docx File (Optional)", padding="10")
        output_frame.pack(fill=tk.X, pady=10)
        
        ttk.Entry(output_frame, textvariable=self.output_path, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(output_frame, text="Browse...", command=self.browse_output).pack(side=tk.RIGHT)
        
        # Create the conversion button
        convert_button = ttk.Button(main_frame, text="Convert", command=self.convert_file)
        convert_button.pack(pady=20)
        
        # Create the status frame
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=10)
        
        ttk.Label(status_frame, text="Status:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(status_frame, textvariable=self.status).pack(side=tk.LEFT)
        
        # Create a progress bar
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=100, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=10)
        
        # Center the window on the screen
        self.center_window()
    
    def center_window(self):
        """Center the window on the screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def browse_input(self):
        """Open a file dialog to select the input .doc file."""
        file_path = filedialog.askopenfilename(
            title="Select .doc file",
            filetypes=[("Word Documents", "*.doc"), ("All Files", "*.*")]
        )
        if file_path:
            self.input_path.set(file_path)
            
            # If output path is empty, set a default based on input path
            if not self.output_path.get():
                base_path = os.path.splitext(file_path)[0]
                self.output_path.set(f"{base_path}.docx")
    
    def browse_output(self):
        """Open a file dialog to select the output .docx file location."""
        file_path = filedialog.asksaveasfilename(
            title="Save .docx file as",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if file_path:
            self.output_path.set(file_path)
    
    def convert_file(self):
        """Convert the selected .doc file to .docx format."""
        input_path = self.input_path.get()
        output_path = self.output_path.get() if self.output_path.get() else None
        
        # Validate input
        if not input_path:
            messagebox.showerror("Error", "Please select an input .doc file.")
            return
        
        if not os.path.exists(input_path):
            messagebox.showerror("Error", f"Input file '{input_path}' does not exist.")
            return
        
        if not input_path.lower().endswith('.doc'):
            messagebox.showerror("Error", f"Input file '{input_path}' is not a .doc file.")
            return
        
        # Start progress bar
        self.progress.start()
        self.status.set("Converting...")
        self.root.update()
        
        # Run the conversion in a separate thread to avoid freezing the GUI
        self.root.after(100, self.run_conversion, input_path, output_path)
    
    def run_conversion(self, input_path, output_path):
        """
        Run the conversion process.
        
        Args:
            input_path (str): Path to the input .doc file.
            output_path (str, optional): Path to save the output .docx file.
        """
        try:
            # Call the conversion function from convertDoc2Docx.py
            result = doc_to_docx(input_path, output_path)
            
            # Stop progress bar
            self.progress.stop()
            
            # Update status based on result
            if result:
                self.status.set("Conversion completed successfully.")
                messagebox.showinfo("Success", f"Converted file saved at:\n{result}")
            else:
                self.status.set("Conversion failed.")
                messagebox.showerror("Error", "Conversion failed. Check console for details.")
        except Exception as e:
            # Stop progress bar
            self.progress.stop()
            
            # Update status
            self.status.set("Error during conversion.")
            messagebox.showerror("Error", f"An error occurred during conversion:\n{str(e)}")


def main():
    """Main function to create and run the GUI application."""
    root = tk.Tk()
    app = DocConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
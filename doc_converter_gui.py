#!/usr/bin/env python3
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import traceback

# Import the existing conversion script
from doc_to_docx_converter import convert_doc_to_docx

class RedirectText:
    """Redirect print statements to a tkinter Text widget"""
    def __init__(self, text_widget):
        self.output = text_widget
    
    def write(self, string):
        self.output.insert(tk.END, string)
        self.output.see(tk.END)
    
    def flush(self):
        pass

class DocConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("Doc Converter")
        master.geometry("500x400")
        
        # File selection frame
        self.file_frame = tk.Frame(master)
        self.file_frame.pack(pady=10, padx=10, fill=tk.X)
        
        self.file_path = tk.StringVar()
        self.file_entry = tk.Entry(self.file_frame, textvariable=self.file_path, width=40)
        self.file_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))
        
        self.browse_button = tk.Button(self.file_frame, text="Browse", command=self.browse_file)
        self.browse_button.pack(side=tk.RIGHT)
        
        # Convert button
        self.convert_button = tk.Button(master, text="Convert", command=self.start_conversion, state=tk.NORMAL)
        self.convert_button.pack(pady=10)
        
        # Exit button (initially disabled)
        self.exit_button = tk.Button(master, text="Exit", command=self.exit_app, state=tk.DISABLED)
        self.exit_button.pack(pady=10)
        
        # Output window
        self.output_window = None
    
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
        
        # Create output window
        self.create_output_window()
        
        # Disable convert button during conversion
        self.convert_button.config(state=tk.DISABLED)
        
        # Start conversion in a separate thread
        conversion_thread = threading.Thread(
            target=self.run_conversion, 
            args=(input_file,), 
            daemon=True
        )
        conversion_thread.start()
    
    def create_output_window(self):
        """Create a new window to display conversion output"""
        if self.output_window and self.output_window.winfo_exists():
            return
        
        self.output_window = tk.Toplevel(self.master)
        self.output_window.title("Conversion Output")
        self.output_window.geometry("600x400")
        
        # Text area for output
        self.output_text = scrolledtext.ScrolledText(
            self.output_window, 
            wrap=tk.WORD, 
            state=tk.NORMAL
        )
        self.output_text.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        
        # Redirect stdout to the text widget
        sys.stdout = RedirectText(self.output_text)
        sys.stderr = RedirectText(self.output_text)
        
        # Close event handler
        self.output_window.protocol("WM_DELETE_WINDOW", self.on_output_window_close)
    
    def run_conversion(self, input_file):
        """Perform the actual conversion"""
        try:
            # Perform conversion
            output_file = convert_doc_to_docx(input_file)
            print(f"Successfully converted file to: {output_file}")
            
            # Re-enable convert button
            self.master.after(0, self.convert_button.config, {"state": tk.NORMAL})
            
            # Enable exit button
            self.master.after(0, self.exit_button.config, {"state": tk.NORMAL})
        
        except Exception as e:
            print(f"Conversion Error: {str(e)}")
            traceback.print_exc()
            
            # Re-enable convert button
            self.master.after(0, self.convert_button.config, {"state": tk.NORMAL})
    
    def on_output_window_close(self):
        """Handle closing of output window"""
        if self.output_window:
            self.output_window.destroy()
            self.output_window = None
            
            # Reset stdout and stderr
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
    
    def exit_app(self):
        """Exit the application"""
        # Only allow exit if output window is closed
        if self.output_window and self.output_window.winfo_exists():
            messagebox.showwarning("Cannot Exit", "Please close the Output Window first.")
            return
        
        self.master.quit()

def main():
    root = tk.Tk()
    app = DocConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

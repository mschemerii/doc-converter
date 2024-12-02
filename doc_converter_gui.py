#!/usr/bin/env python3
# Executable GUI for Doc Converter
# Last updated: 2023-03-08 14:30:00

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
        
        # Configure grid weights for centering
        master.grid_rowconfigure(0, weight=1)
        master.grid_rowconfigure(4, weight=1)
        master.grid_columnconfigure(0, weight=1)
        
        # Main container frame
        self.main_frame = tk.Frame(master)
        self.main_frame.grid(row=1, column=0, sticky="nsew")
        
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
        
        # Show Output button
        self.show_output_button = tk.Button(self.main_frame, text="Show Output", 
                                          command=self.show_output_window,
                                          state=tk.NORMAL, width=15)
        self.show_output_button.grid(row=2, column=0, pady=10)
        
        # Exit button (initially disabled)
        self.exit_button = tk.Button(self.main_frame, text="Exit", command=self.exit_app, 
                                   state=tk.DISABLED, width=15)
        self.exit_button.grid(row=3, column=0, pady=10)
        
        # Output window
        self.output_window = None
        
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
    
    def show_output_window(self):
        """Show the output window with previous output"""
        self.create_output_window()
        # Replay previous output
        if not self.last_output:
            self.output_text.insert(tk.END, "No output available.\n")
        for message in self.last_output:
            self.output_text.insert(tk.END, message)
    
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
            self.output_window.lift()
            return
        
        self.output_window = tk.Toplevel(self.master)
        self.output_window.title("Conversion Output")
        self.output_window.geometry("600x400")
        
        # Configure grid weights for output window
        self.output_window.grid_rowconfigure(0, weight=1)
        self.output_window.grid_columnconfigure(0, weight=1)
        
        # Frame for output content
        output_frame = tk.Frame(self.output_window)
        output_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        output_frame.grid_rowconfigure(0, weight=1)
        output_frame.grid_columnconfigure(0, weight=1)
        
        # Text area for output
        self.output_text = scrolledtext.ScrolledText(
            output_frame, 
            wrap=tk.WORD, 
            state=tk.NORMAL
        )
        self.output_text.grid(row=0, column=0, sticky="nsew")
        
        # Exit button for output window
        output_exit_button = tk.Button(output_frame, text="Close", 
                                     command=self.on_output_window_close, width=15)
        output_exit_button.grid(row=1, column=0, pady=10)
        
        # Custom stdout redirector that also stores output
        class StoringRedirectText(RedirectText):
            def __init__(self, text_widget, output_store):
                super().__init__(text_widget)
                self.output_store = output_store
            
            def write(self, string):
                super().write(string)
                self.output_store.append(string)
        
        # Redirect stdout to the text widget and store output
        sys.stdout = StoringRedirectText(self.output_text, self.last_output)
        sys.stderr = StoringRedirectText(self.output_text, self.last_output)
        
        # Close event handler
        self.output_window.protocol("WM_DELETE_WINDOW", self.on_output_window_close)
        
        # Enable the show output button
        self.show_output_button.config(state=tk.NORMAL)
    
    def run_conversion(self, input_file):
        """Perform the actual conversion"""
        try:
            # Import the full document processing function
            from process_document import process_document
            
            # Perform full document processing
            success = process_document(input_file)
            
            if success:
                print(f"Successfully processed document: {input_file}")
            else:
                print(f"Failed to process document: {input_file}")
            
            # Re-enable convert button
            self.master.after(0, self.convert_button.config, {"state": tk.NORMAL})
            
            # Enable exit button
            self.master.after(0, self.exit_button.config, {"state": tk.NORMAL})
        
        except Exception as e:
            print(f"Document Processing Error: {str(e)}")
            import traceback
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
    
    # Center the window on the screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'+{x}+{y}')
    
    root.mainloop()

if __name__ == "__main__":
    main()

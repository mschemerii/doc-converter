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
    format='%(asctime)s - %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler('doc_converter.log', mode='a'),
        logging.StreamHandler(sys.stdout)
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
        
        # Exit button (always enabled)
        self.exit_button = tk.Button(self.main_frame, text="Exit", command=self.exit_app, 
                                   state=tk.NORMAL, width=15)
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
        """Create or show output window"""
        if self.output_window is None or not self.output_window.winfo_exists():
            # Create new window
            self.output_window = tk.Toplevel(self.master)
            self.output_window.title("Conversion Output")
            self.output_window.geometry("600x400")
            
            # Configure grid
            self.output_window.grid_rowconfigure(0, weight=1)
            self.output_window.grid_columnconfigure(0, weight=1)
            
            # Create frame for text and scrollbar
            text_frame = tk.Frame(self.output_window)
            text_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10,0))
            text_frame.grid_rowconfigure(0, weight=1)
            text_frame.grid_columnconfigure(0, weight=1)
            
            # Add text widget with scrollbar
            self.output_text = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD)
            self.output_text.grid(row=0, column=0, sticky="nsew")
            
            # Create button frame
            button_frame = tk.Frame(self.output_window)
            button_frame.grid(row=1, column=0, pady=10)
            
            # Add Copy button
            copy_button = tk.Button(
                button_frame,
                text="Copy Output",
                command=self.copy_output_to_clipboard,
                width=15
            )
            copy_button.pack(side=tk.LEFT, padx=5)
            
            # Add Close button
            close_button = tk.Button(
                button_frame,
                text="Close",
                command=self.output_window.destroy,
                width=15
            )
            close_button.pack(side=tk.LEFT, padx=5)
            
            # Custom stdout redirector that also stores output
            class StoringRedirectText:
                def __init__(self, text_widget, output_store):
                    self.text_widget = text_widget
                    self.output_store = output_store
                
                def write(self, string):
                    self.text_widget.insert(tk.END, string)
                    self.text_widget.see(tk.END)
                    self.output_store.append(string)
                
                def flush(self):
                    pass
            
            # Redirect stdout and stderr to the text widget and store output
            sys.stdout = StoringRedirectText(self.output_text, self.last_output)
            sys.stderr = StoringRedirectText(self.output_text, self.last_output)
            
            # Set up window close handler
            self.output_window.protocol("WM_DELETE_WINDOW", self.on_output_window_close)
    
    def copy_output_to_clipboard(self):
        """Copy output text to clipboard"""
        if hasattr(self, 'output_text'):
            output_content = self.output_text.get("1.0", tk.END).strip()
            self.master.clipboard_clear()
            self.master.clipboard_append(output_content)
            
            # Show brief confirmation
            self.show_copy_confirmation()
    
    def show_copy_confirmation(self):
        """Show a small popup confirming the copy action"""
        popup = tk.Toplevel(self.master)
        popup.title("")
        
        # Position near the cursor
        x = self.master.winfo_pointerx()
        y = self.master.winfo_pointery()
        popup.geometry(f"+{x+10}+{y+10}")
        
        # Remove window decorations
        popup.overrideredirect(True)
        
        # Add label
        label = tk.Label(popup, text="Copied to clipboard!", padx=10, pady=5)
        label.pack()
        
        # Auto-close after 1 second
        popup.after(1000, popup.destroy)
    
    def run_conversion(self, input_file):
        """Perform the actual conversion"""
        try:
            # Import the full document processing function
            from process_document import process_document
            
            output_file = input_file.replace('.doc', '.docx')  # Define output file path
            
            # Perform full document processing
            logging.info(f"Starting conversion for: {input_file}")
            self.last_output.append(f"Starting conversion for: {input_file}")
            success = process_document(input_file, output_file)
            logging.info(f"Conversion successful: {success}")
            self.last_output.append(f"Conversion successful: {success}")
            
            if success:
                # Show a clear success popup
                self.master.after(0, self.show_success_popup, input_file)
                print(f"Successfully processed document: {input_file}")
                self.last_output.append(f"Successfully processed document: {input_file}")
            else:
                # Show an error popup if processing failed
                self.master.after(0, self.show_error_popup, "Document processing failed")
                print(f"Failed to process document: {input_file}")
                self.last_output.append(f"Failed to process document: {input_file}")
        except Exception as e:
            logging.error(f"Error during conversion: {e}")
            self.last_output.append(f"Error during conversion: {e}")
            self.show_error_popup("An unexpected error occurred during processing.")
    
    def show_success_popup(self, input_file):
        """Show a clear success popup with details about the processed document"""
        # Get the directory where the processed files are located
        directory = os.path.dirname(input_file) or '.'
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        
        # Construct a message with details
        message = (
            f"Document Processing Complete!\n\n"
            f"Original File: {input_file}\n\n"
            f"Processed Files Location: {directory}\n\n"
            f"Files Created:\n"
            f"- {base_name}.docx\n"
            f"- {base_name}_with_headers.docx\n"
            f"- Other modified copies"
        )
        
        # Create a custom popup window
        popup = tk.Toplevel(self.master)
        popup.title("Processing Successful")
        popup.geometry("400x300")
        popup.grab_set()  # Make the popup modal
        
        # Success icon (you can customize this)
        success_label = tk.Label(popup, text="✅", font=("Arial", 48))
        success_label.pack(pady=(20, 10))
        
        # Message text
        message_label = tk.Label(
            popup, 
            text=message, 
            font=("Arial", 10), 
            justify=tk.LEFT,
            wraplength=350
        )
        message_label.pack(padx=20, pady=10)
        
        # Close button
        close_button = tk.Button(
            popup, 
            text="Close", 
            command=popup.destroy, 
            width=15
        )
        close_button.pack(pady=20)
    
    def show_error_popup(self, error_message):
        """Show a clear error popup with details about the processing failure"""
        # Create a custom error popup window
        popup = tk.Toplevel(self.master)
        popup.title("Processing Error")
        popup.geometry("400x250")
        popup.grab_set()  # Make the popup modal
        
        # Error icon (you can customize this)
        error_label = tk.Label(popup, text="❌", font=("Arial", 48))
        error_label.pack(pady=(20, 10))
        
        # Error message text
        message_label = tk.Label(
            popup, 
            text=f"Document Processing Failed:\n\n{error_message}", 
            font=("Arial", 10), 
            fg="red",
            justify=tk.CENTER,
            wraplength=350
        )
        message_label.pack(padx=20, pady=10)
        
        # Close button
        close_button = tk.Button(
            popup, 
            text="Close", 
            command=popup.destroy, 
            width=15
        )
        close_button.pack(pady=20)
    
    def exit_app(self):
        """Handle application exit"""
        # Just close any output window and quit
        if self.output_window and self.output_window.winfo_exists():
            self.output_window.destroy()
        self.master.quit()
    
    def on_output_window_close(self):
        """Handle closing of output window"""
        # Reset stdout and stderr
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        
        if self.output_window is not None:
            self.output_window.destroy()

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

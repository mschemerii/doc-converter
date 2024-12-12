#!/usr/bin/env python3
# Executable GUI for Doc Converter
# Last updated: 2024-03-14

import subprocess
import sys
import os

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import logging
import threading
import traceback

# Debugging
print(f"Python version: {sys.version}")
print(f"Current working directory: {os.getcwd()}")
print(f"Sys path: {sys.path}")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s',
    handlers=[
        logging.FileHandler('doc_converter.log', mode='a'),
        logging.StreamHandler(sys.stdout)
    ]
)

class RedirectText:
    """Redirect print statements to a tkinter Text widget"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
    
    def write(self, string):
        try:
            self.text_widget.insert(tk.END, string)
            self.text_widget.see(tk.END)
            self.text_widget.update()
        except Exception as e:
            logging.error(f"Error writing to text widget: {e}")
    
    def flush(self):
        pass

class GUILogHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)
        self.text_widget.update()

class DocConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("Doc Converter")
        master.geometry("500x400")
        
        # Initialize output window
        self.output_window = None
        self.output_text = None
        
        # Create output window immediately
        self.create_output_window()
        
        # Configure grid weights for centering
        master.grid_rowconfigure(0, weight=1)
        master.grid_rowconfigure(4, weight=1)
        master.grid_columnconfigure(0, weight=1)
        
        # Configure logging to use our GUI handler
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        gui_handler = GUILogHandler(self.output_text)
        gui_handler.setFormatter(logging.Formatter('%(message)s'))
        root_logger.handlers = []  # Remove any existing handlers
        root_logger.addHandler(gui_handler)
        
        # Main container frame
        self.main_frame = tk.Frame(master)
        self.main_frame.grid(row=1, column=0, sticky="nsew")
        
        # Configure main frame grid
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Instructions frame with border
        self.instructions_frame = tk.LabelFrame(self.main_frame, text="Instructions", padx=10, pady=5)
        self.instructions_frame.grid(row=0, column=0, pady=10, padx=20, sticky="ew")
        
        # Instructions text
        instructions = [
            "1. Click 'Browse' to select a .doc file",
            "2. Click 'Convert' to start the conversion process",
            "3. Use 'Show Output' to view the conversion progress",
            "4. Click 'Exit' when finished"
        ]
        for i, instruction in enumerate(instructions):
            tk.Label(self.instructions_frame, text=instruction, anchor="w", justify=tk.LEFT).grid(row=i, column=0, sticky="w")
        
        # File selection frame
        self.file_frame = tk.Frame(self.main_frame)
        self.file_frame.grid(row=1, column=0, pady=10, padx=20, sticky="ew")
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
        self.convert_button.grid(row=2, column=0, pady=10)
        
        # Show Output button
        self.show_output_button = tk.Button(self.main_frame, text="Show Output", 
                                          command=self.show_output_window,
                                          state=tk.NORMAL, width=15)
        self.show_output_button.grid(row=3, column=0, pady=10)
        
        # Exit button (always enabled)
        self.exit_button = tk.Button(self.main_frame, text="Exit", command=self.exit_app, 
                                   state=tk.NORMAL, width=15)
        self.exit_button.grid(row=4, column=0, pady=10)
        
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
        """Show the output window"""
        if self.output_window is None or not self.output_window.winfo_exists():
            self.create_output_window()
        self.output_window.deiconify()
        self.output_window.lift()

    def start_conversion(self):
        """Start conversion process in a separate thread"""
        input_file = self.file_path.get()
        
        if not input_file:
            messagebox.showerror("Error", "Please select a file first")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("Error", "Selected file does not exist")
            return
        
        # Clear previous output
        self.last_output = []
        
        # Show output window
        self.show_output_window()
        
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
        """Create and configure the output window"""
        if self.output_window is not None and self.output_window.winfo_exists():
            self.output_window.lift()
            return

        self.output_window = tk.Toplevel(self.master)
        self.output_window.title("Conversion Output")
        self.output_window.geometry("600x400")
        
        # Configure grid
        self.output_window.grid_rowconfigure(0, weight=1)
        self.output_window.grid_columnconfigure(0, weight=1)
        
        # Frame for text widget and scrollbar
        text_frame = tk.Frame(self.output_window)
        text_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)
        
        # Text widget with scrollbar
        self.output_text = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, width=70, height=20)
        self.output_text.grid(row=0, column=0, sticky="nsew")
        
        # Button frame
        button_frame = tk.Frame(self.output_window)
        button_frame.grid(row=1, column=0, pady=5)
        
        # Copy button
        copy_button = tk.Button(button_frame, text="Copy Output", command=self.copy_output)
        copy_button.pack(side=tk.LEFT, padx=5)
        
        # Close button
        close_button = tk.Button(button_frame, text="Close", command=self.hide_output_window)
        close_button.pack(side=tk.LEFT, padx=5)
        
        # Redirect stdout and stderr to the text widget
        sys.stdout = RedirectText(self.output_text)
        sys.stderr = RedirectText(self.output_text)
        
        # Initially hide the window
        self.output_window.withdraw()
        
        # Close event handler
        self.output_window.protocol("WM_DELETE_WINDOW", self.hide_output_window)

    def run_conversion(self, input_file):
        """Perform the actual conversion"""
        try:
            # Import the full document processing function
            from process_document import process_document
            
            # Perform full document processing
            success = process_document(input_file)
            
            if success:
                # Show a clear success popup
                self.master.after(0, self.show_success_popup, input_file)
                print(f"Successfully processed document: {input_file}")
            else:
                # Show an error popup if processing failed
                self.master.after(0, self.show_error_popup, "Document processing failed")
                print(f"Failed to process document: {input_file}")
            
            # Re-enable convert button
            self.master.after(0, self.convert_button.config, {"state": tk.NORMAL})
            
            # Enable exit button
            self.master.after(0, self.exit_button.config, {"state": tk.NORMAL})
        
        except Exception as e:
            print(f"Document Processing Error: {str(e)}")
            import traceback
            traceback.print_exc()
            
            # Show an error popup
            self.master.after(0, self.show_error_popup, str(e))
            
            # Re-enable convert button
            self.master.after(0, self.convert_button.config, {"state": tk.NORMAL})
    
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
    
    def hide_output_window(self):
        """Hide the output window instead of destroying it"""
        if self.output_window and self.output_window.winfo_exists():
            self.output_window.withdraw()
    
    def exit_app(self):
        """Handle application exit"""
        if self.output_window and self.output_window.winfo_exists():
            self.output_window.destroy()
        self.master.destroy()
    
    def copy_output(self):
        """Copy the contents of the output window to clipboard"""
        if self.output_text:
            output_content = self.output_text.get("1.0", tk.END)
            self.master.clipboard_clear()
            self.master.clipboard_append(output_content)
            messagebox.showinfo("Success", "Output copied to clipboard!")

def main():
    root = tk.Tk()
    
    # Force window to front on macOS
    root.lift()
    root.attributes('-topmost', True)
    app = DocConverterApp(root)
    
    # Center the window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'+{x}+{y}')
    
    # Disable topmost after window appears
    root.after_idle(root.attributes, '-topmost', False)
    
    root.mainloop()

if __name__ == "__main__":
    main()

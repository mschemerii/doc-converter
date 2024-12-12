#!/usr/bin/env python3
import os
import sys
import traceback
import datetime
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

class DocConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Doc Converter")
        self.selected_file = None
        self.output_window = None
        
        # Set window size and position
        window_width = 600
        window_height = 400
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Instructions
        instructions = ttk.Label(
            main_frame,
            text="Select a .doc file to convert to modern format.",
            wraplength=500
        )
        instructions.grid(row=0, column=0, columnspan=2, pady=10)
        
        # File selection frame
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky=tk.W+tk.E)
        
        # Browse button
        self.browse_button = ttk.Button(
            file_frame,
            text="Browse",
            command=self.browse_file
        )
        self.browse_button.pack(side=tk.LEFT, padx=5)
        
        # Selected file label
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Convert button
        self.convert_button = ttk.Button(
            main_frame,
            text="Convert",
            command=self.convert_file,
            state=tk.DISABLED
        )
        self.convert_button.grid(row=2, column=0, pady=10)
        
        # Show Output button
        self.show_output_button = ttk.Button(
            main_frame,
            text="Show Output",
            command=self.show_output,
            state=tk.DISABLED
        )
        self.show_output_button.grid(row=2, column=1, pady=10)
        
        # Exit button
        exit_button = ttk.Button(
            main_frame,
            text="Exit",
            command=self.root.quit
        )
        exit_button.grid(row=3, column=0, columnspan=2, pady=10)
        
        # Configure grid weights
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Document",
            filetypes=[("Word Documents", "*.doc"), ("All Files", "*.*")]
        )
        if file_path:
            self.selected_file = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.convert_button.config(state=tk.NORMAL)
    
    def convert_file(self):
        if not self.selected_file:
            return
            
        try:
            # Import process_document here to ensure it's available
            import process_document
            
            # Process the document
            process_document.process_document(self.selected_file)
            
            messagebox.showinfo("Success", "Document converted successfully!")
            self.show_output_button.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing document: {str(e)}")
    
    def show_output(self):
        if self.output_window is None or not tk.Toplevel.winfo_exists(self.output_window):
            self.output_window = tk.Toplevel(self.root)
            self.output_window.title("Conversion Output")
            self.output_window.geometry("500x300")
            
            output_text = scrolledtext.ScrolledText(
                self.output_window,
                wrap=tk.WORD,
                width=60,
                height=20
            )
            output_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
            
            # Load and display the output
            try:
                with open(os.path.expanduser('~/Desktop/doc_converter_debug.log'), 'r') as f:
                    output_text.insert(tk.END, f.read())
            except Exception as e:
                output_text.insert(tk.END, f"Error loading output: {str(e)}")
            
            output_text.see(tk.END)

def main():
    root = tk.Tk()
    app = DocConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

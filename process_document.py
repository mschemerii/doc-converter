#!/usr/bin/env python3
import sys
import os
import subprocess
import time
from pathlib import Path
import platform
import logging
import traceback

def process_document(doc_path):
    """Process a .doc file through all conversion and modification steps"""
    try:
        logging.info(f"Starting document processing for {doc_path}")
        
        # Validate input file
        if not os.path.exists(doc_path):
            logging.error(f"Input file not found: {doc_path}")
            raise FileNotFoundError(f"Input file not found: {doc_path}")
        
        # Get the directory and base filename
        directory = os.path.dirname(doc_path) or '.'
        filename = os.path.basename(doc_path)
        filename = filename.replace('+', '').replace('+-+', '_').replace(' ', '')
        base_name, ext = os.path.splitext(filename)
        
        if ext.lower() != '.doc':
            logging.error("Error: Input file must be a .doc file")
            return False
        
        # Define the intermediate docx filename
        docx_path = os.path.join(directory, f"{base_name}.docx")
        
        # Step 1: Convert .doc to .docx
        python_cmd = 'python3'
        
        result = subprocess.run(
            [python_cmd, 'doc_to_docx_converter.py', doc_path],
            capture_output=True,
            text=True,
            check=True,
            env=os.environ
        )
        logging.info(result.stdout)
        if result.stderr:
            logging.warning("Warnings: " + result.stderr)
        
        # Small delay to ensure file is ready
        time.sleep(1)
        
        # Step 2: Modify table properties
        result = subprocess.run(
            [python_cmd, 'modify_docx_tables.py', docx_path],
            capture_output=True,
            text=True,
            check=True,
            env=os.environ
        )
        logging.info(result.stdout)
        if result.stderr:
            logging.warning("Warnings: " + result.stderr)
        
        # Step 3: Add table rows
        result = subprocess.run(
            [python_cmd, 'add_table_rows.py', docx_path],
            capture_output=True,
            text=True,
            check=True,
            env=os.environ
        )
        logging.info(result.stdout)
        if result.stderr:
            logging.warning("Warnings: " + result.stderr)
        
        # Step 4: Create renamed copies with headers
        result = subprocess.run(
            [python_cmd, 'rename_docx.py', docx_path],
            capture_output=True,
            text=True,
            check=True,
            env=os.environ
        )
        logging.info(result.stdout)
        if result.stderr:
            logging.warning("Warnings: " + result.stderr)
        
        logging.info("\n=== Processing Complete ===")
        logging.info(f"Original .doc file: {doc_path}")
        logging.info(f"Intermediate .docx file: {docx_path}")
        logging.info("Final files created with appropriate headers and content modifications.")
        return True
    
    except Exception as e:
        logging.critical(f"Document processing failed: {e}")
        logging.critical(traceback.format_exc())
        raise

def main():
    """Main entry point for document processing"""
    try:
        if len(sys.argv) != 2:
            logging.error("Usage: python3 process_document.py <doc_file>")
            return False
        
        input_file = sys.argv[1]
        return process_document(input_file)
    
    except Exception as e:
        logging.critical(f"Main execution failed: {e}")
        logging.critical(traceback.format_exc())
        return False

if __name__ == "__main__":
    # Only configure logging if running as main script
    # When imported by GUI, we'll use GUI's logging configuration
    if not logging.getLogger().handlers:
        logging.basicConfig(
            level=logging.INFO,
            format='%(message)s',
            handlers=[logging.StreamHandler()]
        )
    main()

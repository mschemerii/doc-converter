#!/usr/bin/env python3
import sys
import logging
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Add more detailed logging to capture the flow of execution
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def ensure_element(parent, tag_name):
    """Ensure an XML element exists, create it if it doesn't"""
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    element = parent.find(f'.//w:{tag_name}', namespace)
    if element is None:
        element = parse_xml(f'<w:{tag_name} {nsdecls("w")}/>') 
        parent.append(element)
    return element

def modify_table_properties(docx_path):
    """
    Modify table properties in a .docx file:
    1. Remove auto-resize
    2. Remove fixed column widths
    3. Set table width to 100%
    """
    try:
        # Load the document
        doc = Document(docx_path)
        logging.info(f"Successfully loaded document: {docx_path}")
    except Exception as e:
        logging.error(f"Error loading document {docx_path}: {e}")
        sys.exit(1)

    # Process each table in the document
    for table in doc.tables:
        try:
            logging.info(f"Processing table: {table}")
            # Get or create table properties
            tblPr = ensure_element(table._element, 'tblPr')
            logging.info(f"Ensured table properties element for table: {table}")
            
            # Set table layout to fixed (disables auto-fit)
            layout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
            existing_layout = tblPr.find('.//w:tblLayout', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if existing_layout is not None:
                tblPr.remove(existing_layout)
                logging.info(f"Removed existing table layout for table: {table}")
            tblPr.append(layout)
            logging.info(f"Set table layout to fixed for table: {table}")
            
            # Set table width to 100%
            tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>')
            existing_width = tblPr.find('.//w:tblW', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if existing_width is not None:
                tblPr.remove(existing_width)
                logging.info(f"Removed existing table width for table: {table}")
            tblPr.append(tblW)
            logging.info(f"Set table width to 100% for table: {table}")
            
            # Remove fixed widths from columns
            grid = table._element.find('.//w:tblGrid', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if grid is not None:
                for grid_col in grid.findall('.//w:gridCol', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    if 'w:w' in grid_col.attrib:
                        del grid_col.attrib['w:w']
                        logging.info(f"Removed fixed width from column in table: {table}")
            
            # Remove width settings from cells
            for tc in table._element.findall('.//w:tc', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                tcPr = tc.find('.//w:tcPr', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                if tcPr is not None:
                    tcW = tcPr.find('.//w:tcW', {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
                    if tcW is not None:
                        tcPr.remove(tcW)
                        logging.info(f"Removed width setting from cell in table: {table}")
        except Exception as e:
            logging.error(f"Error processing table: {e}")
            continue

    try:
        # Save the modifications back to the same file
        doc.save(docx_path)
        logging.info(f"Successfully modified tables in {docx_path}")
    except Exception as e:
        logging.error(f"Error saving document {docx_path}: {e}")
        sys.exit(1)

def main():
    if len(sys.argv) != 2:
        print("Usage: python modify_docx_tables.py <path_to_docx_file>")
        sys.exit(1)
    
    try:
        docx_path = sys.argv[1]
        modify_table_properties(docx_path)
        print(f"Successfully modified table properties in: {docx_path}")
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()

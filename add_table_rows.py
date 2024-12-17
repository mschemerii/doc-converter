#!/usr/bin/env python3
import sys
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def has_content(row):
    """Check if a row has any non-empty content."""
    return any(cell.text.strip() for cell in row.cells)

def merge_cells_in_row(row):
    """Merge all cells in a row into a single cell."""
    if len(row.cells) > 1:
        row.cells[0].merge(row.cells[-1])

def should_process_table(table):
    """Determine if a table should be processed."""
    # Keywords for skipping specific tables
    skip_keywords = [
        "Change request numbers",
        "Manifests",
        "Sync Multi Manifest String"
    ]
    # Get text of the first cell in the table's first row
    first_cell_text = table.rows[0].cells[0].text.strip() if table.rows else ""
    return not any(keyword in first_cell_text for keyword in skip_keywords)

def add_and_merge_rows_to_table(table):
    """Add a row with merged cells after rows containing content."""
    added_rows = 0

    # Process table rows
    for i in range(len(table.rows)):
        current_row = table.rows[i + added_rows]

        # Check if this row has content and insert a new row after it
        if has_content(current_row):
            # Insert a new row
            new_row = table.add_row()
            merge_cells_in_row(new_row)  # Merge all cells in the inserted row

            # Move the new row directly after the current row
            table._tbl.remove(new_row._tr)
            table._tbl.insert(current_row._tr.getparent().index(current_row._tr) + 1, new_row._tr)
            added_rows += 1

def process_document(docx_path):
    """Process the document to add rows to specific tables."""
    # Load the document
    doc = Document(docx_path)

    # Process tables
    for table in doc.tables:
        if should_process_table(table):
            add_and_merge_rows_to_table(table)

    # Save the modified document
    doc.save(docx_path)

def main():
    if len(sys.argv) != 2:
        print("Usage: python add_table_rows.py <path_to_docx_file>")
        sys.exit(1)

    try:
        docx_path = sys.argv[1]
        process_document(docx_path)
        print(f"Successfully processed tables in: {docx_path}")
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main()

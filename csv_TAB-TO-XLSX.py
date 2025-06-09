import sys
import os
import csv
from tkinter import Tk, filedialog, messagebox
import xlsxwriter

def is_file_locked(filepath):
    """Check if the file is locked by attempting to open it in write mode."""
    try:
        with open(filepath, 'a') as f:
            return False
    except IOError:
        return True

def handle_locked_file(output_file, root):
    """Handle case where output file is locked by prompting user for action."""
    while True:
        response = messagebox.askyesnocancel(
            "File In Use",
            f"The file {output_file} is open or locked (possibly in Excel).\n\n"
            "Yes: Close the file manually and retry.\n"
            "No: Save with a new filename.\n"
            "Cancel: Exit the script.",
            parent=root
        )
        
        if response is True:  # Yes: Retry
            if not is_file_locked(output_file):
                return output_file
            continue
        elif response is False:  # No: Save as new file
            new_output = filedialog.asksaveasfilename(
                title="Save output Excel file as",
                initialfile=os.path.splitext(os.path.basename(output_file))[0] + '_new.xlsx',
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                parent=root
            )
            if new_output and not is_file_locked(new_output):
                return new_output
            elif not new_output:
                print("No new output file selected. Exiting...")
                sys.exit()
        else:  # Cancel or dialog closed
            print("Operation cancelled. Exiting...")
            sys.exit()

def detect_delimiter(file_path):
    """Detect if file is CSV or tab-delimited by checking the first line."""
    try:
        with open(file_path, 'r', newline='', encoding='utf-8') as f:
            first_line = f.readline()
            # Count commas and tabs in the first line
            comma_count = first_line.count(',')
            tab_count = first_line.count('\t')
            
            # Return delimiter based on which is more prevalent
            if tab_count > comma_count:
                return '\t'
            else:
                return ','
    except Exception as e:
        print(f"Error detecting delimiter: {e}. Defaulting to comma.")
        return ','

# Set working directory to user's home folder for faster file dialog
os.chdir(os.path.expanduser("~"))

# Initialize Tkinter
root = Tk()
root.update()
root.withdraw()
root.attributes('-topmost', True)  # Ensure dialogs are on top

# File picker for input file
input_file = filedialog.askopenfilename(
    title="Select input file",
    filetypes=[("Data files", "*.csv *.txt *.tsv *.xls"), ("All files", "*.*")],
    parent=root
)

if not input_file:
    print("No input file selected. Exiting...")
    root.destroy()
    sys.exit()

# Suggest output filename
input_filename = os.path.basename(input_file)
suggested_output = os.path.splitext(input_filename)[0] + '.xlsx'

# File picker for output file
output_file = filedialog.asksaveasfilename(
    title="Save output Excel file as",
    initialfile=suggested_output,
    defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    parent=root
)

if not output_file:
    print("No output file selected. Exiting...")
    root.destroy()
    sys.exit()

# Check if output file is locked
if is_file_locked(output_file):
    output_file = handle_locked_file(output_file, root)

# Detect delimiter
delimiter = detect_delimiter(input_file)

# Read input file with detected delimiter
with open(input_file, 'r', newline='', encoding='utf-8') as f:
    reader = csv.reader(f, delimiter=delimiter)
    rows = list(reader)

if not rows:
    print("Input file is empty. Exiting...")
    root.destroy()
    sys.exit()

# Create Excel workbook and worksheet with nan_inf_to_errors option
workbook = xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet()

# Define a bold format for headers
bold_header_format = workbook.add_format({'bold': True})

# Define a text format for all cells
text_format = workbook.add_format({'num_format': '@'})

# Track column widths for auto-sizing
col_widths = {}

# Write rows
for row_idx, row in enumerate(rows):
    for col_idx, cell in enumerate(row):
        cell_value = cell.strip()

        # Write all cells as text, with bold format for headers
        if row_idx == 0:
            worksheet.write(row_idx, col_idx, cell_value, bold_header_format)
        else:
            worksheet.write_string(row_idx, col_idx, cell_value, text_format)

        # Update max width
        width = len(cell_value)
        if col_idx not in col_widths or width > col_widths[col_idx]:
            col_widths[col_idx] = width

# Adjust column widths
for col_idx, width in col_widths.items():
    adjusted_width = width + 2
    final_width = min(adjusted_width, 50)  # Cap at 50
    worksheet.set_column(col_idx, col_idx, final_width)

# Close workbook
workbook.close()

print(f"File converted and saved to: {output_file}")
messagebox.showinfo(
    "Success",
    f"File conversion complete!\nSaved to: {output_file}",
    parent=root
)
root.destroy()

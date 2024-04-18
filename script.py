# Standard library imports
import sys
import os
import csv
import json
import xml.etree.ElementTree as ET
import warnings
from tkinter import filedialog, messagebox, ttk

# Related third-party imports
import openpyxl
import tkinter as tk
from tabulate import tabulate
import yaml
import pandas as pd

# Hide UserWarning
warnings.simplefilter(action='ignore', category=UserWarning)

# Initialize Tkinter
root = tk.Tk()
root.title("Excel-to-Mardown-Converter")

# Terminate the script when the window is closed
root.bind("<Destroy>", lambda e: sys.exit())

# Function to open the file selection dialog
def open_file_dialog():
    file_path = filedialog.askopenfilename(
        parent=root,
        title="Please select the Excel document.",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if file_path:
        process_workbook(file_path)

# Function to process the selected Excel file
def process_workbook(file_path):
    try:
        # Load Excel file with openpyxl
        workbook = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:
        messagebox.showinfo("Error", f"Error loading the Excel file: {str(e)}")
        return
    
    # List of available formats
    formats = ["pipe", "yaml", "xml", "html", "csv", "json", "plain", "jira", "mediawiki"]

    # Create a StringVar for the combobox
    selected_format = tk.StringVar()

    # Create a label
    tk.Label(root, text= "Select Excel worksheets to export:").pack(anchor='w')

    # List of available worksheets
    available_sheets = workbook.sheetnames

    # Check if worksheets are available
    if not available_sheets:
        messagebox.showinfo("Error", "The selected Excel file does not contain any worksheets.")
        return

    # Listbox for selecting worksheets
    listbox = tk.Listbox(root, selectmode='multiple', exportselection=0)
    listbox.pack(side='top', fill='both', expand=True)

    # Add worksheets to the listbox
    for sheet in available_sheets:
        listbox.insert(tk.END, sheet)

    # Initialize Checkbox
    export_visible = tk.BooleanVar(value=True)

    # Create Checkbox
    tk.Checkbutton(root, text="Exclude hidden rows/columns when exporting", variable=export_visible).pack(anchor='w')

    # Label for the combobox
    tk.Label(root, text="Choose desired format:").pack(anchor='w')

    # Create the combobox
    combobox = ttk.Combobox(root, textvariable=selected_format, state='readonly')
    combobox['values'] = formats
    combobox.current(0)  # set initial selection
    combobox.pack(anchor='w')

    # Functions for the buttons
    def select_all():
        listbox.select_set(0, tk.END)

    def deselect_all():
        listbox.selection_clear(0, tk.END)

    def cancel():
        root.destroy()

    from openpyxl.utils import column_index_from_string

    def export_sheets():
        selected_indices = listbox.curselection()
        selected_sheets = [listbox.get(i) for i in selected_indices]

        # Check if worksheets were selected
        if not selected_sheets:
            messagebox.showinfo("Information", "Please select worksheets")
            return

        for sheet_name in selected_sheets:
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                # Collect data
                data = []

                for row in sheet.iter_rows(min_row=1, max_col=sheet.max_column, max_row=sheet.max_row):
                    # Check if column is hidden - if checkbox is active
                    if export_visible.get() and sheet.row_dimensions[row[0].row].hidden:
                        continue
                    row_data = []
                    for cell in row:
                        # Check if row is hidden - if checkbox is active
                        column_dimension = sheet.column_dimensions.get(cell.column_letter)
                        if not export_visible.get() or (column_dimension is not None and not column_dimension.hidden):
                            row_data.append(cell.value)
                    # If row is not empty, transfer to data
                    if row_data:
                        data.append(row_data)

                # Convert data to chosen format
                format_table = tabulate(data, tablefmt=selected_format.get(), headers="firstrow", showindex=False)

                # Dictionary mapping formats to file extensions
                format_extensions = {"pipe": ".md", "plain": ".txt", "jira": ".txt", "mediawiki": ".txt", "csv": ".csv", "json": ".json", "yaml": ".yaml", "xml": ".xml", "html": ".html"}

                # Get the file extension for the selected format
                file_extension = format_extensions[selected_format.get()]

                # Path for the Markdown file
                output_file_path = f"{file_path.rsplit('.', 1)[0]}_{sheet_name}{file_extension}"

                # Check if filename already exists
                if os.path.isfile(output_file_path):
                    overwrite = messagebox.askyesno("File already exists", f"The file {output_file_path} already exists. Overwrite?")
                    if not overwrite:
                        return

                # Write the Markdown table into a file
                try:
                    with open(output_file_path, "w", encoding="utf-8") as f:
                        f.write(format_table)
                except Exception as e:
                    messagebox.showinfo("Error", f"Error writing the file {output_file_path}: {str(e)}")
                    return

        # Show success message
        messagebox.showinfo("Success", "File(s) successfully created.")
        root.quit()

    # Buttons
    button_select_all = tk.Button(root, text="Select all", command=select_all)
    button_select_all.pack(side='left', fill='x', expand=True)

    button_deselect_all = tk.Button(root, text="Deselect all", command=deselect_all)
    button_deselect_all.pack(side='left', fill='x', expand=True)

    button_export = tk.Button(root, text="Export", command=export_sheets)
    button_export.pack(side='left', fill='x', expand=True)

    button_cancel = tk.Button(root, text="Cancel", command=cancel)
    button_cancel.pack(side='left', fill='x', expand=True)

# Open the file selection dialog
open_file_dialog()

# Start Tkinter event loop
root.mainloop()

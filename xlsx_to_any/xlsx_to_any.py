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
root.title("Excel-to-Any-Converter")

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
    else:
        sys.exit()
        
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
        selected_sheets = [sheet.get() for sheet in sheet_checkboxes if sheet.get()]
        if not selected_sheets:
            messagebox.showinfo("Information", "Please select worksheets")
            return

        # Dictionary Mapping Formate zu File Extension
        format_extensions = {"pipe": ".md", "plain": ".txt", "jira": ".txt", "mediawiki": ".txt", "csv": ".csv", "json": ".json", "yaml": ".yaml", "xml": ".xml", "html": ".html"}

        for sheet_name in selected_sheets:
            data = read_sheet(sheet_name)

            # Daten in gewähltes Format konvertieren
            if selected_format.get() == "xml":
                output_file_path = f"{sheet_name}.xml"

                # Check ob Dateiname bereits existiert
                if os.path.isfile(output_file_path):
                    overwrite = messagebox.askyesno("File already exists", f"The file {output_file_path} already exists. Overwrite?")
                    if not overwrite:
                        return

                root = ET.Element("root")

                for row in data:
                    row_element = ET.SubElement(root, "row")
                    for cell in row:
                        cell_element = ET.SubElement(row_element, "cell")
                        cell_element.text = str(cell)

                tree = ET.ElementTree(root)

                try:
                    tree.write(output_file_path)
                except Exception as e:
                    messagebox.showinfo("Error", f"Error writing the file {output_file_path}: {str(e)}")
                    return

            elif selected_format.get() == "json":
                output_file_path = f"{sheet_name}.json"
                
                # Check ob Dateiname bereits existiert
                if os.path.isfile(output_file_path):
                    overwrite = messagebox.askyesno("File already exists", f"The file {output_file_path} already exists. Overwrite?")
                    if not overwrite:
                        return

                try:
                    with open(output_file_path, "w") as f:
                        json.dump(data, f)
                except Exception as e:
                    messagebox.showinfo("Error", f"Error writing the file {output_file_path}: {str(e)}")
                    return

            elif selected_format.get() == "yaml":
                output_file_path = f"{sheet_name}.yaml"
                
                # Check ob Dateiname bereits existiert
                if os.path.isfile(output_file_path):
                    overwrite = messagebox.askyesno("File already exists", f"The file {output_file_path} already exists. Overwrite?")
                    if not overwrite:
                        return

                try:
                    with open(output_file_path, "w") as f:
                        yaml.dump(data, f)
                except Exception as e:
                    messagebox.showinfo("Error", f"Error writing the file {output_file_path}: {str(e)}")
                    return

            else:
                output = tabulate(data, tablefmt=selected_format.get())
                output_file_path = f"{sheet_name}{format_extensions[selected_format.get()]}"
                
                # Check ob Dateiname bereits existiert
                if os.path.isfile(output_file_path):
                    overwrite = messagebox.askyesno("File already exists", f"The file {output_file_path} already exists. Overwrite?")
                    if not overwrite:
                        return

                try:
                    with open(output_file_path, "w") as f:
                        f.write(output)
                except Exception as e:
                    messagebox.showinfo("Error", f"Error writing the file {output_file_path}: {str(e)}")
                    return

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

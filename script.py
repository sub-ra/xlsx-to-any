import sys
import os
import openpyxl
import tkinter as tk
import warnings
from tkinter import filedialog, messagebox
from tabulate import tabulate

# Filter UserWarning in Console
warnings.simplefilter(action='ignore', category=UserWarning)

# Initialize Tkinter
root = tk.Tk()
root.title("Excel-to-Mardown-Converter")

# Exit Script on Window Close
root.bind("<Destroy>", lambda e: sys.exit())

# Function to open File Dialog
def open_file_dialog():
    file_path = filedialog.askopenfilename(
        parent=root,
        title="Bitte wählen Sie das Excel-Dokument aus.",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if file_path:
        process_workbook(file_path)

#Function to process selected XLSX File
def process_workbook(file_path):
    try:
        # Load XLSX with openpyxl
        workbook = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:
        messagebox.showinfo("Fehler", f"Fehler beim Laden der Excel-Datei: {str(e)}")
        return

    # Create Label
    tk.Label(root, text= "Excel-Arbeitsblätter zum Exportieren auswählen:").pack(anchor='w')

    # List Sheets
    available_sheets = workbook.sheetnames

    # Errorhandling - Check if Sheets are available
    if not available_sheets:
        messagebox.showinfo("Fehler", "Die ausgewählte Excel-Datei enthält keine Arbeitsblätter.")
        return

    # Listbox for selecting Sheets
    listbox = tk.Listbox(root, selectmode='multiple', exportselection=0)
    listbox.pack(side='top', fill='both', expand=True)

    # Add Sheets to Listbox
    for sheet in available_sheets:
        listbox.insert(tk.END, sheet)

    # Initialize Checkbox
    export_visible = tk.BooleanVar(value=True)

    # Create Checkbox
    tk.Checkbutton(root, text="Ausgeblendete Zeilen/Spalten beim Export nicht einbeziehen", variable=export_visible).pack(anchor='w')

    # Functions for Buttons
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

        # Errorhandling - Check if Sheets selected
        if not selected_sheets:
            messagebox.showinfo("Information", "Bitte wähle Arbeitsmappen aus")
            return

        for sheet_name in selected_sheets:
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                # Collect data
                data = []

                for row in sheet.iter_rows(min_row=1, max_col=sheet.max_column, max_row=sheet.max_row):
                    # Check Column Visibility - if Checkbox active
                    if export_visible.get() and sheet.row_dimensions[row[0].row].hidden:
                        continue
                    row_data = []
                    for cell in row:
                        # Check Row Visibility - if Checkbox active
                        column_dimension = sheet.column_dimensions.get(cell.column_letter)
                        if not export_visible.get() or (column_dimension is not None and not column_dimension.hidden):
                            row_data.append(cell.value)
                    # If Colum not empty -  transfer to Data
                    if row_data:
                        data.append(row_data)

                # Convert Data to Markdown Table
                markdown_table = tabulate(data, tablefmt="pipe", headers="firstrow", showindex=False)

                # Set Path for Markdown-File
                output_file_path = f"{file_path.rsplit('.', 1)[0]}_{sheet_name}.md"

                # Errorhandling - Check if Filename exists
                if os.path.isfile(output_file_path):
                    overwrite = messagebox.askyesno("Datei existiert bereits", f"Die Datei {output_file_path} existiert bereits. Überschreiben?")
                    if not overwrite:
                        return

                # Write Markdown-Table to File
                try:
                    with open(output_file_path, "w", encoding="utf-8") as f:
                        f.write(markdown_table)
                except Exception as e:
                    messagebox.showinfo("Fehler", f"Fehler beim Schreiben der Datei {output_file_path}: {str(e)}")
                    return

        # Success Message
        messagebox.showinfo("Erfolg", "MD-File(s) erfolgreich erstellt.")
        root.quit()

    # Buttons
    button_select_all = tk.Button(root, text="Alle auswählen", command=select_all)
    button_select_all.pack(side='left', fill='x', expand=True)

    button_deselect_all = tk.Button(root, text="Auswahl aufheben", command=deselect_all)
    button_deselect_all.pack(side='left', fill='x', expand=True)

    button_export = tk.Button(root, text="Exportieren", command=export_sheets)
    button_export.pack(side='left', fill='x', expand=True)

    button_cancel = tk.Button(root, text="Abbrechen", command=cancel)
    button_cancel.pack(side='left', fill='x', expand=True)

# Open File Dialog
open_file_dialog()

# Start Tkinter Event-Loop
root.mainloop()

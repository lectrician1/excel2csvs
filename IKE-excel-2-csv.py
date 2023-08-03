import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def select_file():
    filename = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xls *.xlsx')])
    entry_excel.delete(0, tk.END)
    entry_excel.insert(0, filename)

def select_output_dir():
    directory = filedialog.askdirectory()
    entry_output.delete(0, tk.END)
    entry_output.insert(0, directory)

def sort_and_save():
    excel_path = entry_excel.get()
    sheet_names = entry_sheets.get().split(',')
    ring_number = entry_ring_number.get()
    output_dir = entry_output.get()

    try:
        # Load each sheet as a dataframe, add 'Ring_Number' column, sort columns, and save as CSV
        for sheet_name in sheet_names:
            sheet_name = sheet_name.strip()  # Remove leading/trailing whitespaces
            df = pd.read_excel(excel_path, sheet_name=sheet_name)

            # Add the 'Ring_Number' column
            df['Ring_Number'] = ring_number

            # Sort columns
            df = df.sort_index(axis=1)

            # Save to CSV
            output_path = os.path.join(output_dir, f"{sheet_name} {ring_number}.csv")
            df.to_csv(output_path, index=False)

        messagebox.showinfo("Success", "Successfully sorted and saved!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()

tk.Label(root, text="Excel File Path").grid(row=0)
tk.Label(root, text="Sheet Names (comma-separated)").grid(row=1)
tk.Label(root, text="Ring Number").grid(row=2)
tk.Label(root, text="Output Directory").grid(row=3)

entry_excel = tk.Entry(root)
entry_sheets = tk.Entry(root)
entry_ring_number = tk.Entry(root)
entry_output = tk.Entry(root)

entry_excel.grid(row=0, column=1)
entry_sheets.grid(row=1, column=1)
entry_ring_number.grid(row=2, column=1)
entry_output.grid(row=3, column=1)

tk.Button(root, text='Select Excel File', command=select_file).grid(row=0, column=2)
tk.Button(root, text='Select Output Directory', command=select_output_dir).grid(row=3, column=2)
tk.Button(root, text='Sort and Save', command=sort_and_save).grid(row=4, column=1, columnspan=2)

root.mainloop()

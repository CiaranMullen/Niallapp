import subprocess
import sys
import importlib

# Attempt to install missing packages at runtime
def install_dependencies():
    packages = ['pandas', 'numpy', 'openpyxl']
    for package in packages:
        if importlib.util.find_spec(package) is None:
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
                print(f"Installed {package}")
            except Exception as e:
                print(f"Failed to install {package}: {e}")

install_dependencies()

import pandas as pd
import numpy as np
import random
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def select_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
    if not filename:
        messagebox.showwarning("No file selected", "Please select a file.")
        return

    try:
        # Read file with fallback for CSV encoding
        if filename.endswith('.xlsx'):
            df = pd.read_excel(filename)
        elif filename.endswith('.csv'):
            try:
                df = pd.read_csv(filename, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(filename, encoding='latin1')
        else:
            messagebox.showerror("Invalid file", "Only .xlsx or .csv files are allowed.")
            return

        # Check if 'SN' column exists
        if 'SN' not in df.columns:
            messagebox.showerror("Missing Column", "The selected file does not contain an 'SN' column.")
            return

        # Extract unique SN values
        unique_values = df['SN'].dropna().unique()
        if len(unique_values) == 0:
            messagebox.showwarning("No SN Values", "No unique SN values found.")
            return

        random_sn = random.choice(unique_values)
        output_label.config(text=f"Random unique SN: {random_sn}")

        # Filter rows where SN equals random_sn
        filtered_rows = df[df['SN'] == random_sn]
        output_textbox.delete("1.0", tk.END)
        output_textbox.insert(tk.END, f"Rows where SN = {random_sn}:\n{filtered_rows}")

        # Ask user where to save
        save_filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_filename:
            if os.path.exists(save_filename):
                confirm = messagebox.askyesno("Overwrite File?", f"The file {save_filename} already exists. Overwrite?")
                if not confirm:
                    return

            filtered_rows.to_excel(save_filename, index=False)
            messagebox.showinfo("Success", f"Filtered rows have been saved to:\n{save_filename}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

# GUI Setup
root = tk.Tk()
root.title("Random SN Selector")
root.geometry("600x400")

select_button = tk.Button(root, text="Select Excel or CSV File", command=select_file)
select_button.pack(pady=10)

output_label = tk.Label(root, text="")
output_label.pack()

output_textbox = tk.Text(root, height=15, width=70)
output_textbox.pack(pady=10)

root.mainloop()

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def combine_excels():
    folder_path = folder_var.get()
    output_path = output_var.get()

    if not folder_path or not output_path:
        messagebox.showerror("Error", "Please select both input folder and output file.")
        return

    all_data = []

    try:
        for file in os.listdir(folder_path):
            if file.endswith(".xlsx") and not file.startswith("~$"):
                file_path = os.path.join(folder_path, file)
                print(f"Reading: {file}")

                df = pd.read_excel(
                    file_path,
                    header=None,
                    engine="openpyxl"
                )

                df = df.dropna(how="all")

                if df.empty:
                    continue

                df["Source_File"] = file
                all_data.append(df)

        if not all_data:
            messagebox.showwarning("No Data", "No usable Excel data found.")
            return

        combined = pd.concat(all_data, ignore_index=True)
        combined.to_excel(output_path, index=False, header=False)

        messagebox.showinfo("Success", "Excel files combined successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_var.set(folder_selected)


def browse_output():
    file_selected = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    output_var.set(file_selected)


# ===== GUI SETUP =====
root = tk.Tk()
root.title("Excel Combiner")
root.geometry("500x200")

folder_var = tk.StringVar()
output_var = tk.StringVar()

tk.Label(root, text="Input Folder:").pack(pady=5)
tk.Entry(root, textvariable=folder_var, width=60).pack()
tk.Button(root, text="Browse Folder", command=browse_folder).pack(pady=5)

tk.Label(root, text="Output File:").pack(pady=5)
tk.Entry(root, textvariable=output_var, width=60).pack()
tk.Button(root, text="Save As", command=browse_output).pack(pady=5)

tk.Button(root, text="Combine Excel Files", command=combine_excels, bg="lightgreen").pack(pady=15)

root.mainloop()

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def browse_file(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[("All Files", ".*")])
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

def compare_files():
    master_path = master_entry.get()
    slave_path = slave_entry.get()
    key_column = key_entry.get().strip()

    if not master_path or not slave_path or not key_column:
        messagebox.showwarning("Input Error", "Please select both files and enter a key column.")
        return

    try:
        master_df = pd.read_csv(master_path)
        slave_df = pd.read_excel(slave_path)

        key_cols = [k.strip() for k in key_column.split(',')]
        for col in key_cols:
            if col not in master_df.columns or col not in slave_df.columns:
                messagebox.showerror("Error", f"Key column '{col}' not found in both files.")
                return

        master_df.set_index(key_cols, inplace=True)
        slave_df.set_index(key_cols, inplace=True)

        combined = master_df.join(slave_df, lsuffix='_master', rsuffix='_slave', how='outer', sort=False)

        differences = []
        diff_data = []

        for idx, row in combined.iterrows():
            for col in master_df.columns:
                val_master = row.get(f"{col}_master")
                val_slave = row.get(f"{col}_slave")
                if pd.isna(val_master) and pd.isna(val_slave):
                    continue
                if val_master != val_slave:
                    idx_list = list(idx) if isinstance(idx, tuple) else [idx]
                    diff_data.append(idx_list + [col, val_master, val_slave])
                    differences.append(f"[{', '.join(map(str, idx_list))}], Column '{col}': Master='{val_master}' | Slave='{val_slave}'")

        result_text.delete("1.0", tk.END)
        if differences:
            result_text.insert(tk.END, "\n".join(differences))
        else:
            result_text.insert(tk.END, "No differences found.")

        global export_df
        export_df = pd.DataFrame(diff_data, columns=key_cols + ["column", "old value", "new value"])

    except Exception as e:
        messagebox.showerror("Error", str(e))

def export_to_excel():
    global export_df
    if export_df is None or export_df.empty:
        messagebox.showinfo("No Data", "No differences to export.")
        return

    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel Files", "*.xlsx")])
    if filepath:
        export_df.to_excel(filepath, index=False)
        messagebox.showinfo("Exported", f"Differences exported to:\n{filepath}")

# GUI Setup
root = tk.Tk()
root.title("CSV vs Excel Comparator")

export_df = None

tk.Label(root, text="Master CSV File:").grid(row=0, column=0, sticky="e")
master_entry = tk.Entry(root, width=50)
master_entry.grid(row=0, column=1)
tk.Button(root, text="Browse", command=lambda: browse_file(master_entry)).grid(row=0, column=2)

tk.Label(root, text="Slave Excel File:").grid(row=1, column=0, sticky="e")
slave_entry = tk.Entry(root, width=50)
slave_entry.grid(row=1, column=1)
tk.Button(root, text="Browse", command=lambda: browse_file(slave_entry)).grid(row=1, column=2)

tk.Label(root, text="Key Columns (comma-separated):").grid(row=2, column=0, sticky="e")
key_entry = tk.Entry(root, width=30)
key_entry.grid(row=2, column=1, sticky="w")

tk.Button(root, text="Compare Files", command=compare_files).grid(row=3, column=1, pady=10, sticky="w")
tk.Button(root, text="Export to Excel", command=export_to_excel).grid(row=3, column=2, pady=10)

result_text = tk.Text(root, wrap=tk.WORD, width=100, height=25)
result_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10)
scrollbar = tk.Scrollbar(root, command=result_text.yview)
scrollbar.grid(row=4, column=3, sticky='ns')
result_text['yscrollcommand'] = scrollbar.set

root.mainloop()

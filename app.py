import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from excel_to_ms_sql import get_config, get_data_info, create_table, insert_data

class FileChooserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to MS SQL")
        
        self.file_path_var = tk.StringVar()
        
        self.create_widgets()
        
    def create_widgets(self):
        file_path_label = tk.Label(self.root, text="Chosen file:")
        file_path_label.pack(pady=5)
        
        self.file_path_entry = tk.Entry(self.root, textvariable=self.file_path_var, width=50)
        self.file_path_entry.pack(pady=5)
        
        choose_file_button = tk.Button(self.root, text="Choose File", command=self.choose_file)
        choose_file_button.pack(pady=5)
        
        column_info_label = tk.Label(self.root, text="Column Names and Data Types:")
        column_info_label.pack(pady=5)
        
        self.column_info_text = tk.Text(self.root, width=50, height=20)
        self.column_info_text.pack(pady=5)
        
    def choose_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.display_column_info(file_path)
        
    def display_column_info(self, file_path):
        try:
            if file_path.endswith(('.xls', '.xlsx')):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)
            
            column_info = ""
            for col in df.columns:
                column_info += f"{col}: {df[col].dtype}\n"
            self.column_info_text.delete("1.0", tk.END)
            self.column_info_text.insert(tk.END, column_info)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileChooserApp(root)
    root.mainloop()

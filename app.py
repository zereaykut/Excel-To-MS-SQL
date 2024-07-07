import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from excel_to_ms_sql import get_config, get_data_info, create_table, insert_data

class FileChooserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to MS SQL")
        
        self.file_path_var = tk.StringVar()

        self.config = get_config()

        self.create_widgets()
        
        self.retrieve_database()
        self.retrieve_table()
        
    def create_widgets(self):
        database_entry = tk.Entry(self.root, width=50)
        database_entry.pack(padx=20, pady=20)
        
        database_button = tk.Button(root, text="Submit Database", command=self.retrieve_database)
        database_button.pack(pady=10)

        table_entry = tk.Entry(self.root, width=50)
        table_entry.pack(padx=20, pady=20)
        
        table_button = tk.Button(root, text="Submit Table", command=self.retrieve_table)
        table_button.pack(pady=10)

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

        create_table_button = tk.Button(self.root, text="Create Table", command=create_table(
                                                                                        df_info,
                                                                                        self.table_name,
                                                                                        self.config["server_ip"],
                                                                                        self.database,
                                                                                        self.config["username"],
                                                                                        self.config["password"],
                                                                                        self.config["default_data_types"]
                                                                                    ))
        create_table_button.pack(pady=5)

        insert_data_button = tk.Button(self.root, text="Insert Data", command=self.choose_file)
        insert_data_button.pack(pady=5)
    
    def retrieve_database(self):
        self.database = database_entry.get()
    
    def retrieve_table(self):
        self.table_name = table_entry.get()

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
                self.df = pd.read_excel(file_path)
            else:
                self.df = pd.read_csv(file_path)
            
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

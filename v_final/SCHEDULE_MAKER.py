import tkinter as tk
from tkinter import filedialog
import pandas as pd
from function import *

class ScheduleMakerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SCHEDULE MAKER")
        self.root.geometry("700x300")

        self.file_path = None  # instance variable instead of global

        # UI components
        self.label = tk.Label(root, text="Select a file:")
        self.label.pack()

        self.browse_button = tk.Button(root, text="Browse", command=self.browse_file)
        self.browse_button.pack()

        self.log_text = tk.Text(root, height=10, width=60)
        self.log_text.pack()

        self.process_button = tk.Button(root, text="Process File", command=self.process_file)
        self.process_button.pack()

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("All files", "*.*")])
        if self.file_path:
            self.log(f"File selected: {self.file_path}")
        else:
            self.log("No file selected.")

    def process_file(self):
        if not self.file_path:
            self.log("Please select a file first.")
            return

        try:
            df = pd.read_csv(self.file_path)
            online_df,inperson_df = remove_cols(df)
            total_online_students_col_len = create_new_schedule(online_df,inperson_df)
            colour_cells(total_online_students_col_len)
            remove_empty_sheets_and_rows()
            remove_specific_rows()
            consolidate()
            
            self.log(f"Processed file saved!")
        except Exception as e:
            self.log(f"Error: {e}")

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleMakerApp(root)
    root.mainloop()

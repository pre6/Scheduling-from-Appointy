import tkinter as tk
from tkinter import filedialog
import pandas as pd
from function import *

import os


class ScheduleMakerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Schedule Maker")
        self.root.geometry("750x450")
        self.root.resizable(False, False)

        self.file_path = None
        self.template_path = None
        self.template_highschool_path = None

     # ----- File Select Frame -----
        file_frame = tk.Frame(root, padx=10, pady=10)
        file_frame.pack(fill='x')

        tk.Label(file_frame, text="Schedule CSV:", width=20, anchor="w").grid(row=0, column=0, sticky="w")
        self.file_label = tk.Label(file_frame, text="No file selected", width=50, anchor="w", fg="gray")
        self.file_label.grid(row=0, column=1, padx=5)
        tk.Button(file_frame, text="Browse", command=self.browse_file, width=10).grid(row=0, column=2)

        tk.Label(file_frame, text="Template File:", width=20, anchor="w").grid(row=1, column=0, sticky="w", pady=5)
        self.template_label = tk.Label(file_frame, text="No file selected", width=50, anchor="w", fg="gray")
        self.template_label.grid(row=1, column=1, padx=5)
        tk.Button(file_frame, text="Browse", command=self.browse_template, width=10).grid(row=1, column=2)

        tk.Label(file_frame, text="Highschool Template:", width=20, anchor="w").grid(row=2, column=0, sticky="w")
        self.template_highschool_label = tk.Label(file_frame, text="No file selected", width=50, anchor="w", fg="gray")
        self.template_highschool_label.grid(row=2, column=1, padx=5)
        tk.Button(file_frame, text="Browse", command=self.browse_template_highschool, width=10).grid(row=2, column=2)

        # ----- Log Output -----
        log_frame = tk.Frame(root, padx=10, pady=10)
        log_frame.pack(fill='x')

        tk.Label(log_frame, text="Log Output:").pack(anchor="w")
        self.log_text = tk.Text(log_frame, height=12, width=90, bg="#f9f9f9", relief="solid", borderwidth=1)
        self.log_text.pack()

        # ----- Process Button -----
        button_frame = tk.Frame(root, pady=10)
        button_frame.pack()
        self.process_button = tk.Button(button_frame, text="Process File", command=self.process_file, width=30, height=2, bg="#4CAF50", fg="white", font=('Helvetica', 10, 'bold'))
        self.process_button.pack()

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if self.file_path:
            self.file_label.config(text=self.file_path, fg="black")
            self.log(f"Schedule file selected: {self.file_path}")
        else:
            self.log("No schedule file selected.")

    def browse_template(self):
        self.template_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.template_path:
            self.template_label.config(text=self.template_path, fg="black")
            self.log(f"Template file selected: {self.template_path}")
        else:
            self.log("No template file selected.")

    def browse_template_highschool(self):
        self.template_highschool_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.template_highschool_path:
            self.template_highschool_label.config(text=self.template_highschool_path, fg="black")
            self.log(f"Highschool Template file selected: {self.template_highschool_path}")
        else:
            self.log("No highschool template file selected.")


    def process_file(self):
        if not self.file_path or not self.template_path or not self.template_highschool_path:
            self.log("❌ Please make sure all files are selected.")
            return
        
        # Ensure 'Schedule' folder exists
        output_dir = "Schedule"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            self.log("Created 'Schedule' folder.")

        try:
            df = pd.read_csv(self.file_path)
            earliest_date = df['Appointment_Date'].min()
            output_path = os.path.join("Schedule", str(earliest_date) + ".xlsx")

            online_df, inperson_df = remove_cols(df)
            total_online_students_col_len = create_new_schedule(online_df, inperson_df, earliest_date)
            colour_cells(total_online_students_col_len, earliest_date)
            remove_empty_sheets_and_rows(earliest_date)
            remove_specific_rows(earliest_date)
            consolidate(earliest_date)

            copy_the_template(self.template_path, earliest_date)
            copy_highschool(self.template_highschool_path, earliest_date)

            self.log("✅ Processing complete. Files saved.")
        except Exception as e:
            self.log(f"❌ Error: {e}")

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log



if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleMakerApp(root)
    root.mainloop()

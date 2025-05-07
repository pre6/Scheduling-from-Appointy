from utils import *
import tkinter as tk
from tkinter import filedialog
from tkinter import Text, filedialog




# Define a variable to store the selected file path
selected_file_path = ""

# Function to browse for a file
def browse_file():
    global selected_file_path
    selected_file_path = filedialog.askopenfilename()
    log_text.insert(tk.END, f"Selected file: {selected_file_path}\n")

# Define a function to perform the entire process

missing_students = []


def process_file():
    global missing_students
    today = datetime.datetime.today()

    # Calculate the date of the next Monday
    days_until_monday = (7 - today.weekday()) % 7
    next_monday = today + datetime.timedelta(days=days_until_monday)
    '''This is a string of the date on monday'''
    monday_date = next_monday.strftime('%Y-%m-%d')


    

    global selected_file_path
    if not selected_file_path:
        log_text.insert(tk.END, "Please select a file first.\n")
        return
    


    log_text.insert(tk.END, "Getting TXT file location\n")

    dictionary = read_config_file('C:/Users/preet/Documents/GitHub/Scheduling-from-Appointy/system_files/text_file.txt')


    appoint_path = dictionary['appointments folder']
    print(appoint_path)
    student_delete_path = dictionary['student list folder']
    print(student_delete_path)
    template_path = dictionary['template_path']
    print(template_path)
    

    log_text.insert(tk.END, "Done!\n")



    
    log_text.insert(tk.END, "Cleaning CSV\n")

    cleaning_path = student_delete_path+monday_date+'.xlsx'
    create_appointments_path = appoint_path+monday_date+'.xlsx'

    convert(selected_file_path,cleaning_path)
    date_time_convert(cleaning_path)

    delete_bad_columns(cleaning_path)
    fix_intake(cleaning_path)
    no_student_name(cleaning_path)
    home_online(cleaning_path)

    log_text.insert(tk.END, "Done!\n")



    
    log_text.insert(tk.END, "Creating Excel\n")

    copy_template(create_appointments_path,template_path)
    create_schedule(cleaning_path,create_appointments_path,missing_students)

    log_text.insert(tk.END, "Done!\n")



    
    


    for student_info in missing_students:
        result_text.insert(tk.END, student_info + '\n')


    

if __name__ == "__main__":
    missing_students = []

    # Create the main window
    root = tk.Tk()
    root.title("SCHEDULE MAKER")
    root.geometry("700x450")

    label = tk.Label(root, text="Select a file:")
    label.pack()
    # Create a "Browse" button
    browse_button = tk.Button(root, text="Browse", command=browse_file)
    browse_button.pack()

    # Create a Text widget for displaying log messages
    log_text = tk.Text(root, height=10, width=60)
    log_text.pack()

    # Create another Text widget for displaying the result or additional information
    result_text = tk.Text(root, height=10, width=60)
    result_text.pack()

    # Create a button to trigger the entire process
    process_button = tk.Button(root, text="Process File", command=process_file)
    process_button.pack()



    # Start the GUI event loop
    root.mainloop()


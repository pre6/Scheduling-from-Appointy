import pandas as pd
import re




# Replace with your actual file path
df = pd.read_csv("appointmentsReport.csv")

# Drop unwanted columns
df.drop(columns=['Booking_Date', 'Booking_Time', 'Customer_Email', 'Customer_Address', 'Customer_Contact', 'Customer_Timezone', 'Booked_By', 'Service_Duration_mins', 'Amount', 'Staff_Name', 'Staff_Email', 'Resource', 'Customer_Remark', 'Appointment_Status', 'Arrival_Status', 'Admin_Note', 'Payment_Info', 'AddOn'], inplace=True)




# fix strings 
df['Intake_Form'] = df['Intake_Form'].str.replace(r'Intake form:\n|Student Name:|Student #2: ,\n|Student #2:|Student #3:|Student #2 Name \(if applicable\): ,|Student #3 Name \(if applicable\): |,\n|\n', '', regex=True).str.strip()
print(df.head())

df['Intake_Form'] = df.apply(lambda row: row['Customer_Name'] if row['Intake_Form'] == '-' else row['Intake_Form'], axis=1)

# print(df['Intake_Form'].tolist())
print(df.columns)

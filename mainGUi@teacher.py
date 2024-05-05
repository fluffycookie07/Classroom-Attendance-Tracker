import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import pandas as pd

class AttendanceTrackerApp:
    def __init__(self, root, subjects, total_students):
        # Initialize the GUI
        self.root = root
        self.root.title("Attendance Tracker")
        self.root.geometry("1280x720")
        self.subjects = subjects
        self.total_students = total_students
        self.current_subject = None
        self.create_main_frame()
        
    def create_main_frame(self):
        # Create main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Create buttons to switch between views
        view_button = ttk.Button(main_frame, text="View Attendance", command=self.show_view_tab)
        view_button.grid(row=0, column=0, padx=10, pady=10)
        
        enter_button = ttk.Button(main_frame, text="Enter Attendance", command=self.show_enter_tab)
        enter_button.grid(row=0, column=1, padx=10, pady=10)
        
        add_student_button = ttk.Button(main_frame, text="Add Student", command=self.show_add_student_tab)
        add_student_button.grid(row=0, column=2, padx=10, pady=10)
        
        overall_button = ttk.Button(main_frame, text="Overall Attendance", command=self.calculate_overall_attendance)
        overall_button.grid(row=0, column=3, padx=10, pady=10)
        
        self.view_frame = ttk.Frame(main_frame)
        self.enter_frame = ttk.Frame(main_frame)
        self.add_student_frame = ttk.Frame(main_frame)
        
        self.create_view_tab()
        self.create_enter_tab()
        self.create_add_student_tab()
        
    def create_view_tab(self):
        # Create the UI elements for viewing attendance
        subject_label = ttk.Label(self.view_frame, text="Select Subject:")
        subject_label.grid(row=0, column=0, padx=(0, 10))
        self.subject_combobox_view = ttk.Combobox(self.view_frame, values=self.subjects)
        self.subject_combobox_view.grid(row=0, column=1)
        date_label_view = ttk.Label(self.view_frame, text="Select Date:")
        date_label_view.grid(row=0, column=2, padx=(10, 0))
        self.date_entry_view = DateEntry(self.view_frame, date_pattern='dd-mm-yyyy')
        self.date_entry_view.grid(row=0, column=3)
        view_attendance_button = ttk.Button(self.view_frame, text="View Attendance", command=self.display_attendance)
        view_attendance_button.grid(row=0, column=4, padx=(10, 0))
        calculate_total_button = ttk.Button(self.view_frame, text="Calculate Total", command=self.calculate_total)
        calculate_total_button.grid(row=0, column=5, padx=(10, 0))
        self.attendance_tree = ttk.Treeview(self.view_frame, columns=("Roll No.", "Name", "Status"))
        self.attendance_tree.heading("#0", text="")
        self.attendance_tree.heading("Roll No.", text="Roll No.")
        self.attendance_tree.heading("Name", text="Name")
        self.attendance_tree.heading("Status", text="Status")
        self.attendance_tree.grid(row=1, column=0, columnspan=6, padx=10, pady=10)
        self.view_frame.grid_columnconfigure(0, weight=1)
        
    def create_enter_tab(self):
        # Create the UI elements for entering attendance
        subject_label_enter = ttk.Label(self.enter_frame, text="Subject:")
        subject_label_enter.grid(row=0, column=0, sticky="w")
        self.subject_entry_enter = ttk.Entry(self.enter_frame)
        self.subject_entry_enter.grid(row=0, column=1)
        date_label_enter = ttk.Label(self.enter_frame, text="Date:")
        date_label_enter.grid(row=1, column=0, sticky="w")
        self.date_entry_enter = DateEntry(self.enter_frame, date_pattern='dd-mm-yyyy')
        self.date_entry_enter.grid(row=1, column=1)
        absent_label_enter = ttk.Label(self.enter_frame, text="Absent Roll No.:")
        absent_label_enter.grid(row=2, column=0, sticky="w")
        self.absent_entry_enter = ttk.Entry(self.enter_frame)
        self.absent_entry_enter.grid(row=2, column=1)
        absent_help_label_enter = ttk.Label(self.enter_frame, text="Enter comma-separated roll numbers")
        absent_help_label_enter.grid(row=3, column=1, sticky="w")
        submit_button_enter = ttk.Button(self.enter_frame, text="Submit", command=self.submit_attendance)
        submit_button_enter.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        self.enter_frame.grid_columnconfigure(0, weight=1)
        
    def create_add_student_tab(self):
        # Create the UI elements for adding students
        name_label = ttk.Label(self.add_student_frame, text="Name:")
        name_label.grid(row=0, column=0, sticky="w")
        self.name_entry = ttk.Entry(self.add_student_frame)
        self.name_entry.grid(row=0, column=1)
        add_button = ttk.Button(self.add_student_frame, text="Add Student", command=self.add_student)
        add_button.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        self.add_student_frame.grid_columnconfigure(0, weight=1)
        
    def show_view_tab(self):
        # Switch to the view attendance tab
        self.enter_frame.grid_forget()
        self.add_student_frame.grid_forget()
        self.view_frame.grid(row=1, column=0, sticky="nsew")
        self.current_subject = None
        
    def show_enter_tab(self):
        # Switch to the enter attendance tab
        self.view_frame.grid_forget()
        self.add_student_frame.grid_forget()
        self.enter_frame.grid(row=1, column=0, sticky="nsew")
        self.current_subject = None
        
    def show_add_student_tab(self):
        # Switch to the add student tab
        self.view_frame.grid_forget()
        self.enter_frame.grid_forget()
        self.add_student_frame.grid(row=1, column=0, sticky="nsew")
        
    def display_attendance(self):
        # Function to display attendance
        subject = self.subject_combobox_view.get()
        date = self.date_entry_view.get()
        if not subject or not date:
            return
        try:
            df = pd.read_excel("AttendanceData.xlsx", sheet_name=f"{subject}_Attendance")
            self.attendance_tree.delete(*self.attendance_tree.get_children())
            if df.empty:
                print(f"Attendance data for {subject} is empty.")
                return
            for idx, row in df.iterrows():
                status = "Present" if row[date] == 1 else "Absent"
                self.attendance_tree.insert("", "end", values=(row['Roll No.'], row['Name'], status))
        except FileNotFoundError:
            print(f"Attendance data for {subject} not found.")
        except Exception as e:
            print(f"An error occurred: {e}")
        
    def submit_attendance(self):
        # Function to submit attendance
        subject = self.subject_entry_enter.get()
        date = self.date_entry_enter.get()
        absent = self.absent_entry_enter.get().split(',')
        try:
            try:
                with pd.ExcelFile("AttendanceData.xlsx") as xls:
                    df = pd.read_excel(xls, sheet_name=f"{subject}_Attendance")
            except FileNotFoundError:
                df = pd.DataFrame(columns=['Roll No.', 'Name'])
            if df.empty:
                df['Roll No.'] = self.total_students.keys()
                df['Name'] = self.total_students.values()
            if date not in df.columns:
                df[date] = ""
            absent_list = [int(x.strip()) for x in absent]
            for roll, student in self.total_students.items():
                if roll not in absent_list:
                    df.loc[df['Roll No.'] == roll, date] = 1
                else:
                    df.loc[df['Roll No.'] == roll, date] = 0
            with pd.ExcelWriter("AttendanceData.xlsx", mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=f"{subject}_Attendance", index=False)
            print(f"\nAttendance for {subject} on {date} has been updated.")
        except Exception as e:
            print(f"An error occurred: {e}")
            
    def add_student(self):
        # Function to add a new student
        name = self.name_entry.get().strip()
        if name:
            roll_no = max(self.total_students.keys()) + 1
            self.total_students[roll_no] = name
            print(f"Student '{name}' added with Roll No. {roll_no}.")
            self.name_entry.delete(0, tk.END)
        else:
            print("Please enter a valid name.")
    
    def calculate_total(self):
        # Function to calculate total attendance
        subject = self.subject_combobox_view.get()
        if not subject:
            return
        try:
            with pd.ExcelFile("AttendanceData.xlsx") as xls:
                df = pd.read_excel(xls, sheet_name=f"{subject}_Attendance")
            if df.empty:
                print(f"Attendance data for {subject} is empty.")
                return
            
            total_days = len(df.columns) - 2 
            
            detained_students = []
            
            for idx, row in df.iterrows():
                total_attendance = sum(row[2:])
                total_percentage = (total_attendance / total_days) * 100
                if total_percentage < 75:
                    detained_students.append((subject, row['Roll No.']))
                df.loc[idx, 'Total Attendance'] = total_attendance
                df.loc[idx, 'Percentage'] = total_percentage
            
            with pd.ExcelWriter("AttendanceData.xlsx", mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=f"{subject}_Attendance", index=False)

            if detained_students:
                print("Detained Students:")
                for detained_student in detained_students:
                    detained_subject, detained_roll = detained_student
                    detained_name = self.total_students.get(detained_roll)
                    if detained_name:
                        print(f"{detained_name} in {detained_subject}")
                    else:
                        print(f"Roll No. {detained_roll} in {detained_subject} (Name not found)")
            else:
                print("No students detained.")
        
        except FileNotFoundError:
            print(f"Attendance data for {subject} not found.")
        except Exception as e:
            print(f"An error occurred: {e}")
            
    def calculate_overall_attendance(self):
        # Function to calculate overall attendance
        try:
            # Define file paths and sheet names
            file_path = "AttendanceData.xlsx"
            destination_sheet_name = "OverallSubjects_Attendance"  # Destination sheet

            # Read the existing DataFrame from the Excel file
            df_destination = pd.read_excel(file_path, sheet_name=destination_sheet_name)

            FDS_df_source = pd.read_excel(file_path, sheet_name="FDS_Attendance")
            PSP_df_source = pd.read_excel(file_path, sheet_name="PSP_Attendance")
            Maths_df_source = pd.read_excel(file_path, sheet_name="Maths_Attendance")
            Physics_df_source = pd.read_excel(file_path, sheet_name="Physics_Attendance")

            # Define new data to add to df_destination
            new_data = {
                "Roll No.": FDS_df_source["Roll No."],
                "Name": FDS_df_source['Name'],
                "FDS%": FDS_df_source['Percentage'],
                "PSP%": PSP_df_source['Percentage'],
                "Maths%": Maths_df_source['Percentage'],
                "Physics%": Physics_df_source['Percentage']
            }

            # Create a new DataFrame from the new_data dictionary
            df_new_data = pd.DataFrame(new_data)

            # Append the new data to df_destination using concat
            df_destination = pd.concat([df_destination, df_new_data], ignore_index=True)

            # Calculate the average attendance and add it as a new column
            df_destination['Average Attendance'] = df_destination[['FDS%', 'PSP%', 'Maths%', 'Physics%']].mean(axis=1)

            # Filter detained students (average attendance < 75)
            detained_students = df_destination[df_destination['Average Attendance'] < 75]

            # Save the updated df_destination to an Excel file
            df_destination.to_excel(file_path, index=False, sheet_name=destination_sheet_name)

            if not detained_students.empty:
                print("Detained Students:")
                for idx, student in detained_students.iterrows():
                    print(f"{student['Roll No.']}: {student['Name']}")
            else:
                print("No students detained.")

            print("Updated data saved to Excel successfully.")
        except Exception as e:
            print(f"An error occurred: {e}")

if __name__ == "__main__":
    subjects = ["FDS", "PSP", "Physics", "Maths"]
    total_students = {1: 'ADHAV SANGHARSH BHAGWAT', 2: 'AHER GANESH SANJAY', 3: 'ANAP SHRADDHA SUNIL', 4: 'ANJALI CHANDRABHAN SONAWANE', 5: 'AUTI ARTI BALASAHEB', 6: 'BAGUL SURAJ SIDDHARTH', 7: 'BAJAJ MANJIT AJIT', 8: 'BANKAR SARTHAK SANDIP', 9: 'BANKAR SMITRAJ DINKAR', 10: 'BANSODE NIKITA SANJAY',
                     11: 'BHAKARE TANISHKA SHARAD', 12: 'BHANDARE ROHIT RAJENDRA', 13: 'BHAWAR SARTHAK VIJAY', 14: 'BHGINGARDIVE SAMARTH DEEPAK', 15: 'BHORUNDE PRATIKA DILIP', 16: 'BHOSALE SNEHAL ANNASAHEB', 17: 'BHOYE POOJA KAMALAKAR', 18: 'BORDE SARTHAK RAJENDRA', 19: 'BORSE VAISHNAVI SANDESH', 20: 'BORUDE YASH AMBADAS',
                      21: 'CHAUDHARI ABHISHEK ANIL', 22: 'CHAUDHARI TEJAS VIKAS', 23: 'CHAVAN ANIKET PRAKASH', 24: 'CHAVAN SAMRUDDHI PANDIT', 25: 'DABHADE UNNATI VIJAY', 26: 'DANGE SHRADDHA SACHIN', 27: 'DEOKAR ANUSHKA SANTOSH', 28: 'DEOKAR PRANAV BALASAHEB', 29: 'DESHMUKH DHANRAJ MAHENDRA', 30: 'DESHMUKH RITESH NARENDRA', 
                      31: 'DESHMUKH SAYALI VINOD', 32: 'DEVADHE PRATIK PANDURANG', 33: 'DEVKATE RAJENDRA BABASAHEB', 34: 'DHADGE VEDANT SANJAY', 35: 'DIWATE ROSHAN BHAUSAHEB', 36: 'DYAVANE ABHAY SURESH', 37: 'GADHE SUVARNA BABASAHEB', 38: 'GAGARE YASH BHIMRAJ', 39: 'GAIKWAD PRANJALI AJAY', 40: 'GAIKWAD SAIRAM AJIT',
                      41: 'GAIKWAD VAISHNAVI KALYANRAO', 42: 'GAIKWAD VEDANT JITENDRA', 43: 'GALHATE DNYANESHWAR GOKUL', 44: 'GAYKE PRATHAMESH RAVINDRA', 45: 'GIRASE JAYESH KEWALSING', 46: 'GIRI ABHISHEK ANKUSH', 47: 'GUNJAL PALLAVI SANTOSH', 48: 'JADHAV ADITYA NARAYAN', 49: 'JADHAV ISHANT RANJAN', 50: 'JADHAV MANOJ RAMBHAU',
                      51: 'JADHAV PRATIKSHA PRAKASH', 52: 'JADHAV VAIBHAV VILAS', 53: 'JAGADALE KUNAL SUNIL', 54: 'JAGDALE SUPRIYA NANDKISHORE',55: 'JAMDHADE AARTI RAHUL', 56: 'KADAM KRISHNA BAPUSAHEB', 57: 'KADAM OM VIKRAM', 58: 'KADAM RUTUJA ANIL', 59: 'KADAM SNEHAL SHIVAJI',
                      60: 'KADAM VIRAJ KRISHNA', 61: 'KALE MAYUR VIJAY', 62: 'KALE SANIKA DNYANESHWAR', 63: 'KAMODKAR SAKSHI NITIN', 64: 'KARAD MAYUR SUDHAKAR', 65: 'KASAR ABHIJIT BALASAHEB', 66: 'KEKAN SAKSHI KIRAN', 67: 'KEVAL ATHARVA YOGESH', 68: 'KHEDKAR SAKSHI SANTOSH', 69: 'KHILLARE AMAN SUNIL'
    }      

    root = tk.Tk()
    app = AttendanceTrackerApp(root, subjects, total_students)

    root.mainloop()

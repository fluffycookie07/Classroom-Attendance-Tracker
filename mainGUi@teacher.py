import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import pandas as pd

class AttendanceTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Attendance Tracker")
        self.root.geometry("1280x720")
        self.total_students = {
            1: 'Vaishnavi', 2: 'Tanishka', 3: 'Nikita', 4: 'Anjali', 5: 'Manjit',
            6: 'Tejas', 7: 'Abhishek', 8: 'Smitraj', 9: 'Yash', 10: 'Sarthak',
            11: 'Komal', 12: 'Suresh', 13: 'Amit', 14: 'Deepak', 15: 'Riya',
            16: 'Shivani', 17: 'Raj', 18: 'Pooja', 19: 'Alok', 20: 'Neha',
            21: 'Karan', 22: 'Anuradha', 23: 'Vikas', 24: 'Kavita', 25: 'Harish',
            26: 'Geeta', 27: 'Sohan', 28: 'Anushka', 29: 'Vinod', 30: 'Meena', 31: 'Pranjali'
        }
        self.create_main_tab()
        
    def create_main_tab(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill="both", expand=True)
        view_frame = ttk.Frame(notebook)
        enter_frame = ttk.Frame(notebook)
        notebook.add(view_frame, text="View Attendance")
        notebook.add(enter_frame, text="Enter Attendance")
        self.create_view_tab(view_frame)
        self.create_enter_tab(enter_frame)
        
    def create_view_tab(self, view_frame):
        subject_label = ttk.Label(view_frame, text="Select Subject:")
        subject_label.grid(row=0, column=0, padx=(0, 10))
        self.subject_combobox_view = ttk.Combobox(view_frame, values=["FDS", "PSP", "Physics", "Maths"])
        self.subject_combobox_view.grid(row=0, column=1)
        view_attendance_button = ttk.Button(view_frame, text="View Attendance", command=self.display_attendance)
        view_attendance_button.grid(row=0, column=2, padx=(10, 0))
        calculate_button = ttk.Button(view_frame, text="Calculate Total", command=self.calculate_total)
        calculate_button.grid(row=0, column=3, padx=(10, 0))
        self.attendance_tree = ttk.Treeview(view_frame, columns=("Roll No.", "Name", "Date", "Status"))
        self.attendance_tree.heading("#0", text="")
        self.attendance_tree.heading("Roll No.", text="Roll No.")
        self.attendance_tree.heading("Name", text="Name")
        self.attendance_tree.heading("Date", text="Date")
        self.attendance_tree.heading("Status", text="Status")
        self.attendance_tree.grid(row=1, column=0, columnspan=5, padx=10, pady=10)
        
    def create_enter_tab(self, enter_frame):
        subject_label_enter = ttk.Label(enter_frame, text="Subject:")
        subject_label_enter.grid(row=0, column=0, sticky="w")
        self.subject_entry_enter = ttk.Entry(enter_frame)
        self.subject_entry_enter.grid(row=0, column=1)
        date_label_enter = ttk.Label(enter_frame, text="Date:")
        date_label_enter.grid(row=1, column=0, sticky="w")
        self.date_entry_enter = DateEntry(enter_frame, date_pattern='dd-mm-yyyy')
        self.date_entry_enter.grid(row=1, column=1)
        absent_label_enter = ttk.Label(enter_frame, text="Absent Roll No.:")
        absent_label_enter.grid(row=2, column=0, sticky="w")
        self.absent_entry_enter = ttk.Entry(enter_frame)
        self.absent_entry_enter.grid(row=2, column=1)
        absent_help_label_enter = ttk.Label(enter_frame, text="Enter comma-separated roll numbers")
        absent_help_label_enter.grid(row=3, column=1, sticky="w")
        submit_button_enter = ttk.Button(enter_frame, text="Submit", command=self.submit_attendance)
        submit_button_enter.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        
    def display_attendance(self):
        subject = self.subject_combobox_view.get()
        if not subject:
            return
        file_name = f"{subject}_attendance.xlsx"
        try:
            df = pd.read_excel(file_name)
            self.attendance_tree.delete(*self.attendance_tree.get_children())
            if df.empty:
                print(f"Attendance file for {subject} is empty.")
                return
            for idx, row in df.iterrows():
                status = "1" if row[self.date_entry_enter.get()] == 1 else "0"
                self.attendance_tree.insert("", "end", values=(row['Roll No.'], row['Name'], self.date_entry_enter.get(), status))
        except FileNotFoundError:
            print(f"Attendance file for {subject} not found.")
        except Exception as e:
            print(f"An error occurred: {e}")
        
    def submit_attendance(self):
        subject = self.subject_entry_enter.get()
        date = self.date_entry_enter.get()
        absent = self.absent_entry_enter.get().split(',')
        file_name = f"{subject}_attendance.xlsx"
        try:
            try:
                df = pd.read_excel(file_name)
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
            df.to_excel(file_name, index=False)
            print(f"\nAttendance for {subject} on {date} has been updated.")
        except Exception as e:
            print(f"An error occurred: {e}")
        
    def calculate_total(self):
        subject = self.subject_combobox_view.get()
        if not subject:
            return
        file_name = f"{subject}_attendance.xlsx"
        try:
            df = pd.read_excel(file_name)
            if df.empty:
                print(f"Attendance file for {subject} is empty.")
                return
            total_attendance = df.drop(['Roll No.', 'Name'], axis=1).sum().sum()
            total_days = len(df.columns) - 2 
            total_percentage = (total_attendance / (len(self.total_students) * total_days)) * 100 
            detained_students = [roll for roll in df.columns[2:] if (df[roll].sum() / total_days) * 100 < 75]
            detained_students_names = [self.total_students[int(roll)] for roll in detained_students]
            detained = ", ".join(detained_students_names) if detained_students_names else "None"
            print(f"Total Attendance: {total_attendance}")
            print(f"Total Percentage: {total_percentage:.2f}%")
            print(f"Detained Students: {detained}")
            
            # Add total count column to the excel sheet
            df['Total Attendance'] = df.drop(['Roll No.', 'Name'], axis=1).sum(axis=1)
            df.to_excel(file_name, index=False)
            
        except FileNotFoundError:
            print(f"Attendance file for {subject} not found.")
        except Exception as e:
            print(f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceTrackerApp(root)
    root.mainloop()

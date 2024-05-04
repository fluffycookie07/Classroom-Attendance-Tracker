import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import pandas as pd

class AttendanceTrackerApp:
    def __init__(self, root, subjects, total_students):
        self.root = root
        self.root.title("Attendance Tracker")
        self.root.geometry("1280x720")
        self.subjects = subjects
        self.total_students = total_students
        self.current_subject = None
        self.create_main_frame()
        
    def create_main_frame(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Create buttons to switch between views
        view_button = ttk.Button(main_frame, text="View Attendance", command=self.show_view_tab)
        view_button.grid(row=0, column=0, padx=10, pady=10)
        
        enter_button = ttk.Button(main_frame, text="Enter Attendance", command=self.show_enter_tab)
        enter_button.grid(row=0, column=1, padx=10, pady=10)
        
        self.view_frame = ttk.Frame(main_frame)
        self.enter_frame = ttk.Frame(main_frame)
        
        self.create_view_tab()
        self.create_enter_tab()
        
    def create_view_tab(self):
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
        self.attendance_tree = ttk.Treeview(self.view_frame, columns=("Roll No.", "Name", "Status"))
        self.attendance_tree.heading("#0", text="")
        self.attendance_tree.heading("Roll No.", text="Roll No.")
        self.attendance_tree.heading("Name", text="Name")
        self.attendance_tree.heading("Status", text="Status")
        self.attendance_tree.grid(row=1, column=0, columnspan=5, padx=10, pady=10)
        self.view_frame.grid_columnconfigure(0, weight=1)
        
    def create_enter_tab(self):
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
        
    def show_view_tab(self):
        self.enter_frame.grid_forget()
        self.view_frame.grid(row=1, column=0, sticky="nsew")
        self.current_subject = None
        
    def show_enter_tab(self):
        self.view_frame.grid_forget()
        self.enter_frame.grid(row=1, column=0, sticky="nsew")
        self.current_subject = None
        
    def display_attendance(self):
        subject = self.subject_combobox_view.get()
        date = self.date_entry_view.get()
        if not subject or not date:
            return
        try:
            df = pd.read_excel("Attendance Data.xlsx", sheet_name=f"{subject}_Attendance")
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
        subject = self.subject_entry_enter.get()
        date = self.date_entry_enter.get()
        absent = self.absent_entry_enter.get().split(',')
        try:
            try:
                with pd.ExcelFile("Attendance Data.xlsx") as xls:
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
            with pd.ExcelWriter("Attendance Data.xlsx", mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=f"{subject}_Attendance", index=False)
            print(f"\nAttendance for {subject} on {date} has been updated.")
        except Exception as e:
            print(f"An error occurred: {e}")

if __name__ == "__main__":
    subjects = ["FDS", "PSP", "Physics", "Maths"]
    total_students = {
        1: 'Vaishnavi', 2: 'Tanishka', 3: 'Nikita', 4: 'Anjali', 5: 'Manjit',
        6: 'Tejas', 7: 'Abhishek', 8: 'Smitraj', 9: 'Yash', 10: 'Sarthak',
        11: 'Komal', 12: 'Suresh', 13: 'Amit', 14: 'Deepak', 15: 'Riya',
        16: 'Shivani', 17: 'Raj', 18: 'Pooja', 19: 'Alok', 20: 'Neha',
        21: 'Karan', 22: 'Anuradha', 23: 'Vikas', 24: 'Kavita', 25: 'Harish',
        26: 'Geeta', 27: 'Sohan', 28: 'Anushka', 29: 'Vinod', 30: 'Meena', 31: 'Pranjali'
    }
    root = tk.Tk()
    app = AttendanceTrackerApp(root, subjects, total_students)
    root.mainloop()

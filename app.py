import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, StringVar
from concatFiles import ConcatFiles

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("COMPARE")
        self.root.geometry("400x400")
        self.path_protime = r'./Protime'
        self.path_agency = r'./Agency'
        
        # Creating a frame
        self.frame = ttk.Frame(self.root)
        self.frame.pack(pady=20)

        # Label for Protime folder path
        self.label_protime = ttk.Label(self.frame, text=f"Protime Folder: {self.path_protime}", font=("Helvetica", 12))
        self.label_protime.pack(pady=10)

        # Label for Agency folder path
        self.label_agency = ttk.Label(self.frame, text=f"Agency Folder: {self.path_agency}", font=("Helvetica", 12))
        self.label_agency.pack(pady=10)

        # Buttons to select folder
        self.button_protime = ttk.Button(self.frame, text="Select Protime Folder", command=self.select_protime_folder)
        self.button_protime.pack(pady=5)

        self.button_agency = ttk.Button(self.frame, text="Select Agency Folder", command=self.select_agency_folder)
        self.button_agency.pack(pady=5)

        # Dropdown for start and end week selection
        self.label_week = ttk.Label(self.frame, text="Select Start and End Week", font=("Helvetica", 12))
        self.label_week.pack(pady=10)

        # Start week dropdown
        self.start_week = ttk.Combobox(self.frame, values=[f"Week {i}" for i in range(1, 53)], state="readonly")
        self.start_week.set("Start Week")
        self.start_week.pack(pady=5)

        # End week dropdown
        self.end_week = ttk.Combobox(self.frame, values=[f"Week {i}" for i in range(1, 53)], state="readonly")
        self.end_week.set("End Week")
        self.end_week.pack(pady=5)

        # File format option (radio buttons for .xlsx or .csv)
        self.label_format = ttk.Label(self.frame, text="Select Report Format", font=("Helvetica", 12))
        self.label_format.pack(pady=10)

        self.file_format = StringVar(value=".xlsx")
        self.radio_xlsx = ttk.Radiobutton(self.frame, text="Excel (.xlsx)", variable=self.file_format, value=".xlsx")
        self.radio_xlsx.pack(pady=2)

        self.radio_csv = ttk.Radiobutton(self.frame, text="CSV (.csv)", variable=self.file_format, value=".csv")
        self.radio_csv.pack(pady=2)

        # Button to start comparison
        self.compare_button = ttk.Button(self.frame, text="Compare Data", command=self.compare_data)
        self.compare_button.pack(pady=10)

        # Button to save the report
        self.save_button = ttk.Button(self.frame, text="Save Report", command=self.save_report)
        self.save_button.pack(pady=10)

    def select_protime_folder(self):
        folder_selected = filedialog.askdirectory(title="Select Protime Folder")
        if folder_selected:
            self.path_protime = folder_selected
            self.label_protime.config(text=f"Protime Folder: {self.path_protime}")

    def select_agency_folder(self):
        folder_selected = filedialog.askdirectory(title="Select Agency Folder")
        if folder_selected:
            self.path_agency = folder_selected
            self.label_agency.config(text=f"Agency Folder: {self.path_agency}")

    def compare_data(self):
        # Get start and end weeks
        start_week = self.start_week.get()
        end_week = self.end_week.get()

        if "Start Week" in start_week or "End Week" in end_week:
            print("Please select valid start and end weeks.")
            return

        # Load and compare data from both folders
        protime = ConcatFiles(self.path_protime)
        agency = ConcatFiles(self.path_agency)
        
        protime.combine_files()  # Combine files in Protime folder
        agency.combine_files()   # Combine files in Agency folder
        
        # Compare data based on the selected weeks (start_week and end_week)
        print(f"Comparing data from {start_week} to {end_week}...")

        # (Implement comparison logic here based on the selected weeks)

    def save_report(self):
        # Get file format
        file_format = self.file_format.get()

        # File dialog for saving the report
        file_path = filedialog.asksaveasfilename(defaultextension=file_format, filetypes=[(f"{file_format.upper()} files", f"*{file_format}")])

        if not file_path:
            print("No file selected.")
            return

        # Save the comparison report to the specified format
        print(f"Saving report as {file_format} at {file_path}...")

        # (Implement logic to save the report to Excel or CSV)

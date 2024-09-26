# view.py
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import tkinter as tk

class DataView:
    def __init__(self, master, run_callback):
        self.master = master
        self.master.title("Data Comparison Tool")
        self.agency_path = ''
        self.protime_path = ''
        self.run_callback = run_callback  # Accept the callback function
        self.create_widgets()

    def create_widgets(self):
        # Frame for Agency Data Selection
        agency_frame = ttk.Frame(self.master, padding=10)
        agency_frame.pack(fill=X)
        agency_label = ttk.Label(agency_frame, text="Select Agency Data Folder:")
        agency_label.pack(side=LEFT)
        self.agency_path_label = ttk.Label(agency_frame, text="No folder selected", foreground='red')
        self.agency_path_label.pack(side=LEFT, padx=5)
        agency_button = ttk.Button(agency_frame, text="Browse", command=self.select_agency_folder)
        agency_button.pack(side=RIGHT)

        # Frame for Protime Data Selection
        protime_frame = ttk.Frame(self.master, padding=10)
        protime_frame.pack(fill=X)
        protime_label = ttk.Label(protime_frame, text="Select Protime Data Folder:")
        protime_label.pack(side=LEFT)
        self.protime_path_label = ttk.Label(protime_frame, text="No folder selected", foreground='red')
        self.protime_path_label.pack(side=LEFT, padx=5)
        protime_button = ttk.Button(protime_frame, text="Browse", command=self.select_protime_folder)
        protime_button.pack(side=RIGHT)

        # Frame for Threshold
        threshold_frame = ttk.Frame(self.master, padding=10)
        threshold_frame.pack(fill=X)
        threshold_label = ttk.Label(threshold_frame, text="Set Hours Difference Threshold (Minutes):")
        threshold_label.pack(side=LEFT)
        self.threshold_minutes_var = tk.StringVar(value='')
        threshold_minutes_entry = ttk.Entry(threshold_frame, width=10, textvariable=self.threshold_minutes_var)
        threshold_minutes_entry.pack(side=LEFT, padx=5)

        # Frame for Week Selection
        week_frame = ttk.Frame(self.master, padding=10)
        week_frame.pack(fill=X)
        week_label = ttk.Label(week_frame, text="Select Week Range:")
        week_label.pack(side=LEFT)
        self.start_week_var = tk.StringVar()
        self.start_week_entry = ttk.Entry(week_frame, width=5, textvariable=self.start_week_var)
        self.start_week_entry.pack(side=LEFT, padx=5)
        week_to_label = ttk.Label(week_frame, text="to")
        week_to_label.pack(side=LEFT)
        self.end_week_var = tk.StringVar()
        self.end_week_entry = ttk.Entry(week_frame, width=5, textvariable=self.end_week_var)
        self.end_week_entry.pack(side=LEFT, padx=5)

        # Frame for Progress Bar
        progress_frame = ttk.Frame(self.master, padding=10)
        progress_frame.pack(fill=X)
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.pack(fill=X)

        # Frame for Run Button
        run_frame = ttk.Frame(self.master, padding=10)
        run_frame.pack(fill=X)
        self.run_button = ttk.Button(run_frame, text="Run", command=self.run_callback)  # Use the callback
        self.run_button.pack(side=LEFT)
        self.status_label = ttk.Label(run_frame, text="")
        self.status_label.pack(side=LEFT, padx=10)

    def select_agency_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.agency_path_label.config(text=folder_selected, foreground='green')
            self.agency_path = folder_selected
        else:
            self.agency_path_label.config(text="No folder selected", foreground='red')
            self.agency_path = ''

    def select_protime_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.protime_path_label.config(text=folder_selected, foreground='green')
            self.protime_path = folder_selected
        else:
            self.protime_path_label.config(text="No folder selected", foreground='red')
            self.protime_path = ''

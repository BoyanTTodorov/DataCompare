import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog
from concatFiles import ConcatFiles

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("COMPARE")
        self.root.geometry("400x200")
        
        # Creating a frame
        self.frame = ttk.Frame(self.root)
        self.frame.pack(pady=20)

        # Creating a label
        self.label = ttk.Label(self.frame, text="Default Folder Agency", font=("Helvetica", 16))
        self.label.pack(pady=10)

        # Creating a button
        self.button = ttk.Button(self.frame, text="Select another folder", command=self.open_file_dialog)
        self.button.pack(pady=10)

    def open_file_dialog(self):
        folder_selected = filedialog.askdirectory(title="Select a Folder",filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        agency = ConcatFiles(folder_selected)
        print(agency)
# controller.py
import ttkbootstrap as ttk
from tkinter import Tk, messagebox, filedialog
from view import DataView
from model import DataModel
import threading
import os

class DataController:
    def __init__(self):
        self.root = ttk.Window(themename="flatly")
        self.model = DataModel()
        # Pass the run_processing method as a callback to the view
        self.view = DataView(self.root, self.run_processing)
        print("Controller initialized")

    def run_processing(self):
        """
        Main process function starts modeling
        """
        print("Run button clicked")
        # Get inputs from the view
        try:
            self.model.set_agency_path(self.view.agency_path)
            self.model.set_protime_path(self.view.protime_path)
            self.model.set_threshold_minutes(self.view.threshold_minutes_var.get())
            start_week = self.view.start_week_var.get()
            end_week = self.view.end_week_var.get()
            print(f"Start Week Entry: {start_week}")
            print(f"End Week Entry: {end_week}")
            self.model.set_week_range(start_week, end_week)
            print(f"Agency Path: {self.view.agency_path}")
            print(f"Protime Path: {self.view.protime_path}")
            print(f"Threshold Minutes: {self.model.threshold_minutes}")
            print(f"Start Week: {self.model.start_week}")
            print(f"End Week: {self.model.end_week}")
            # Disable the run button during processing
            self.view.run_button.config(state='disabled')
            self.view.status_label.config(text="Processing...")
            self.view.progress_bar.start()
            threading.Thread(target=self.process_data_thread).start()
            print("Processing thread started")
        except Exception as e:
            # Show error message
            print(f"Error occurred: {e}")
            messagebox.showerror("Error", str(e))
            self.view.status_label.config(text=f"Error: {e}")

    def process_data_thread(self):
        try:
            print("Thread: Starting data processing")
            # Load data
            success = self.model.load_data()
            print("Thread: Data loaded")
            # Process data
            success = self.model.process_data()
            print("Thread: Data processed")

            # Ask user for save directory
            save_directory = filedialog.askdirectory(title="Select Save Directory")
            if not save_directory:
                raise Exception("No save directory selected.")
            print(f"Save Directory: {save_directory}")

            # Create 'Reports' folder in the save directory
            reports_folder = os.path.join(save_directory, 'Reports')
            os.makedirs(reports_folder, exist_ok=True)
            print(f"Reports Folder: {reports_folder}")

            # Save report
            report_name = f"diff_report_{self.get_timestamp()}.xlsx"
            output_file = os.path.join(reports_folder, report_name)
            self.model.save_report(output_file)
            print(f"Report saved to {output_file}")

            # Show success message
            messagebox.showinfo("Success", f"Report saved to {output_file}")

            self.view.status_label.config(text=f"Report saved to {output_file}")
        except Exception as e:
            # Show error message
            print(f"Error occurred: {e}")
            messagebox.showerror("Error", str(e))
            self.view.status_label.config(text=f"Error: {e}")
        finally:
            # Re-enable the run button
            self.view.run_button.config(state='normal')
            self.view.progress_bar.stop()
            print("Processing thread finished")

    def get_timestamp(self):
        from datetime import datetime
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    def start(self):
        print("Application started")
        self.root.mainloop()

if __name__ == '__main__':
    app = DataController()
    app.start()

import os
import json
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from workplace_manager import WorkplaceManager

DATA_FILE = "data.json"

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Schedule Manager")
        self.root.geometry("600x500")
        self.root.configure(bg="#f8f9fa")

        # load or start the data
        self.data = self.load_data()

        # setup UI
        self.create_widgets()

    def load_data(self):
        """Load saved data from file or initialize a new structure."""
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r") as file:
                return json.load(file)
        else:
            return {"workplaces": []}

    def save_data(self):
        """Save current data to file."""
        with open(DATA_FILE, "w") as file:
            json.dump(self.data, file, indent=4)

    def create_widgets(self):
        """Create the main menu interface."""
        title = ttk.Label(self.root, text="Schedule Manager", font=("Arial", 24, "bold"))
        title.pack(pady=20)

        # buttons for accessing various features
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        workplaces_btn = ttk.Button(btn_frame, text="Manage Workplaces", command=self.open_workplace_manager)
        workplaces_btn.pack(fill=tk.X, pady=5)

        view_schedule_btn = ttk.Button(btn_frame, text="📅 View Schedule", command=self.open_schedule_viewer)
        view_schedule_btn.pack(fill=tk.X, pady=5)

        generate_schedule_btn = ttk.Button(btn_frame, text="✏️ Create Weekly Schedule", command=self.open_weekly_schedule)
        generate_schedule_btn.pack(fill=tk.X, pady=5)

        exit_btn = ttk.Button(btn_frame, text="❌ Exit", command=self.exit_app)
        exit_btn.pack(fill=tk.X, pady=5)

    def open_workplace_manager(self):
        """Open the Workplace Manager window."""
        WorkplaceManager(self.root, self.data, self.save_data)

    def open_schedule_viewer(self):
        """Open the Schedule Viewer script in a subprocess."""
        try:
            if os.path.exists("schedule_viewer.py"):
                subprocess.Popen(["python", "schedule_viewer.py"])
            else:
                messagebox.showerror("Error", "Cannot find schedule_viewer.py!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Schedule Viewer: {e}")

    def open_weekly_schedule(self):
        """Open the Weekly Schedule Generator script in a subprocess."""
        try:
            if os.path.exists("weekly_schedule.py"):
                subprocess.Popen(["python", "weekly_schedule.py"])
            else:
                messagebox.showerror("Error", "Cannot find weekly_schedule.py!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Weekly Schedule Generator: {e}")

    def exit_app(self):
        """Save the data and exit the application."""
        self.save_data()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()

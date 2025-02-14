import os
import subprocess
import tkinter as tk
from tkinter import messagebox

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Main Schedule App")
        self.root.geometry("400x200")

        # title label
        self.title_label = tk.Label(root, text="Schedule App", font=("Helvetica", 16, "bold"))
        self.title_label.pack(pady=10)

        # button to open the schedule viewer
        self.schedule_viewer_button = tk.Button(
            root,
            text="Open Schedule Viewer",
            command=self.open_schedule_viewer,
            width=25
        )
        self.schedule_viewer_button.pack(pady=10)

        # button to open the weekly schedule Maker
        self.weekly_schedule_button = tk.Button(
            root,
            text="Open Weekly Schedule Maker",
            command=self.open_weekly_schedule,
            width=25
        )
        self.weekly_schedule_button.pack(pady=10)

        # exit button
        self.exit_button = tk.Button(
            root,
            text="Exit",
            command=self.exit_app,
            width=25
        )
        self.exit_button.pack(pady=10)

    def open_schedule_viewer(self):
        """Opens the Schedule Viewer script in a new process."""
        try:
            # make sure the schedule_viewer.py script exists on this device
            if os.path.exists("schedule_viewer.py"):
                subprocess.Popen(["python", "schedule_viewer.py"])
            else:
                messagebox.showerror("File Not Found", "The schedule_viewer.py file is missing.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Schedule Viewer: {e}")

    def open_weekly_schedule(self):
        """Opens the Weekly Schedule Maker script in a new process."""
        try:
            # make sure the weekly_schedule.py script exists on this device
            if os.path.exists("weekly_schedule.py"):
                subprocess.Popen(["python", "weekly_schedule.py"])
            else:
                messagebox.showerror("File Not Found", "The weekly_schedule.py file is missing.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Weekly Schedule Maker: {e}")

    def exit_app(self):
        """Exits the main application."""
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()

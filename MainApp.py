import os
import json
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from workplace_manager import WorkplaceManager

DATA_FILE = "data.json"

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Schedule Manager")
        self.root.geometry("600x500")
        self.root.configure(bg="#f8f9fa")

        # load existing data or start fresh
        self.data = self.load_data()

        # setup UI
        self.create_widgets()

    def load_data(self):
        if os.path.exists(DATA_FILE):
            with open(DATA_FILE, "r") as file:
                return json.load(file)
        else:
            return {"workplaces": []}

    def save_data(self):
        with open(DATA_FILE, "w") as file:
            json.dump(self.data, file, indent=4)

    def create_widgets(self):
        # title
        title = ttk.Label(self.root, text="Schedule Manager", font=("Arial", 24, "bold"))
        title.pack(pady=20)

        # buttons
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(pady=10)

        workplaces_btn = ttk.Button(btn_frame, text="Manage Workplaces", command=self.open_workplace_manager)
        workplaces_btn.pack(fill=tk.X, pady=5)

        exit_btn = ttk.Button(btn_frame, text="Exit", command=self.exit_app)
        exit_btn.pack(fill=tk.X, pady=5)

    def open_workplace_manager(self):
        """Launch the Workplace Manager window."""
        WorkplaceManager(self.root, self.data, self.save_data)

    def exit_app(self):
        self.save_data()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()

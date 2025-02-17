import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import json

# ------------------------------------------------------------------------------------------------------------------------

class WorkplaceManager:
    def __init__(self, root, data, save_callback):
        self.root = tk.Toplevel(root)
        self.root.title("Manage Workplaces")
        self.root.geometry("600x400")
        self.data = data
        self.save_callback = save_callback

        # workplace list
        self.workplace_listbox = tk.Listbox(self.root)
        self.workplace_listbox.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        self.workplace_listbox.bind("<<ListboxSelect>>", self.load_workplace_details)

        # control buttons
        control_frame = tk.Frame(self.root)
        control_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)

        tk.Button(control_frame, text="Add Workplace", command=self.add_workplace).pack(fill=tk.X, pady=5)
        tk.Button(control_frame, text="Delete Workplace", command=self.delete_workplace).pack(fill=tk.X, pady=5)
        tk.Button(control_frame, text="Modify Hours", command=self.modify_workplace_hours).pack(fill=tk.X, pady=5)
        tk.Button(control_frame, text="Save", command=self.save_data).pack(fill=tk.X, pady=5)

        # populate workplaces
        self.refresh_workplace_list()

    def refresh_workplace_list(self):
        self.workplace_listbox.delete(0, tk.END)
        for workplace in self.data["workplaces"]:
            self.workplace_listbox.insert(tk.END, workplace["name"])

    def add_workplace(self):
        new_name = simpledialog.askstring("Add Workplace", "Enter workplace name:")
        if new_name:
            self.data["workplaces"].append({"name": new_name, "hours": {}})
            self.refresh_workplace_list()

    def delete_workplace(self):
        selected_index = self.workplace_listbox.curselection()
        if selected_index:
            del self.data["workplaces"][selected_index[0]]
            self.refresh_workplace_list()
            self.save_data()

    def modify_workplace_hours(self):
        selected_index = self.workplace_listbox.curselection()
        if not selected_index:
            messagebox.showerror("Error", "Please select a workplace to modify.")
            return

        workplace = self.data["workplaces"][selected_index[0]]

        # modify hours for each day
        for day in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]:
            default_hours = workplace["hours"].get(day, "9 AM - 5 PM")
            new_hours = simpledialog.askstring(f"Modify Hours: {day}", f"Enter working hours for {day}:", initialvalue=default_hours)
            if new_hours:
                workplace["hours"][day] = new_hours

        messagebox.showinfo("Success", f"Updated hours for {workplace['name']}")

    def save_data(self):
        self.save_callback()
        self.refresh_workplace_list()
        messagebox.showinfo("Save", "Workplaces saved successfully!")

    def load_workplace_details(self, event):
        selected_index = self.workplace_listbox.curselection()
        if selected_index:
            workplace = self.data["workplaces"][selected_index[0]]
            messagebox.showinfo("Workplace Details", json.dumps(workplace, indent=4))


if __name__ == "__main__":
    print("Run the MainApp.py module to launch the full application.")

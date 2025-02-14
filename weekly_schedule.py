import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageDraw, ImageFont
from docx import Document
import random
import time  # importing time for randomizing

# --------------------------------------------------------------------------------------------------------------- #

class WeeklyScheduleGenerator:
    def __init__(self, root):
        """Initialize main GUI components."""
        self.root = root
        self.root.title("Weekly Schedule Generator")
        self.root.geometry("400x250")

        # buttons for the application workflow
        self.load_button = tk.Button(root, text="Load Excel Schedule", command=self.load_schedule)
        self.load_button.pack(pady=10)

        self.generate_availability_button = tk.Button(root, text="Generate Availability", command=self.generate_availability)
        self.generate_availability_button.pack(pady=10)

        self.generate_schedule_button = tk.Button(root, text="Generate Schedule", command=self.generate_schedule)
        self.generate_schedule_button.pack(pady=10)

        self.save_button = tk.Button(root, text="Save Word File", command=self.save_word_file)
        self.save_button.pack(pady=10)

        # data containers
        self.availability_data = None  # who's available for each time slot
        self.schedule_data = None  # final worker assignment for each time slot
        self.image_path = "final_schedule.png"  # path to save the generated image

# --------------------------------------------------------------------------------------------------------------- #

    def load_schedule(self):
        """Loads the Excel schedule file into a DataFrame."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        try:
            # load data into a pandas DataFrame
            self.schedule_df = pd.read_excel(file_path)

            # making sure all columns are their
            required_columns = {'First Name', 'Last Name', 'Email', 'Sunday', 'Monday', 'Tuesday',
                                'Wednesday', 'Thursday', 'Friday', 'Saturday'}
            if not required_columns.issubset(self.schedule_df.columns):
                messagebox.showerror("Error", f"Excel file must contain the following columns: {', '.join(required_columns)}")
                return

            messagebox.showinfo("Success", "Schedule file loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load schedule file: {e}")
            self.schedule_df = None

# --------------------------------------------------------------------------------------------------------------- #

    def generate_availability(self):
        """Processes the loaded schedule to create availability data per day and shift."""
        if self.schedule_df is None:
            messagebox.showerror("Error", "Please load a schedule file first.")
            return

        # defining available shifts and times
        shifts = {
            "Sunday": ["12 PM - 4 PM", "4 PM - 7 PM", "7 PM - 10 PM", "10 PM - 12 AM"],
            "Monday": ["2 PM - 5 PM", "5 PM - 8 PM", "8 PM - 12 AM"],
            "Tuesday": ["2 PM - 5 PM", "5 PM - 8 PM", "8 PM - 12 AM"],
            "Wednesday": ["2 PM - 6 PM", "6 PM - 9 PM", "9 PM - 12 AM"],
            "Thursday": ["2 PM - 4 PM", "4 PM - 8 PM", "8 PM - 12 AM"],
            "Friday": ["2 PM - 7 PM", "7 PM - 9 PM", "9 PM - 12 AM"],
            "Saturday": ["12 PM - 4 PM", "4 PM - 8 PM", "8 PM - 12 AM"]
        }

        # making a nested dictionary for availability
        availability = {day: {shift: [] for shift in shifts[day]} for day in shifts.keys()}

        for _, row in self.schedule_df.iterrows():
            for day in availability.keys():
                if pd.notna(row[day]) and row[day].lower() != 'na':  # Skip "na" or empty values
                    for shift in availability[day]:
                        availability[day][shift].append(f"{row['First Name']} {row['Last Name']}")

        self.availability_data = availability
        messagebox.showinfo("Success", "Availability data created successfully!")

# --------------------------------------------------------------------------------------------------------------- #

    def generate_schedule(self):
        """Randomly generates schedule based on availability data."""
        if not self.availability_data:
            messagebox.showerror("Error", "Please generate availability first.")
            return

        # initialize a schedule dictionary with the same structure as availability
        schedule = {day: {shift: None for shift in shifts} for day, shifts in self.availability_data.items()}

        # use current time as a random seed to ensure randomized results
        random.seed(time.time())

        # assign workers to shifts
        for day, shifts in self.availability_data.items():
            assigned_workers = set()  # keep track of who has already been assigned

            for shift, workers in shifts.items():
                random.shuffle(workers)  # randomize the order of available workers
                for worker in workers:
                    if worker not in assigned_workers:
                        schedule[day][shift] = worker
                        assigned_workers.add(worker)
                        break

                # if no one is available, leave the shift unassigned
                if not schedule[day][shift] and workers:
                    schedule[day][shift] = workers[0]  # Assign someone to avoid gaps

        self.schedule_data = schedule

        # make the PNG with final schedule
        self.create_calendar_image(schedule, title="Final Weekly Schedule")
        messagebox.showinfo("Success", "Final weekly schedule generated and image saved!")

# --------------------------------------------------------------------------------------------------------------- #

    def create_calendar_image(self, data, title="Weekly Schedule"):
        """Generates a PNG visual representation of the schedule."""
        # picture settings
        width, height = 1200, 1000
        header_height = 100
        color_bg = (255, 255, 255)
        color_header = (70, 130, 180)
        color_text = (0, 0, 0)

        # create a blank image
        img = Image.new("RGB", (width, height), color_bg)
        draw = ImageDraw.Draw(img)

        # add title to the image
        font_title = ImageFont.truetype("arial.ttf", 40)
        draw.text((width // 2 - 200, 10), title, fill=color_header, font=font_title)

        # draw schedule data
        font_body = ImageFont.truetype("arial.ttf", 20)
        row_height = (height - header_height) // len(data)

        for i, (day, shift_data) in enumerate(data.items()):
            y_start = header_height + i * row_height

            # draw the day header
            draw.text((10, y_start), day, fill=color_header, font=font_body)

            # add shifts and corresponding workers
            y_shift = y_start + 40
            for shift, worker in shift_data.items():
                shift_text = f"{shift}: {worker if worker else 'No one assigned'}"
                draw.text((40, y_shift), shift_text, fill=color_text, font=font_body)
                y_shift += 30

        # save the image as a PNG file
        img.save(self.image_path)

# --------------------------------------------------------------------------------------------------------------- #

def save_word_file(self):
    """Saves the generated schedule and worker summary to a Word document with improved formatting."""
    if not self.schedule_data:
        messagebox.showerror("Error", "Please generate the final schedule first.")
        return

    try:
        # make a new Word document
        doc = Document()
        
        # make a title
        title = doc.add_heading("Weekly Schedule", level=1)
        title.alignment = 1  # centered
        doc.add_paragraph()  # add spacing
        
        # track worker hours with shift length
        worker_hours = {}
        
        # helper func to calculate shift length
        def calculate_shift_duration(shift_time):
            start, end = shift_time.split(" - ")
            
            # change to 24 hour format for calculation
            def convert_to_24hr(time_str):
                time = int(time_str.split()[0])
                if "PM" in time_str and time != 12:
                    time += 12
                elif "AM" in time_str and time == 12:
                    time = 0
                return time
            
            start_hour = convert_to_24hr(start)
            end_hour = convert_to_24hr(end)
            
            # handle shifts that cross midnight
            if end_hour < start_hour:
                end_hour += 24
                
            return end_hour - start_hour

        # write schedule data to the Word document with improved formatting
        for day, shifts in self.schedule_data.items():
            # add day header with styling
            day_heading = doc.add_heading(day, level=2)
            day_heading.style.font.color.rgb = None  # reset the color to default
            
            # make a table for each day's shifts
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            table.allow_autofit = True
            
            # set header row
            header_cells = table.rows[0].cells
            header_cells[0].text = "Shift Time"
            header_cells[1].text = "Worker"
            
            # add shifts to table
            for shift, worker in shifts.items():
                row_cells = table.add_row().cells
                row_cells[0].text = shift
                row_cells[1].text = worker if worker else 'Unassigned'
                
                # calculate and track hours for each worker
                if worker:
                    duration = calculate_shift_duration(shift)
                    worker_hours[worker] = worker_hours.get(worker, 0) + duration
            
            doc.add_paragraph()  # add spacing between days

        # add worker summary section
        doc.add_paragraph()  # add spacing
        summary_heading = doc.add_heading("Weekly Hours Summary", level=2)
        summary_heading.style.font.color.rgb = None  # reset color to default
        
        # make table for hours summary
        summary_table = doc.add_table(rows=1, cols=2)
        summary_table.style = 'Table Grid'
        
        # make a summary table header
        header_cells = summary_table.rows[0].cells
        header_cells[0].text = "Worker Name"
        header_cells[1].text = "Total Hours"
        
        # sort workers by hours (descending)
        sorted_workers = sorted(worker_hours.items(), key=lambda x: x[1], reverse=True)
        
        # add worker hours to summary table
        for worker, hours in sorted_workers:
            row_cells = summary_table.add_row().cells
            row_cells[0].text = worker
            row_cells[1].text = f"{hours} hours"
        
        # add footer with generation date
        doc.add_paragraph()
        footer = doc.add_paragraph(f"Generated on: {time.strftime('%Y-%m-%d %H:%M:%S')}")
        footer.alignment = 1  # center alignment

        # save word file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")],
            initialfile="Weekly_Schedule.docx"
        )
        if file_path:
            doc.save(file_path)
            messagebox.showinfo("Success", f"Word file saved successfully at {file_path}.")
            
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save Word file: {e}")

# --------------------------------------------------------------------------------------------------------------- #

if __name__ == "__main__":
    root = tk.Tk()
    app = WeeklyScheduleGenerator(root)
    root.mainloop()

# --------------------------------------------------------------------------------------------------------------- #

import os
import sys
import shutil
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk  # Import ttk for progress bar
from datetime import timedelta

# Function to install required libraries
def install_packages():
    packages = ["pandas", "tkcalendar", "openpyxl"]
    for package in packages:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except subprocess.CalledProcessError:
            print(f"Failed to install {package}. Please install it manually.")

# Check for pandas and install packages if missing
try:
    import pandas as pd
except ImportError:
    install_packages()
    import pandas as pd  # Retry importing after installation

# Attempt to import tkcalendar, install if missing
try:
    from tkcalendar import Calendar
except ImportError:
    install_packages()
    from tkcalendar import Calendar  # Retry importing after installation

# Load local folder path from file
def load_local_folder_path():
    if os.path.exists("local_folder_path.txt"):
        with open("local_folder_path.txt", "r") as f:
            return f.read().strip()
    return None

# Save local folder path to file
def save_local_folder_path(path):
    with open("local_folder_path.txt", "w") as f:
        f.write(path)

# Function to create/update the local folder from the shared folder with progress
def update_local_folder(shared_folder_path, local_folder_path):
    # Display start message and initialize progress
    progress_label.config(text="Copying started...")
    app.update_idletasks()

    # Remove the existing local folder if it exists
    if os.path.exists(local_folder_path):
        shutil.rmtree(local_folder_path)

    # Create the local folder
    os.makedirs(local_folder_path)

    # Get the total number of files and directories to copy
    total_items = sum(len(files) for _, _, files in os.walk(shared_folder_path))

    # Copy the contents of the shared folder to the local folder with progress update
    copied_items = 0
    for root, dirs, files in os.walk(shared_folder_path):
        local_root = os.path.join(local_folder_path, os.path.relpath(root, shared_folder_path))
        os.makedirs(local_root, exist_ok=True)

        for file in files:
            shutil.copy2(os.path.join(root, file), local_root)
            copied_items += 1
            progress_percentage = (copied_items / total_items) * 100
            progress_label.config(text=f"Copying... {progress_percentage:.2f}%")
            app.update_idletasks()

    os.startfile(local_folder_path)
    progress_label.config(text="Copy Successful!")
    update_local_info(local_folder_path)
    app.update_idletasks()

    # Enable the search button once copying is complete
    search_button.config(state=tk.NORMAL)


# Function to search for multiple unit numbers or phrases in an Excel or CSV file
def search_in_file(file_path, search_values):
    found_units = []
    not_found_units = set(search_values)  # Track units not found initially

    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
            dfs = {'CSV': df}  # Treat CSV as a single "sheet" named 'CSV'
        else:
            excel_file = pd.ExcelFile(file_path)
            dfs = {sheet: pd.read_excel(file_path, sheet_name=sheet) for sheet in excel_file.sheet_names}

        for search_value in search_values:
            found = False
            search_value_str = str(search_value).strip()  # Convert to string and strip whitespace

            # Loop through each sheet (or CSV)
            for sheet_name, df in dfs.items():
                # Convert all DataFrame cells to strings and use contains() for partial matching
                if df.astype(str).apply(lambda x: x.str.contains(search_value_str, case=False, na=False)).any().any():
                    found_units.append((search_value, sheet_name))
                    not_found_units.discard(search_value)  # Remove from not found
                    found = True
                    break  # Stop after finding in one sheet

    except Exception as e:
        print(f"Error reading {file_path}: {e}")

    return found_units, not_found_units

# Function to search for the latest updated file
def search_latest_file(search_values, directory_path):
    latest_file, latest_mod_date = None, None
    found_units, not_found_units = [], set()

    for file_name in os.listdir(directory_path):
        if file_name.endswith(('.xlsx', '.xls', '.csv')):
            file_path = os.path.join(directory_path, file_name)
            mod_time = os.path.getmtime(file_path)
            if latest_mod_date is None or mod_time > latest_mod_date:
                latest_file, latest_mod_date = file_path, mod_time

    if latest_file:
        found_units, not_found_units = search_in_file(latest_file, search_values)
        latest_mod_date = datetime.fromtimestamp(latest_mod_date)

    return found_units, not_found_units, latest_file, latest_mod_date

# Function to search all files within a date range
def search_all_files(search_values, directory_path, start_date, end_date):
    results, not_found_units = [], set(search_values)

    for file_name in os.listdir(directory_path):
        if file_name.endswith(('.xlsx', '.xls', '.csv')):
            file_path = os.path.join(directory_path, file_name)
            file_date = datetime.fromtimestamp(os.path.getmtime(file_path))

            if start_date <= file_date <= end_date:
                found_units, units_not_found = search_in_file(file_path, search_values)
                for unit, sheet_name in found_units:
                    results.append((file_name, sheet_name, file_date, unit))
                    not_found_units.discard(unit)

    return results, not_found_units

# Function to open a calendar for date selection
def open_calendar(entry):
    if not hasattr(open_calendar, 'cal_window') or open_calendar.cal_window is None or not open_calendar.cal_window.winfo_exists():
        open_calendar.cal_window = tk.Toplevel(app)
        open_calendar.cal_window.title("Select Date")
        cal = Calendar(open_calendar.cal_window, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=10)

        def select_date():
            selected_date = cal.get_date()
            entry.delete(0, tk.END)
            entry.insert(0, selected_date)
            open_calendar.cal_window.destroy()

        tk.Button(open_calendar.cal_window, text="Select", command=select_date).pack(pady=10)
    else:
        open_calendar.cal_window.destroy()
        open_calendar.cal_window = None

# Function to toggle date entry states based on the checkbox
def toggle_date_entries():
    if latest_file_var.get():
        start_date_entry.config(state=tk.DISABLED)
        end_date_entry.config(state=tk.DISABLED)
    else:
        start_date_entry.config(state=tk.NORMAL)
        end_date_entry.config(state=tk.NORMAL)

# Function to browse for a shared directory
def browse_shared_directory():
    shared_directory = filedialog.askdirectory()
    if shared_directory:
        shared_directory_entry.delete(0, tk.END)
        shared_directory_entry.insert(0, shared_directory)
        local_folder_path = os.path.join(os.path.expanduser("~/Desktop"), "Local_CycleCount")
        update_local_folder(shared_directory, local_folder_path)

# Skip copy and go directly to the search page
def skip_copy():
    local_folder_path = os.path.join(os.path.expanduser("~/Desktop"), "Local_CycleCount")
    if os.path.exists(local_folder_path):
        update_local_info(local_folder_path)
        search_button.config(state=tk.NORMAL)  # Enable search button if local folder exists
    switch_to_search_page()

# Switch to search page after folder creation or skipping
def switch_to_search_page():
    search_button.config(state=tk.NORMAL)

# Function to update local folder info
def update_local_info(local_folder_path):
    last_mod_time = datetime.fromtimestamp(os.path.getmtime(local_folder_path))
    folder_info_label.config(text=f"Local Folder: {os.path.basename(local_folder_path)}\nLast Copied/Updated: {last_mod_time}")
    save_local_folder_path(local_folder_path)  # Save path for persistence

def perform_search():
    local_folder_path = os.path.join(os.path.expanduser("~/Desktop"), "Local_CycleCount")
    search_values = [value.strip() for value in search_entry.get().split(',')]
    show_latest = latest_file_var.get()

    if not os.path.exists(local_folder_path):
        messagebox.showerror("Error", "The local folder does not exist. Please try again.")
        return

    try:
        if not show_latest:
            start_date = datetime.strptime(start_date_entry.get(), '%Y-%m-%d')
            # Extend the end_date to include the full day
            end_date = datetime.strptime(end_date_entry.get(), '%Y-%m-%d') + timedelta(days=1) - timedelta(seconds=1)
        else:
            start_date = end_date = None
    except ValueError:
        messagebox.showerror("Error", "Please enter dates in the format YYYY-MM-DD.")
        return

    results, not_found_units = [], set()

    if show_latest:
        found_units, not_found_units, latest_file, latest_mod_date = search_latest_file(search_values, local_folder_path)
        if found_units:
            for unit, sheet_name in found_units:
                results.append((os.path.basename(latest_file), sheet_name, latest_mod_date, unit))
    else:
        results, not_found_units = search_all_files(search_values, local_folder_path, start_date, end_date)

    # Display results in the text box
    result_text.delete(1.0, tk.END)

    if results:
        result_text.insert(tk.END, "Found the following units in the files:\n\n")
        for file_name, sheet_name, mod_date, found_value in results:
            result_text.insert(tk.END, f"Unit: {found_value}\nFile: {file_name}\nSheet: {sheet_name if sheet_name else 'N/A (CSV)'}\nLast Saved: {mod_date}\n" + "-" * 50 + "\n")
    else:
        result_text.insert(tk.END, "No units found.\n")

    if not_found_units:
        result_text.insert(tk.END, "\nThe following units were not found:\n")
        for unit in not_found_units:
            result_text.insert(tk.END, f"{unit}\n")

# Set up the main application window
app = tk.Tk()
app.title("Local Cycle Count Manager")

# Set up the frames
frame1 = tk.Frame(app)
frame1.pack(pady=10)

frame2 = tk.Frame(app)
frame2.pack(pady=10)

frame3 = tk.Frame(app)
frame3.pack(pady=10)

# Set up the shared directory entry
shared_directory_label = tk.Label(frame1, text="Shared Folder:")
shared_directory_label.grid(row=0, column=0)

shared_directory_entry = tk.Entry(frame1, width=50)
shared_directory_entry.grid(row=0, column=1)

browse_button = tk.Button(frame1, text="Browse", command=browse_shared_directory)
browse_button.grid(row=0, column=2)

# Set up the buttons
copy_button = tk.Button(frame2, text="Copy Files", command=browse_shared_directory)
copy_button.grid(row=0, column=0, padx=5)

skip_button = tk.Button(frame2, text="Skip Copy", command=skip_copy)
skip_button.grid(row=0, column=1, padx=5)

# Set up the folder info label
folder_info_label = tk.Label(frame3, text="")
folder_info_label.pack()

# Set up the progress bar and label
progress_label = tk.Label(frame3, text="")
progress_label.pack()

# Search frame
search_frame = tk.Frame(app)
search_frame.pack(pady=10)

search_label = tk.Label(search_frame, text="Search for Unit Numbers (comma-separated):")
search_label.grid(row=0, column=0)

search_entry = tk.Entry(search_frame, width=50)
search_entry.grid(row=0, column=1)

# Date selection for range search
latest_file_var = tk.BooleanVar()
latest_file_check = tk.Checkbutton(search_frame, text="Search Latest File", variable=latest_file_var, command=toggle_date_entries)
latest_file_check.grid(row=1, column=0, columnspan=2)

start_date_label = tk.Label(search_frame, text="Start Date (YYYY-MM-DD):")
start_date_label.grid(row=2, column=0)

start_date_entry = tk.Entry(search_frame, width=20)
start_date_entry.grid(row=2, column=1)

# Replace the button with a larger emoji label
start_date_emoji = tk.Label(search_frame, text="📅", font=("Arial", 16), cursor="hand2")  # Increased font size
start_date_emoji.grid(row=2, column=2)  # No padding for close alignment

# Function to handle the emoji click to open the calendar
start_date_emoji.bind("<Button-1>", lambda e: open_calendar(start_date_entry))

end_date_label = tk.Label(search_frame, text="End Date (YYYY-MM-DD):")
end_date_label.grid(row=3, column=0)

end_date_entry = tk.Entry(search_frame, width=20)
end_date_entry.grid(row=3, column=1)

# Replace the button with a larger emoji label
end_date_emoji = tk.Label(search_frame, text="📅", font=("Arial", 16), cursor="hand2")  # Increased font size
end_date_emoji.grid(row=3, column=2)  # No padding for close alignment

# Function to handle the emoji click to open the calendar
end_date_emoji.bind("<Button-1>", lambda e: open_calendar(end_date_entry))


search_button = tk.Button(search_frame, text="Search", command=perform_search, state=tk.DISABLED)
search_button.grid(row=4, column=1, pady=5)

# Set up the result display area
result_text = scrolledtext.ScrolledText(app, wrap=tk.WORD, width=70, height=20)
result_text.pack(pady=10)

# Load the local folder path and initialize the app
local_folder_path = load_local_folder_path()
if local_folder_path:
    update_local_info(local_folder_path)  # Update UI with loaded path
    search_button.config(state=tk.NORMAL)  # Enable search button if local folder exists

# Add fine print watermark at the bottom
watermark_label = tk.Label(
    app,
    text="Version: 1.0.1 - Developer: Ibrahim Sanduqah",
    font=("Arial", 9),
    fg="gray"
)
watermark_label.pack(side=tk.BOTTOM, pady=5)

# Start the application
app.mainloop()
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import os
from pandas import DataFrame, concat, ExcelWriter, ExcelFile, read_excel
import report_generator as rg
import threading
from datetime import datetime, timezone, timedelta
import webbrowser

program_name = 'LockitCheck'
program_version = '0.1b'
program_update_latest = '2023-11-20'
program_url = 'https://github.com/wtigga/lockitcheck'


def open_html_in_browser(file_path):
    # Check if the file exists
    if not os.path.exists(file_path):
        print("File not found:", file_path)
        return

    # Open the file in the default web browser
    webbrowser.open('file://' + os.path.realpath(file_path))


def browse_folder():
    start_progress()
    buttons_disable()

    def main_logic():
        folder_location = filedialog.askdirectory()
        folder_path_var.set(folder_location)
        # print(folder_location)
        if folder_location:
            print("Extracting headers...")
            extract_header_names(folder_location)
        buttons_enable()
        stop_progress()

    main_thread = threading.Thread(target=main_logic)
    main_thread.start()


def process_files():
    buttons_disable()

    def task():
        now = datetime.now()
        timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")  # Required for the filename
        report_save_path = os.path.join(os.getcwd(), 'report_{}.html'.format(timestamp))
        print(source_var.get())
        print(target_var.get())
        print(id_var.get())
        rg.create_report(folder_source=folder_path_var.get(), col_src=source_var.get(), col_tgt=target_var.get(),
                         col_id=id_var.get(), find_latin=check_latin.get(), find_chinese=check_chinese.get(),
                         output_file=report_save_path)
        open_html_in_browser(report_save_path)

        # Stop the progress bar and update the GUI
        stop_progress()
        buttons_enable()
        print("Finished")

    # Start the progress bar
    start_progress()

    # Run the task in a separate thread
    threading.Thread(target=task).start()


class ExcelProcessingError(Exception):
    """Custom exception for errors during extract_header_names function."""


def buttons_disable():
    browse_folder_button.config(state=tk.DISABLED)
    combobox_id.config(state=tk.DISABLED)
    combobox_target.config(state=tk.DISABLED)
    combobox_source.config(state=tk.DISABLED)
    process_button.config(state=tk.DISABLED)
    print("Disabled")


def buttons_enable():
    browse_folder_button.config(state=tk.NORMAL)
    combobox_id.config(state='readonly')
    combobox_target.config(state='readonly')
    combobox_source.config(state='readonly')
    process_button.config(state='readonly')
    print("Disabled")


def extract_header_names(folder):
    try:
        header_set = set()  # Initialize an empty set to store unique headers

        # Get a list of Excel files in the input folder
        input_file_list = [file for file in os.listdir(folder) if file.endswith('.xlsx')]

        if not input_file_list:
            messagebox.showinfo(f'No files in {folder}', f"Please note that no Excel files were found in {folder}.")

        # Loop through each Excel file
        for file in input_file_list:

            file_path = os.path.join(folder, file)

            # Open the Excel file
            try:
                excel = ExcelFile(file_path)
            except Exception as e:
                error_opening_excel = 'Error opening Excel file'
                raise ExcelProcessingError(f"{error_opening_excel} '{file}': {e}")

            # Iterate through each sheet in the Excel file
            for sheet_title in excel.sheet_names:
                # Read the sheet into a DataFrame
                try:
                    dataframe = excel.parse(sheet_title)
                except Exception as e:
                    error_reading_sheet = 'Error reading sheet'
                    in_file = 'in file'
                    raise ExcelProcessingError(f"{error_reading_sheet} '{sheet_title}' {in_file} '{file}': {e}")
                header_set.update(dataframe.columns)
        header_set = sorted(list(header_set))
        combobox_source['values'] = header_set
        combobox_target['values'] = header_set
        combobox_id['values'] = header_set
        source_var.set(header_set[0])
        id_var.set(header_set[0])
        target_var.set(header_set[0])
        print("Combobox updated")
        return header_set
    except Exception as e:
        # Handle any unexpected exceptions by raising a custom error
        sys.exit()


root = tk.Tk()
title = f"{program_name}, {program_version} ({program_update_latest})"
root.title(title)

# Folder browsing
folder_path_var = tk.StringVar()
browse_folder_button = tk.Button(root, text="Browse Folder", command=browse_folder)
browse_folder_button.grid(row=0, column=0, padx=10, pady=10, sticky='w')
folder_entry = ttk.Entry(root, textvariable=folder_path_var, width=60)
folder_entry.grid(row=0, column=1, padx=10, pady=10, sticky='w', columnspan=2)

# Labels and Dropdown comboboxes
tk.Label(root, text="Source Column").grid(row=1, column=0, padx=10, pady=5)
tk.Label(root, text="Target Column").grid(row=1, column=1, padx=10, pady=5)
tk.Label(root, text="ID Column").grid(row=1, column=2, padx=10, pady=5)

source_var = tk.StringVar()
target_var = tk.StringVar()
id_var = tk.StringVar()

combobox_source = ttk.Combobox(root, textvariable=source_var, values='',
                               width=20,
                               state='readonly')
combobox_source.grid(row=2, column=0, padx=10, pady=5)

combobox_target = ttk.Combobox(root, textvariable=target_var, values='',
                               width=20,
                               state='readonly')
combobox_target.grid(row=2, column=1, padx=10, pady=5)

combobox_id = ttk.Combobox(root, textvariable=id_var, values='',
                           width=20,
                           state='readonly')
combobox_id.grid(row=2, column=2, padx=10, pady=5)

# Checkboxes
check_latin = tk.BooleanVar()
check_latin.set(True)
check_chinese = tk.BooleanVar()
check_chinese.set(True)
checkbox_latin = ttk.Checkbutton(root, text="Check for Latin in Target", variable=check_latin).grid(row=3, column=0,
                                                                                                    padx=10, pady=10)
checkbox_chinese = ttk.Checkbutton(root, text="Check for Chinese in Target", variable=check_chinese).grid(row=3,
                                                                                                          column=1,
                                                                                                          padx=10,
                                                                                                          pady=10)

# Process button
process_button = ttk.Button(root, text="Process", command=process_files)
process_button.grid(row=4, column=0, padx=10, pady=10, sticky='w')


# Progress bar

def start_progress():
    # Set the progress bar to the 'active' style
    progress['mode'] = 'indeterminate'
    progress.start(10)


def stop_progress():
    # Stop the progress bar and switch to 'inactive' style
    progress.stop()
    progress['mode'] = 'determinate'


def open_source_code(event):
    # the event is not used, but required for mouse-click in gui
    webbrowser.open_new(program_url)


progress_var = tk.IntVar(value=0)
progress = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress.grid(row=4, column=1, padx=10, pady=10, columnspan=2, sticky='w')

manual_label = tk.Label(root, text="https://github.com/wtigga/lockitcheck", justify="left", cursor="hand2", fg="blue")
manual_label.bind("<Button-1>", open_source_code)
manual_label.grid(row=5, column=2, padx=10, pady=0, sticky='e')

root.mainloop()

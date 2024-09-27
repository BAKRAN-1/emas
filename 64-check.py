from decimal import Decimal
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog, Listbox, ttk  # These are still used from Tkinter
import tkinter as tk
import os
import pickle
import re
from tkcalendar import DateEntry
from datetime import datetime
import pyodbc
import json
import random
import string
import logging
import sys
import requests

# File to store the session state
STATE_FILE = "session_state.pkl"
 
# Temporary storage for processed dataframes and hospital names
processed_dataframes = []
extracted_hospitals = []
recon_data = {}

is_data_extracted = False  # Tracks if new data has been extracted
is_data_saved = True  # Tracks if the data has been saved after extraction

extracted_sheets = []  # Use a set to keep track of extracted sheets

# Add a global variable to store the file path
selected_file_path = None
sheet_df = None

# Function to save user input to the master Excel sheet
def save_to_master_file(master_file_path):
    try:
        # Load the master file
        df = pd.read_excel(master_file_path)

        # Check if columns for Recon By, Recon Date, and SOA Date exist, add if not
        if 'Recon By' not in df.columns:
            df['Recon By'] = pd.NA
        if 'Recon Date' not in df.columns:
            df['Recon Date'] = pd.NA
        if 'SOA Date' not in df.columns:
            df['SOA Date'] = pd.NA

        # Append new row with recon info
        new_row = {'Recon By': recon_data['Recon By'],
                   'Recon Date': recon_data['Recon Date'],
                   'SOA Date': recon_data['SOA Date']}
        df = df.append(new_row, ignore_index=True)

        # Save back to the file
        df.to_excel(master_file_path, index=False)
        print("Data successfully saved to Excel.")
    except Exception as e:
        print(f"Error while saving to Excel: {e}")

# Updated sign_in_window function with requested behavior
def sign_in_window():
    global form_submitted  # Ensure we're modifying the global form_submitted variable

    ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "dark-blue", "green"

    sign_in = ctk.CTk()  # Create the sign-in window
    form_submitted = tk.BooleanVar(value=False)  # Initialize BooleanVar after the window is created
    sign_in.title("Sign-In Window")
    sign_in.geometry("400x400")

    # Disable maximize/restore down button by making the window not resizable
    sign_in.resizable(False, False)

    # Function to handle window close event (regardless of "Recon By" field)
    def on_close():
        # Simply close both windows and exit the program if the user closes the window
        messagebox.showwarning("Exiting", "You closed the window without submitting. Exiting the program.")
        sign_in.destroy()  # Close the sign-in window
        sys.exit()  # Exit the whole program

    # Set the protocol to handle window close (X button)
    sign_in.protocol("WM_DELETE_WINDOW", on_close)

    # Welcome message
    ctk.CTkLabel(sign_in, text="Welcome to the Reconciliation Tool", font=("Helvetica", 18)).pack(pady=10)

    # Recon By Input
    ctk.CTkLabel(sign_in, text="Recon By:", font=("Helvetica", 14)).pack(pady=5)
    recon_by_entry = ctk.CTkEntry(sign_in, placeholder_text="Enter your name", font=("Helvetica", 14))
    recon_by_entry.pack(pady=5)

    # Recon Date Input (auto-set to today's date, larger font)
    ctk.CTkLabel(sign_in, text="Recon Date:", font=("Helvetica", 14)).pack(pady=5)
    recon_date_entry = DateEntry(sign_in, date_pattern="y/mm/dd", font=("Helvetica", 16), width=15)
    recon_date_entry.pack(pady=5)
    recon_date_entry.set_date(datetime.today())  # Set to current date automatically

    # SOA Date Input with calendar, larger font
    ctk.CTkLabel(sign_in, text="SOA Date:", font=("Helvetica", 14)).pack(pady=5)
    soa_date_entry = DateEntry(sign_in, date_pattern="y/mm/dd", font=("Helvetica", 16), width=15)
    soa_date_entry.pack(pady=5)

    # Submit Button
    submit_button = ctk.CTkButton(sign_in, text="Submit", font=("Helvetica", 16),
                                  command=lambda: submit_action(sign_in, recon_by_entry, recon_date_entry, soa_date_entry))
    submit_button.pack(pady=20)

    # Function to handle submit button click
    def submit_action(window, recon_by, recon_date, soa_date):
        recon_by_value = recon_by.get().strip()  # Get the value and remove any spaces
        recon_date_value = recon_date.get()
        soa_date_value = soa_date.get()

        # If the "Recon By" field is empty, close everything
        if not recon_by_value:
            messagebox.showwarning("Missing Name", "You must enter a name for 'Recon By'. Exiting...")
            window.destroy()  # Close the sign-in window
            sys.exit()  # Exit the whole program

        # Validate input fields
        if recon_by_value and recon_date_value and soa_date_value:
            recon_data['Recon By'] = recon_by_value
            recon_data['Recon Date'] = recon_date_value
            recon_data['SOA Date'] = soa_date_value
            messagebox.showinfo("Success", "Data saved successfully.")
            window.destroy()  # Close the sign-in window
            main_interface()  # Launch the main interface after saving data
        else:
            messagebox.showwarning("Missing Information", "All fields are required.")

    # Enable navigation with Enter key or Up/Down keys
    def on_enter_key(event, current_entry, next_entry):
        if next_entry:
            next_entry.focus_set()

    # Bind the fields for up/down navigation and enter navigation
    recon_by_entry.bind("<Return>", lambda event: on_enter_key(event, recon_by_entry, recon_date_entry))
    recon_date_entry.bind("<Return>", lambda event: on_enter_key(event, recon_date_entry, soa_date_entry))
    soa_date_entry.bind("<Return>", lambda event: on_enter_key(event, soa_date_entry, submit_button))

    recon_by_entry.bind("<Down>", lambda event: recon_date_entry.focus_set())
    recon_by_entry.bind("<Up>", lambda event: recon_by_entry.focus_set())

    recon_date_entry.bind("<Down>", lambda event: soa_date_entry.focus_set())
    recon_date_entry.bind("<Up>", lambda event: recon_by_entry.focus_set())

    soa_date_entry.bind("<Down>", lambda event: submit_button.focus_set())
    soa_date_entry.bind("<Up>", lambda event: recon_date_entry.focus_set())

    # Start the Tkinter event loop
    sign_in.mainloop()

# Placeholder for the main interface function (you can implement this as needed)
def main_interface():
    print("Main interface loaded.")

# Launch the sign-in window on start
sign_in_window()

# Function to change appearance mode
def change_appearance(mode):
    ctk.set_appearance_mode(mode)  # Set the appearance mode (Light, Dark, or System)

# Function to handle pinning/unpinning the window
def toggle_pin():
    if pin_var.get():
        root.attributes("-topmost", True)  # Set the window to always stay on top
    else:
        root.attributes("-topmost", False)  # Allow the window to be overlapped by others

# Function to generate a random 3-letter string
def generate_random_string(length=3):
    letters = string.ascii_uppercase
    return ''.join(random.choice(letters) for i in range(length))

# Function to create 'reconref' with VARCHAR(11) length
def create_reconref(soa_date, recon_date):
    # Ensure date strings are in 'YYYY-MM-DD' format
    if pd.isna(soa_date) or pd.isna(recon_date):
        return ''  # Return empty string if either date is NaN
    
    day_month_soa = soa_date[8:10] + soa_date[5:7]  # Extract day and month from soa_date
    day_month_recon = recon_date[8:10] + recon_date[5:7]  # Extract day and month from recon_date
    random_str = generate_random_string()
    base_key = f"{day_month_soa}{day_month_recon}{random_str}"
    return base_key[:11]  # Ensure the length is 11 characters

def load_sheet(file_path, sheet_name):
    global sheet_df
    try:
        sheet_df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load sheet: {e}")
        return None
    return sheet_df

# Function to update the available column suggestions and disable already selected ones
def update_column_suggestions():
    if sheet_df is not None:
        columns = sheet_df.columns
        # Convert index to Excel column letters (A, B, C, etc.)
        column_letters = [chr(65 + i) for i in range(len(columns))]

        # Track the selected columns
        selected_columns = [
            bill_date_combobox.get(),
            bill_number_combobox.get(),
            bill_amount_combobox.get(),
            second_bill_number_combobox.get()
        ]

        # Create available columns by excluding already selected ones
        available_columns = [letter for letter in column_letters if letter not in selected_columns]

        # Update the Combobox values dynamically based on selected columns
        if not bill_date_combobox.get():
            bill_date_combobox['values'] = available_columns
        if not bill_number_combobox.get():
            bill_number_combobox['values'] = available_columns
        if not bill_amount_combobox.get():
            bill_amount_combobox['values'] = available_columns
        if not second_bill_number_combobox.get():
            second_bill_number_combobox['values'] = available_columns

class connection: 
    
    def pandas_dataframe(BillNo ):
        with open('db.json') as app_data:
            d = json.load(app_data)

        cnxn_string = 'Driver={SQL Server};SERVER='+d["db connection"]["server"]+';DATABASE='+d["db connection"]["database"]+';UID='+d["db connection"]["username"]+';PWD='+d["db connection"]["password"]
        cnxn = pyodbc.connect(cnxn_string)
        cursor = cnxn.cursor()

        if BillNo:
            placeholders = ','.join('?' for _ in BillNo )
            query = f"SELECT * FROM VW_SOAReconInfo WHERE BillNo IN ({placeholders})"
            cursor.execute(query, BillNo )

        else:
            # Handle the case where BillNo is empty
            query = "SELECT * FROM VW_SOAReconInfo WHERE 1=0"  # This will return an empty result
            cursor.execute(query)

        rows = cursor.fetchall()
        columns = [column[0] for column in cursor.description]
        df = pd.DataFrame.from_records(rows, columns=columns)
        return df

    def table_recon(BillNo ):
        with open('db.json') as app_data:
            d = json.load(app_data)
        cnxn_string = 'Driver={SQL Server};SERVER='+d["db connection"]["server"]+';DATABASE='+d["db connection"]["database"]+';UID='+d["db connection"]["username"]+';PWD='+d["db connection"]["password"]
        cnxn = pyodbc.connect(cnxn_string)
        cursor = cnxn.cursor()

        query = f"SELECT * FROM tbl_SOARecon WHERE BillNo = '{BillNo}'"
        print(query)
        cursor.execute(query)

        rows = cursor.fetchall()
        columns = [column[0] for column in cursor.description]
        df = pd.DataFrame.from_records(rows, columns=columns)
        return df

        # Function to convert complex objects to string representations
    def convert_data(data):
        for item in all_data:
            for key, value in item.items():
                if isinstance(value, (Decimal, datetime)):
                    item[key] = str(value)  # Convert to string for Excel compatibility
        return data
all_data = connection.pandas_dataframe([])

def fetch_claim_data(DMS_ID):
    # URL with dynamic claim number
    url = f"https://dms.emastpa.com.my/emasdms/api/v1/Billrecord?select=claimnumber&maxSize=5&offset=0&where[0][type]=equals&where[0][attribute]=claimnumber&where[0][value]={DMS_ID}"
    
    # Headers with API Key
    headers = {
        'X-Api-Key': '4009add19c9daee9915d3e0d682d56c2'
    }

    # Sending GET request to the API
    response = requests.get(url, headers=headers)

    # Checking if the request was successful (status code 200)
    if response.status_code == 200:
        data = response.json()  # Parsing the JSON response

        if data['total'] > 0:
            DMS_ID = data['list'][0]['id']
            print(f"DMSID: {DMS_ID}")
            #print(data)
            return DMS_ID
        else:
            print(f"No DMSID found for Claim Number: {DMS_ID}")
            return None
    else:
        print(f"Failed to fetch data for Claim Number: {DMS_ID}, Status Code: {response.status_code}")
        return None

# Function to load the session state
def load_session_state():
    global processed_dataframes, extracted_hospitals, recon_data, is_data_extracted, is_data_saved, extracted_sheets, selected_file_path, sheet_df

    session_file = filedialog.askopenfilename(
        title="Select Session File",
        filetypes=[("Pickle files", "*.pkl")]
    )

    if session_file:
        try:
            with open(session_file, "rb") as f:
                state = pickle.load(f)
                
                # Load state data or initialize if not present
                processed_dataframes = state.get("processed_dataframes", [])
                extracted_hospitals = state.get("extracted_hospitals", [])
                recon_data = state.get("recon_data", {})
                extracted_sheets = state.get("extracted_sheets", [])
                selected_file_path = state.get("selected_file_path", None)
                sheet_df = state.get("sheet_df", None)
                is_data_extracted = state.get("is_data_extracted", False)
                is_data_saved = state.get("is_data_saved", True)

            # Check if there's no saved data
            if not processed_dataframes and not extracted_hospitals and not extracted_sheets:
                messagebox.showerror(
                    "No Data", "The selected session file contains no data. Starting a new session."
                )
                start_new_session()
            else:
                messagebox.showinfo(
                    "Session Loaded", "Session data loaded successfully from the selected file."
                )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load session: {e}")
    else:
        messagebox.showinfo(
            "No Selection", "No session file selected. Load session canceled."
        )

# Function to save the session state to a selected file
def save_session_state():
    global is_data_extracted, is_data_saved  # Use global to modify the flags

    session_file = filedialog.asksaveasfilename(
        title="Save Session As",
        defaultextension=".pkl",
        filetypes=[("Pickle files", "*.pkl")]
    )
    if session_file:
        try:
            state = {
                "processed_dataframes": processed_dataframes,
                "extracted_hospitals": extracted_hospitals,
                "extracted_sheets": extracted_sheets
            }
            with open(session_file, "wb") as f:
                pickle.dump(state, f)
            messagebox.showinfo(
                "Save Session", "Session state saved successfully!")

            # Reset the flags after saving
            is_data_extracted = False
            is_data_saved = True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save session: {e}")
    else:
        messagebox.showinfo("No Selection", "Session save canceled.")

# Function to start a new session
def start_new_session():
    result = messagebox.askyesno(
        "Confirm New Session", "Are you sure you want to start a new session? All previous data will be cleared.")

    if result:
        global processed_dataframes, extracted_hospitals, extracted_sheets
        processed_dataframes = []
        extracted_hospitals = []
        extracted_sheets = []
        messagebox.showinfo(
            "New Session", "Starting a new session. All previous data has been cleared.")
    else:
        messagebox.showinfo(
            "New Session", "Continuing with the current session.")
        
# Function to preview data in a new window
def preview_data(df, title="Data Preview"):
    # Create a new CustomTkinter Toplevel window
    preview_window = ctk.CTkToplevel(root)
    preview_window.title(title)

    # This makes the preview window independent from the parent window
    preview_window.transient(root)
    preview_window.grab_set()

    # Set a fixed size or allow resizing
    preview_window.geometry("800x400")
    preview_window.resizable(True, True)

    # Apply a professional theme/style to the Treeview
    style = ttk.Style(preview_window)
    style.theme_use('clam')  # or 'default', 'alt', etc.
    style.configure("Treeview",
                    background="#f0f0f0",
                    foreground="black",
                    rowheight=25,
                    fieldbackground="#f0f0f0")
    style.configure("Treeview.Heading",
                    background="#003366",
                    foreground="white",
                    font=('Arial', 10, 'bold'))
    style.map('Treeview', background=[
        ('selected', '#003366')], foreground=[('selected', 'white')])

    # Create a Treeview widget
    tree = ttk.Treeview(preview_window, selectmode="browse", style="Treeview")
    tree.pack(fill='both', expand=True)

    # Define columns
    tree["column"] = list(df.columns)
    tree["show"] = "headings"

    # Add column headings with formatting
    for col in df.columns:
        tree.heading(col, text=col, anchor=ctk.CENTER)
        tree.column(col, anchor=ctk.CENTER, width=120)

    # Add rows with striped background for better readability
    for index, row in df.iterrows():
        tree.insert("", "end", values=list(row), tags=(
            'oddrow' if index % 2 == 0 else 'evenrow',))

    # Style for alternating row colors
    tree.tag_configure('oddrow', background='#e6f2ff')
    tree.tag_configure('evenrow', background='#ffffff')

    # Add a full vertical scrollbar
    scrollbar = ctk.CTkScrollbar(preview_window, orientation="vertical", command=tree.yview)

    # Create the horizontal scrollbar
    horizontal_scrollbar = ctk.CTkScrollbar(preview_window, orientation='horizontal', command=tree.xview)

    # Configure the Treeview to work with the scrollbars
    tree.configure(yscrollcommand=scrollbar.set)
    tree.configure(xscrollcommand=horizontal_scrollbar.set)

    # Use grid layout to place the Treeview and scrollbars
    tree.grid(row=0, column=0, sticky='nsew')
    scrollbar.grid(row=0, column=1, sticky='ns')
    horizontal_scrollbar.grid(row=1, column=0, sticky='ew')

    # Make the Treeview and Scrollbar expand to fill the window
    preview_window.grid_rowconfigure(0, weight=1)
    preview_window.grid_columnconfigure(0, weight=1)

    # Handle closing of the preview window
    preview_window.protocol("WM_DELETE_WINDOW", preview_window.destroy)

def letter_to_index(col_letter):
    col_letter = col_letter.upper()
    index = 0
    for char in col_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

# Function to safely add data to all lists
def add_extracted_data(sheet_name, extracted_data):
    extracted_sheets.add(sheet_name)
    processed_dataframes.append(extracted_data)

def remove_extracted_data(sheet_name):
    global extracted_sheets, processed_dataframes, extracted_hospitals

    try:
        # Ensure the sheet exists in extracted_sheets
        if sheet_name in extracted_sheets:
            # Convert set to list to maintain index order and get index of the sheet
            index = extracted_sheets.index(sheet_name)

            # Remove from extracted_sheets
            extracted_sheets.pop(index)

            # Remove corresponding data from processed_dataframes and extracted_hospitals
            if index < len(processed_dataframes):
                processed_dataframes.pop(index)
            else:
                print(f"Warning: No corresponding dataframe at index {index}.")

            if index < len(extracted_hospitals):
                extracted_hospitals.pop(index)
            else:
                print(f"Warning: No corresponding hospital at index {index}.")

            print(f"Successfully removed data for '{sheet_name}' at index {index}.")
        else:
            print(f"Sheet '{sheet_name}' not found in extracted_sheets.")
    
    except IndexError:
        print("Error: Index out of range.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def extract_data(bill_date_letter, bill_number_letter, bill_amount_letter, file_path, sheet_name):
    global extracted_sheets, processed_dataframes, extracted_hospitals, is_data_extracted, is_data_saved

    # Check if the same column is selected for multiple fields
    selected_columns = [bill_date_letter, bill_number_letter, bill_amount_letter]

    # If there are duplicates, alert the user and return
    if len(selected_columns) != len(set(selected_columns)):
        messagebox.showerror("Selection Error",
                             "You have selected the same column for more than one field. Please choose different columns for each field.")
        return  # Stop the extraction process

    try:
        # Load the data from the selected sheet
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        # Convert column letters to indices
        bill_date_idx = letter_to_index(bill_date_letter)
        bill_number_idx = letter_to_index(bill_number_letter)
        bill_amount_idx = letter_to_index(bill_amount_letter)

        # Skip the header row by slicing the DataFrame
        data_without_header = df.iloc[1:]

        # Extract data using indices
        extracted_data = data_without_header.iloc[:, [bill_date_idx, bill_number_idx, bill_amount_idx]]
        extracted_data.columns = ['Bill Date', 'Bill Number', 'Bill Amount']

        # Convert 'Bill Date' to datetime and remove time component
        try:
            extracted_data['Bill Date'] = pd.to_datetime(extracted_data['Bill Date'], dayfirst=True).dt.date
        except Exception as date_error:
            raise ValueError(f"Date conversion failed: {date_error}")

        # Bill Number remains unchanged; no date conversion or time removal applied
        extracted_data['Bill Number'] = extracted_data['Bill Number'].apply(str)

        # Handle the second bill number column if enabled
        if second_bill_number_var.get():
            second_bill_number_idx = letter_to_index(second_bill_number_combobox.get())
            extracted_data['Second Bill Number'] = data_without_header.iloc[:, second_bill_number_idx]

            # Combine 'Bill Number' and 'Second Bill Number' with conditions
            extracted_data['Bill Number'] = extracted_data.apply(
                lambda row: combine_bill_numbers(str(row['Bill Number']), str(row['Second Bill Number'])),
                axis=1
            )
            extracted_data.drop(columns=['Second Bill Number'], inplace=True)

        # Clean Bill Number if checkbox is selected
        if remove_non_numeric.get():
            extracted_data['Bill Number'] = extracted_data['Bill Number'].apply(
                lambda x: re.sub(r'[a-zA-Z]', '', str(x))
            )

        # Check if the sheet is already extracted
        if sheet_name in extracted_sheets:
            user_choice = messagebox.askyesno("Duplicate Sheet Found",
                                              f"The sheet '{sheet_name}' has already been extracted. Do you want to replace the existing data?")
            if user_choice:
                remove_extracted_data(sheet_name)  # Remove the existing data
            else:
                messagebox.showinfo("Data Kept", f"The existing data for sheet '{sheet_name}' has been kept.")
                return

        # Append to the lists
        extracted_sheets.append(sheet_name)

        # Fetch SQL data using the unique Bill Numbers from the extracted data
        bill_numbers = extracted_data['Bill Number'].unique().tolist()
        sql_data = connection.pandas_dataframe(bill_numbers)
        # print(bill_numbers['ClaimID'])

        # Merge extracted data with SQL data
        combined_data = pd.merge(extracted_data, sql_data, how='left', left_on='Bill Number', right_on='BillNo')
        # print(combined_data['ClaimID'])

        # Ensure 'ClaimID' is available in the combined SQL data for fetching DMSID
        if 'ClaimID' in combined_data.columns:
            combined_data['ClaimID'] = combined_data['ClaimID'].astype(pd.Int64Dtype()) 
            # Fetch DMS_ID for each ClaimID and store them in a list
            DMS_ID = combined_data['ClaimID'].apply(fetch_claim_data)
            
            # Append the DMSID list as a new column in combined_data
            combined_data['DMSID'] = DMS_ID
        else:
            raise ValueError("SQL data does not contain 'ClaimID' column.")

        # Ensure 'Recon By', 'Recon Date', and 'SOA Date' columns are added
        combined_data['Recon By'] = recon_data.get('Recon By', '')
        combined_data['Recon Date'] = recon_data.get('Recon Date', '')
        combined_data['SOA Date'] = recon_data.get('SOA Date', '')

        # Convert 'bill number' and 'bill no' to string (no issues expected here)
        combined_data['Bill Number'] = combined_data['Bill Number'].astype(str)
        combined_data['BillNo'] = combined_data['BillNo'].astype(str)
        combined_data['BillNo'] = combined_data['BillNo'].replace('NaN', '')

        # Handle non-numeric values for 'bill amount'
        combined_data['Bill Amount'] = pd.to_numeric(combined_data['Bill Amount'], errors='coerce')  # Converts invalid parsing to NaN
        combined_data['Bill Amount'] = combined_data['Bill Amount'].apply(lambda x: f"{x:.2f}" if pd.notna(x) else '')

        # Handle non-numeric values for 'claim id' and 'provider id'
        # combined_data['ClaimID'] = pd.to_(combined_data['ClaimID'], errors='coerce')  # Converts invalid parsing to NaN
        # combined_data['ProviderID'] = pd.to_numeric(combined_data['ProviderID'], errors='coerce')  # Converts invalid parsing to NaN

        # Convert types explicitly
        combined_data['ClaimID'] = combined_data['ClaimID'].astype(pd.Int64Dtype())  # Use 'Int64Dtype' to keep NaN values
        combined_data['ProviderID'] = combined_data['ProviderID'].astype(pd.Int64Dtype())  # Use 'Int64Dtype' to keep NaN values

        # Create the 'reconref' column
        combined_data['Reconrefno'] = combined_data.apply(lambda row: create_reconref(row['SOA Date'], row['Recon Date']), axis=1)

        #Create row for missing values 
        combined_data['ModifyDate'] = ''
        combined_data['ModifyBy'] = ''
        combined_data['ReconRemarks'] = ''
        #combined_data['DMSID'] = ''
        combined_data['CreateDate'] = datetime.now().date()     # only gets date 

        # Add combined data to processed dataframes
        if len(processed_dataframes) == 0: 
            processed_dataframes.append(combined_data)
        else:
            # processed_dataframes[0] = processed_dataframes[0].append(combined_data, ignore_index=True)
            processed_dataframes[0] = pd.concat([processed_dataframes[0], combined_data], ignore_index=True)
        
        print(processed_dataframes)

        is_data_extracted = True
        is_data_saved = False

        # Show preview of the combined data
        preview_data(combined_data, title=f"Preview")
        messagebox.showinfo("Success", f"Data extracted from {os.path.basename(file_path)} successfully!")

    except ValueError as ve:
        messagebox.showerror("Input Error", f"Input Error: {ve}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred during extraction: {e}")


def edit_data():
    # Create a copy of the DataFrame for editing
    if isinstance(processed_dataframes, list) and all(isinstance(df, pd.DataFrame) for df in processed_dataframes):
        df = pd.concat(processed_dataframes, ignore_index=True).copy()
        # print(len(processed_dataframes))
        # print(processed_dataframes[0])
        # print(type(processed_dataframes[0]))
    else:
        print("Invalid input: processed_dataframes should be a list of DataFrames.")
        return

    # Variable to track the original bill number for saving
    original_bill_number = None

    def on_double_click(event, tree):
        try:
            item = tree.selection()[0]
            column = tree.identify_column(event.x)
            col = int(column.split('#')[-1]) - 1
            if col == 1:  # Allow editing only the Bill Number in the second column
                start_edit(tree, item, col)
        except IndexError:
            print("No item selected")

    def start_edit(tree, item, col):
        nonlocal original_bill_number  # Access the variable to track the original bill number
        value = tree.item(item, 'values')[col]
        original_bill_number = tree.item(item, 'values')[1]  # Save the original bill number
        edit_window = ctk.CTkToplevel(tree)
        edit_window.title("Edit Bill Number")
        edit_window.geometry("200x100")

        edit_entry = ctk.CTkEntry(edit_window, width=150)
        edit_entry.insert(0, value)
        edit_entry.grid(row=0, column=0, padx=10, pady=10)
        edit_entry.focus()

        save_button = ctk.CTkButton(edit_window, text="Save", 
                                    command=lambda: save_edit(tree, item, col, edit_entry.get(), edit_window))
        save_button.grid(row=1, column=0, padx=5, sticky=tk.W)
        cancel_button = ctk.CTkButton(edit_window, text="Cancel", command=edit_window.destroy)
        cancel_button.grid(row=1, column=1, padx=5, sticky=tk.E)

    def save_edit(tree, item, col, new_value, edit_window):
        values = list(tree.item(item, 'values'))
        old_billnumber = values[1]
        # original_bill_number = values[1]  # The old bill number before editing
        new_list = values
        new_list[col] = new_value
        
        sql_data = connection.pandas_dataframe([new_value])

        # Update the local copy of the DataFrame
        for original_df in processed_dataframes:
            if len(sql_data) == 1:
                new_list[3] = sql_data["ClaimID"].astype(pd.Int64Dtype()).tolist()[0]
                new_list[4] = sql_data["BillNo"].astype(str).tolist()[0]
                new_list[5] = sql_data["ProviderID"].astype(pd.Int64Dtype()).tolist()[0]
                new_list[6] = sql_data["ProviderName"].tolist()[0]
                new_list[7] = sql_data["InvoiceNo"].astype(str).tolist()[0]
                new_list[8] = sql_data["ClaimAmount"].apply(pd.to_numeric).tolist()[0]
                new_list[9] = sql_data["CompanyName"].tolist()[0]
                new_list[10] = sql_data["ClientPaidDate"].apply(pd.to_datetime).tolist()[0]
                new_list[11] = sql_data["PayOrderNo"].tolist()[0]
                new_list[12] = sql_data["Status"].tolist()[0]


                original_df.loc[original_df['Bill Number'] == old_billnumber, 'Bill Number'] = new_value
                
                original_df.loc[original_df['Bill Number'] == new_value, 'ClaimID'] = sql_data["ClaimID"].astype(pd.Int64Dtype()).tolist()[0]
                original_df.loc[original_df['Bill Number'] == new_value, 'BillNo'] = sql_data["BillNo"].astype(str).tolist()[0]
                original_df.loc[original_df['Bill Number'] == new_value,'ProviderID'] = sql_data["ProviderID"].astype(pd.Int64Dtype()).tolist()[0]
                original_df.loc[original_df['Bill Number'] == new_value,'ProviderName'] = sql_data["ProviderName"].tolist()[0]
                original_df.loc[original_df['Bill Number'] == new_value,'InvoiceNo'] = sql_data["InvoiceNo"].astype(str).tolist()[0]
                original_df.loc[original_df['Bill Number'] == new_value,'ClaimAmount'] = sql_data["ClaimAmount"].apply(pd.to_numeric).tolist()[0]
                original_df.loc[original_df['Bill Number'] == new_value,'CompanyName'] = sql_data["CompanyName"].tolist()[0]
                original_df.loc[original_df['Bill Number'] == new_value,'ClientPaidDate'] = sql_data["ClientPaidDate"].apply(pd.to_datetime).tolist()[0]
                original_df.loc[original_df['Bill Number'] == new_value,'PayOrderNo'] = sql_data["PayOrderNo"].tolist()[0]
                original_df.loc[original_df['Bill Number'] == new_value,'Status'] = sql_data["Status"].tolist()[0]

                DMS_ID = original_df.loc[original_df['Bill Number'] == new_value, 'ClaimID'].apply(fetch_claim_data)
                original_df.loc[original_df['Bill Number'] == new_value,'DMSID'] = DMS_ID

                new_list[13] = DMS_ID
            
            elif len(sql_data) == 0:
                print("jjj")
                new_list[3] = None
                new_list[4] = None
                new_list[5] = None
                new_list[6] = None
                new_list[7] = None
                new_list[8] = None
                new_list[9] = None
                new_list[10] = None
                new_list[11] = None
                new_list[12] = None

                original_df.loc[original_df['Bill Number'] == old_billnumber, 'Bill Number'] = new_value
                
                original_df.loc[original_df['Bill Number'] == new_value, 'ClaimID'] = None
                original_df.loc[original_df['Bill Number'] == new_value, 'BillNo'] = None
                original_df.loc[original_df['Bill Number'] == new_value,'ProviderID'] = None
                original_df.loc[original_df['Bill Number'] == new_value,'ProviderName'] = None
                original_df.loc[original_df['Bill Number'] == new_value,'InvoiceNo'] = None
                original_df.loc[original_df['Bill Number'] == new_value,'ClaimAmount'] = None
                original_df.loc[original_df['Bill Number'] == new_value,'CompanyName'] = None
                original_df.loc[original_df['Bill Number'] == new_value,'ClientPaidDate'] = None
                original_df.loc[original_df['Bill Number'] == new_value,'PayOrderNo'] = None
                original_df.loc[original_df['Bill Number'] == new_value,'Status'] = None

                original_df.loc[original_df['Bill Number'] == new_value,'DMSID'] = None
                print(original_df)
                new_list[13] = None
        print(processed_dataframes[0])
            #     print(sql_data["ClaimID"].astype(pd.Int64Dtype()).tolist()[0])
            #     print(original_df.loc[original_df['Bill Number'] == new_value, 'ClaimID'])
            # # processed_dataframes[0].loc[processed_dataframes[0]['Bill Number'] == new_value, "ClaimID"] = sql_data["ClaimID"].astype(pd.Int64Dtype()).tolist()[0]
            # print(processed_dataframes[0].loc[processed_dataframes[0]['Bill Number'] == new_value])
        tree.item(item, values=new_list)
        edit_window.destroy()    

    def add_remarks(tree, remarks_entry, selected_item):
        if selected_item:
            for original_df in processed_dataframes:
                remarks = remarks_entry.get()
                bill_number = tree.item(selected_item, 'values')[1]  # Bill Number is in column 1
                 
                original_df.loc[original_df['Bill Number'] == bill_number, 'ReconRemarks'] = remarks
                print(original_df)
                print(bill_number)
                updated_row = original_df[original_df['Bill Number'] == bill_number].iloc[0]
                print(updated_row)
                print(original_df[original_df['Bill Number'] == bill_number].iloc[0])
                tree.item(selected_item, values=list(updated_row))
            remarks_entry.delete(0, tk.END)

    def on_row_select(event, tree, remarks_entry):
        selected_item = tree.selection()[0]
        remarks = tree.item(selected_item, 'values')[-1]  # Remarks are in the last column
        remarks_entry.delete(0, tk.END)
        remarks_entry.insert(0, remarks)

    def save_changes_to_original():
        # Update the original processed_dataframes with the changes made
        for original_df in processed_dataframes:
            for index, row in df.iterrows():
                # Update the original DataFrame based on edits
                original_df.loc[original_df['Bill Number'] == row['Bill Number'], 'ReconRemarks'] = row['ReconRemarks']
                # If the Bill Number has changed, update it as well
                if row['Bill Number'] != original_bill_number:
                    original_df.loc[original_df['Bill Number'] == original_bill_number, 'Bill Number'] = row['Bill Number']
        print("Changes saved to the original DataFrame.")

    def create_editable_treeview():
        window = ctk.CTkToplevel(root)
        window.title("Editable Treeview Example")

        # Create a frame for the Treeview and scrollbar
        frame = ctk.CTkFrame(window)
        frame.pack(fill=tk.BOTH, expand=True)

        columns = list(df.columns)
        
        # Create the Treeview
        tree = ttk.Treeview(frame, columns=columns, show='headings')
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create the horizontal scrollbar
        horizontal_scrollbar = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
        horizontal_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Attach the scrollbar to the Treeview 
        tree.configure(xscrollcommand=horizontal_scrollbar.set)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        for index, row in df.iterrows():
            values = tuple(row.fillna(''))  # Fill NaNs with empty strings for display
            tree.insert('', tk.END, iid=index, values=values)

        # Remarks entry and buttons
        remarks_entry = ctk.CTkEntry(window, width=300)
        remarks_entry.pack(padx=20, pady=10)
        print(remarks_entry)

        remarks_button = ctk.CTkButton(window, text="Update Remarks", 
                                command=lambda: add_remarks(tree, remarks_entry, 
                                tree.selection()[0] if tree.selection() else None))
        remarks_button.pack(padx=20, pady=10)

        save_button = ctk.CTkButton(window, text="Save Changes", command=save_changes_to_original)
        save_button.pack(padx=20, pady=10)

        # Configure window resizing
        window.grid_rowconfigure(0, weight=1)
        window.grid_columnconfigure(0, weight=1)

        # Bind events
        tree.bind('<Double-1>', lambda event: on_double_click(event, tree))
        tree.bind('<<TreeviewSelect>>', lambda event: on_row_select(event, tree, remarks_entry))

    create_editable_treeview()

class DatabaseManager:
    @staticmethod
    def create_connection(config_file):
        try:
           with open('db.json') as app_data:
            d = json.load(app_data)

            cnxn_string = (
                'Driver={SQL Server};'
                'SERVER=' + d["db connection"]["server"] + ';'
                'DATABASE=' + d["db connection"]["database"] + ';'
                'UID=' + d["db connection"]["username"] + ';'
                'PWD=' + d["db connection"]["password"]
            )
            logging.info("Connection string created successfully.")

            # Establish the connection
            connection = pyodbc.connect(cnxn_string)
            return connection
        except Exception as e:
            logging.error(f"Error creating connection: {e}")
            raise

    @staticmethod
    def prepare_dataframe(df):
        # Convert columns to the appropriate data types
        # df['Bill Date'] = pd.to_datetime(df['Bill Date'])
        # df['ClientPaidDate'] = pd.to_datetime(df['ClientPaidDate'])
        # df['Recon Date'] = pd.to_datetime(df['Recon Date'])
        # df['SOA Date'] = pd.to_datetime(df['SOA Date'])
        # df['Bill Date'] = str(df['Bill Date'])
        # df['ClientPaidDate'] = str(df['ClientPaidDate'])
        # df['Recon Date'] = str(df['Recon Date'])
        # df['SOA Date'] = str(df['SOA Date'])
        # print(df['Bill Date'])

        df['Bill Amount'] = pd.to_numeric(df['Bill Amount'], errors='coerce')
        df['ClaimAmount'] = pd.to_numeric(df['ClaimAmount'], errors='coerce')

        df['ClaimID'] = pd.to_numeric(df['ClaimID'], errors='coerce').astype(pd.Int64Dtype())
        df['ProviderID'] = pd.to_numeric(df['ProviderID'], errors='coerce').astype(pd.Int64Dtype())

        # Ensure the varchar fields are strings
        varchar_fields = [
            'Bill Number', 'BillNo', 'ProviderName', 'InvoiceNo',
            'CompanyName', 'PayOrderNo', 'Status', 'Recon By', 'Reconrefno'
        ]
        for field in varchar_fields:
            df[field] = df[field].astype(str)

        return df

    @staticmethod
    def insert_data(connection, row):
        if connection is None:
            raise RuntimeError("Connection is not established.")

        try:
            with connection.cursor() as cursor:
                logging.debug(f"Inserting row: {row}")  # Log the row being inserted
                cursor.execute('''INSERT INTO tbl_SOARecon ( CreateDate, BillNo,
                    BillAmount, ReconReferenceNo, ReconDate, SOADate,  ProviderID, ClaimID,
                    DMSID, ReconBy, ModifyDate, ModifyBy, ReconRemarks,
                    ReconStatus
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
                    #str(row['Bill Date']),
                    str(row['CreateDate']),
                    str(row['Bill Number']) if str(row['Bill Number']) != 'nan' else '',
                    #str(row['BillNo']) if str(row['BillNo']) != 'nan' else '',
                    str(row['Bill Amount']),
                    str(row['Reconrefno']),
                    str(row['Recon Date']).replace("/","-"),
                    str(row['SOA Date']).replace("/","-"),
                    str(row['ProviderID']) if str(row['ProviderID']) != 'None' else '0',
                    str(row['ClaimID']) if str(row['ClaimID']) != 'None' else '0',
                    str(row['DMSID']),
                    str(row['Recon By']),
                    str(row['ModifyDate']),
                    str(row['ModifyBy']),
                    str(row['ReconRemarks']),
                    str(row['Status']) if str(row['Status']) != 'nan' else '',    
                    #str(row['ProviderName']) if str(row['ProviderName']) != 'nan' else '',
                    #str(row['InvoiceNo']) if str(row['InvoiceNo']) != 'nan' else '',
                    #str(row['ClaimAmount']) if str(row['ClaimAmount']) != 'nan' else '0',
                    #str(row['CompanyName']) if str(row['CompanyName']) != 'nan' else '',
                    #str(row['ClientPaidDate'].date()).replace("/","-") if str(row['ClientPaidDate']) != 'NaT' else '',
                    #str(row['PayOrderNo']) if str(row['PayOrderNo']) != 'nan' else '',      
                ))
            # Commit the transaction
            connection.commit()
            logging.info("Data inserted successfully.")
        except Exception as e:
            logging.error(f"Error inserting data: {e}")
            connection.rollback()  # Rollback in case of error

    @staticmethod
    def update_data(connection, query, updatedata):
        if connection is None:
            raise RuntimeError("Connection is not established.")

        try:
            with connection.cursor() as cursor:
                # logging.debug(f"Updating row: {row}")  # Log the row being inserted
                cursor.execute(query, updatedata)
            # Commit the transaction
            connection.commit()
            print("Data updated successfully.")
            logging.info("Data updated successfully.")
        except Exception as e:
            logging.error(f"Error updating data: {e}")
            connection.rollback()  # Rollback in case of error

    @staticmethod
    def close_connection(connection):
        if connection:
            connection.close()
            logging.info("Connection closed.")


def insert_data_to_db():
    global processed_dataframes, is_data_extracted

    if is_data_extracted and processed_dataframes:
        # try:
        # Initialize and create a connection
        config_file = 'path_to_config_file.json'  # Adjust as necessary
        db_manager = DatabaseManager()
        connect = db_manager.create_connection(config_file)
        # Insert each processed DataFrame into the database
        for df in processed_dataframes:
            df = DatabaseManager.prepare_dataframe(df)  # Prepare the DataFrame
            # print(df)
            df = df.to_dict('records')
            
            for row in df:
                update = False 
                updatelist = {}
                billno = row['Bill Number']

                if billno != "":
                    df2 = connection.table_recon(str(row['Bill Number']))
                    if len(df2) > 0:
                        print(df2["BillAmount"].tolist()[0])
                        if df2["BillAmount"].tolist()[0] != Decimal(str(row['Bill Amount'])).quantize(Decimal('0.00')):
                            update  = True
                            updatelist["Bill Amount"] =  'BillAmount'
                        elif str(df2["ProviderID"].tolist()[0]) != str(row['ProviderID']) if str(row['ProviderID']) != 'None' else '0':
                            update  = True
                            updatelist["ProviderID"] = 'ProviderID'
                            print(str(df2["ProviderID"].tolist()[0]))
                            print(str(row['ProviderID']) if str(row['ProviderID']) != 'None' else '0')
                        elif str(df2["ClaimID"].tolist()[0]) != str(row['ClaimID']) if str(row['ClaimID']) != 'None' else '0':
                            update = True 
                            updatelist["ClaimID"] = 'ClaimID'
                        elif df2["DMSID"].tolist()[0] != str(row['DMSID']):
                            print(df2["DMSID"].tolist()[0])
                            print(str(row['DMSID']))
                            update = True 
                            updatelist["DMSID"] = 'DMSID'
                        elif df2["ReconDate"].tolist()[0] != row['Recon Date']: 
                            update  = True
                            updatelist["Recon Date"] = 'ReconDate'
                        elif df2["ReconRemarks"].tolist()[0] !=  str(row['ReconRemarks']): 
                            update  = True 
                            updatelist["ReconRemarks"] = 'ReconRemarks'
                        elif df2["ReconStatus"].tolist()[0] != str(row['Status']) if str(row['Status']) != 'nan' else '': 
                            update = True
                            updatelist["Status"] ='ReconStatus'
                        
                        updatelist["Recon By"] ='ModifyBy'
                        if update:
                            str1 = "update tbl_SOARecon set "
                            str2 = ""
                            str3 = f" where BillNo = ?"
                            updatedata = []
                            for columname in updatelist:
                                str2 = f"{str2} [{updatelist[columname]}] = ?,"
                                updatedata.append(str(row[columname]).replace('/','-') if str(row[columname]) != 'None' else '0')
                            str2 = str2 + ' ModifyDate = getdate()'
                            query = str1 + str2 + str3
                            updatedata.append(billno)
                            updatedata = tuple(updatedata)
                            print(query)
                            print(updatedata)
                            db_manager.update_data(connect, query, updatedata)
                        # print(update)
                        # print(updatelist)
                    else: 
                    # if str(row['Bill Number']) =='1 087761':
                        # print(str(row['SOA Date']).replace("/","-"))
                        # print(str(row['Recon Date']).replace("/","-"))
                        # print(str(row['ClientPaidDate']).replace("/","-"))
                    # print(type(df))
                    # print((
                    # str(row['CreateDate']),
                    # str(row['Bill Number']) if str(row['Bill Number']) != 'nan' else '',
                    # #str(row['BillNo']) if str(row['BillNo']) != 'nan' else '',
                    # str(row['Bill Amount']),
                    # str(row['Reconrefno']),
                    # str(row['Recon Date']).replace("/","-"),
                    # str(row['SOA Date']).replace("/","-"),
                    # str(row['ProviderID']) if str(row['ProviderID']) != 'None' else '0',
                    # str(row['ClaimID']) if str(row['ClaimID']) != 'None' else '0',
                    # str(row['DMSID']),
                    # str(row['Recon By']),
                    # str(row['ModifyDate']),
                    # str(row['ModifyBy']),
                    # str(row['ReconRemarks']),
                    # str(row['Status']) if str(row['Status']) != 'nan' else '',    
                    # ))
                    # print(row)
                #     print((
                #     row['Bill Date'] if row['Bill Date'] is not None or row['Bill Date'] != 'NaN'.lower() or row['Bill Date'] !=  '<NA>' or row['Bill Date'] != 'NaT'  else '',
                #     row['Bill Number'] if row['Bill Number'] is not None or row['Bill Number'] != 'NaN'.lower() or row['Bill Number'] !=  '<NA>' or row['Bill Number'] != 'NaT' else '',
                #     row['Bill Amount'] if row['Bill Amount'] is not None or row['Bill Amount'] != 'NaN'.lower() or row['Bill Amount'] !=  '<NA>' or row['Bill Amount'] != 'NaT'  else '',
                #     row['ClaimID'] if row['ClaimID'] is not None or row['ClaimID'] != 'NaN'.lower() or row['ClaimID'] !=  '<NA>' or row['ClaimID'] != 'NaT'  else '',
                #     row['BillNo'] if row['BillNo'] is not None or row['BillNo'] != 'NaN'.lower() or row['BillNo'] !=  '<NA>' or row['BillNo'] != 'NaT'  else '',
                #     row['ProviderID'] if row['ProviderID'] is not None or row['ProviderID'] != 'NaN'.lower() or row['ProviderID'] !=  '<NA>' or row['ProviderID'] != 'NaT'  else '',
                #     row['ProviderName'] if row['ProviderName'] is not None or row['ProviderName'] != 'NaN'.lower() or row['ProviderName'] !=  '<NA>' or row['ProviderName'] != 'NaT'  else '',
                #     row['InvoiceNo'] if row['InvoiceNo'] is not None or row['InvoiceNo'] != 'NaN'.lower() or row['InvoiceNo'] !=  '<NA>' or row['InvoiceNo'] != 'NaT'  else '',
                #     row['ClaimAmount'] if row['ClaimAmount'] is not None or row['ClaimAmount'] != 'NaN'.lower() or row['ClaimAmount'] !=  '<NA>' or row['ClaimAmount'] != 'NaT'  else '',
                #     row['CompanyName'] if row['CompanyName'] is not None or row['CompanyName'] != 'NaN'.lower() or row['CompanyName'] !=  '<NA>' or row['CompanyName'] != 'NaT'  else '',
                #     row['ClientPaidDate'] if row['ClientPaidDate'] is not None or row['ClientPaidDate'] != 'NaN'.lower() or row['ClientPaidDate'] !=  '<NA>' or row['ClientPaidDate'] != 'NaT'  else '',
                #     row['PayOrderNo'] if row['PayOrderNo'] is not None or row['PayOrderNo'] != 'NaN'.lower() or row['PayOrderNo'] !=  '<NA>' or row['PayOrderNo'] != 'NaT'  else '',
                #     row['Status'] if row['Status'] is not None or row['Status'] != 'NaN'.lower() or row['Status'] !=  '<NA>' or row['Status'] != 'NaT'  else '',
                #     row['Recon By'] if row['Recon By'] is not None or row['Recon By'] != 'NaN'.lower() or row['Recon By'] !=  '<NA>' or row['Recon By'] != 'NaT'  else '',
                #     row['Recon Date'] if row['Recon Date'] is not None or row['Recon Date'] != 'NaN'.lower() or row['Recon Date'] !=  '<NA>' or row['Recon Date'] != 'NaT'  else '',
                #     row['SOA Date'] if row['SOA Date'] is not None or str(row['SOA Date']).lower() != 'nan' or row['SOA Date'] !=  '<NA>' or row['SOA Date'] != 'NaT'  else '',
                #     row['Reconrefno']if row['Reconrefno'] is not None or row['Reconrefno'] != 'NaN'.lower() or row['Reconrefno'] !=  '<NA>' or row['Reconrefno'] != 'NaT'  else ''
                # ))
                        db_manager.insert_data(connect, row)
            
        messagebox.showinfo("Success", "Data inserted into the database successfully!")
        # except Exception as e:
        #     messagebox.showerror("Error", f"An error occurred during data insertion: {e}")
        # finally:
        #     db_manager.close_connection(connect)
    else:
        messagebox.showwarning("No Data", "No data to insert. Please extract data first.")

def start_extraction():
    global selected_file_path
    bill_date_letter = bill_date_combobox.get()
    bill_number_letter = bill_number_combobox.get()
    bill_amount_letter = bill_amount_combobox.get()
    file_path = selected_file_path  # Use the global variable instead of entry field
    sheet_name = sheet_selection_combobox.get()

    if not file_path:
        messagebox.showwarning(
            "No File", "Please select a file to extract data from.")
        return

    if not all([bill_date_letter, bill_number_letter, bill_amount_letter]):
        messagebox.showwarning("Missing Information",
                               "Please fill in all required fields.")
        return

    if not sheet_name:
        messagebox.showwarning("No Sheet Selected",
                               "Please select a sheet to extract data from.")
        return

    try:
        extract_data(bill_date_letter, bill_number_letter,
                     bill_amount_letter, file_path, sheet_name)

        # Clear the column fields after successful extraction, but keep the file path and sheet selection
        bill_date_combobox.set('')  # Reset the combobox selection
        bill_number_combobox.set('')  # Reset the combobox selection
        bill_amount_combobox.set('')  # Reset the combobox selection
        second_bill_number_combobox.set('')  # Reset the combobox selection

        # Reset checkboxes to their default states
        second_bill_number_var.set(False)
        remove_non_numeric.set(False)
        second_bill_number_combobox.config(state="disabled")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def browse_file():
    global selected_file_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:  # Only update the entry if a file was selected
        selected_file_path = file_path
        entry_file_path.delete(0, ctk.END)
        entry_file_path.insert(0, file_path)

        # Populate the sheet selection combobox
        try:
            xl = pd.ExcelFile(file_path)
            sheets = xl.sheet_names
            sheet_selection_combobox['values'] = sheets
            sheet_selection_combobox.set(sheets[0] if sheets else "")

            # Automatically load the first sheet
            if sheets:
                load_sheet(file_path, sheets[0])
                update_column_suggestions()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheets: {e}")

def on_sheet_selected(event):
    sheet_name = sheet_selection_combobox.get()
    if sheet_name and selected_file_path:
        load_sheet(selected_file_path, sheet_name)
        update_column_suggestions()

def format_bill_amount(amount):
    """
    Format the bill amount to ensure it has no more than 2 decimal points.
    """
    try:
        # Convert amount to float and round to 2 decimal places
        formatted_amount = round(float(amount), 2)
        return f"{formatted_amount:.2f}"
    except ValueError:
        # If conversion fails, return a default value or handle the error as needed
        return "0.00"

def remove_letters_from_number(bill_number):
    """
    Remove all letters from a bill number but keep non-numeric characters.
    """
    return re.sub(r'[a-zA-Z]', '', bill_number)

def remove_numbers_after_decimal(bill_number):
    """
    Remove all numbers after the decimal point in a bill number.
    """
    return re.sub(r'\.\d+', '', bill_number)

def combine_bill_numbers(bill_number_1, bill_number_2, remove_letters=False):
    """
    Combine two bill numbers based on the following rules:
    1. If bill_number_1's numeric value is '1', combine with a leading '0' between.
    2. If bill_number_1's numeric value is '2', combine without any separator.
    3. If remove_letters is True, remove letters from both bill numbers before combining.
    4. Remove numbers after the decimal points from both bill numbers before combining.
    """
    # Remove numbers after the decimal point
    bill_number_1 = remove_numbers_after_decimal(bill_number_1)
    bill_number_2 = remove_numbers_after_decimal(bill_number_2)

    if remove_letters:
        # Clean both bill numbers by removing letters
        bill_number_1 = remove_letters_from_number(bill_number_1)
        bill_number_2 = remove_letters_from_number(bill_number_2)

    # Determine if cleaned bill_number_1 is numeric and check its value
    bill_number_1_value = remove_letters_from_number(bill_number_1)

    if bill_number_1_value == '1':
        combined = f"{bill_number_1}0{bill_number_2}"
    elif bill_number_1_value == '2':
        combined = f"{bill_number_1} {bill_number_2}"
    else:
        combined = f"{bill_number_1}/{bill_number_2}"

    return combined

def remove_selected_data():
    if not extracted_sheets:
        messagebox.showwarning("No Data", "No sheets available to show.")
        return

    checklist_window = ctk.CTkToplevel(root)
    checklist_window.title("Sheet Checklist")

    def update_listbox(search_term=""):
        listbox.delete(0, ctk.END)
        for sheet in extracted_sheets:
            if search_term.lower() in sheet.lower():
                listbox.insert(ctk.END, sheet)

    def on_search(event):
        search_term = search_entry.get()
        update_listbox(search_term)

    search_frame = ctk.CTkFrame(checklist_window)
    search_frame.pack(padx=10, pady=10)

    search_label = ctk.CTkLabel(search_frame, text="Search:")
    search_label.pack(side=ctk.LEFT)

    search_entry = ctk.CTkEntry(search_frame)
    search_entry.pack(side=ctk.LEFT, fill=ctk.X, expand=True)
    search_entry.bind("<KeyRelease>", on_search)

    listbox = Listbox(
        checklist_window, selectmode="multiple", height=10, width=50)
    listbox.pack(padx=10, pady=10)

    update_listbox()  # Initial listbox population

    # Button to remove selected sheets
    ctk.CTkButton(checklist_window, text="Remove Selected",
                  command=lambda: remove_selected(listbox, checklist_window)).pack(pady=5)
    ctk.CTkButton(checklist_window, text="Close",
                  command=checklist_window.destroy).pack(pady=10)

def update_listbox(listbox, search_term=""):
    listbox.delete(0, ctk.END)
    for sheet in extracted_sheets:
        if search_term.lower() in sheet.lower():
            listbox.insert(ctk.END, sheet)

def remove_selected(listbox, checklist_window):
    selected_indices = listbox.curselection()

    if not selected_indices:
        messagebox.showwarning("No Selection", "Please select sheets to remove.")
        return

    selected_sheets = [listbox.get(i) for i in selected_indices]

    global processed_dataframes, extracted_sheets, extracted_hospitals

    extracted_sheets_list = list(extracted_sheets)  # Convert set to list once

    for sheet in selected_sheets:
        if sheet in extracted_sheets:
            index = extracted_sheets_list.index(sheet)
            
            # Ensure index is within range for all lists
            if index < len(processed_dataframes) and index < len(extracted_hospitals):
                print(f"Removing sheet: {sheet} at index: {index}")
                del processed_dataframes[index]
                del extracted_hospitals[index]
                extracted_sheets.remove(sheet)
            else:
                print(f"Index out of range for sheet: {sheet}")
                messagebox.showerror("Error", f"Unable to remove '{sheet}', index out of range.")
        else:
            print(f"Sheet '{sheet}' not found in extracted_sheets.")

    update_listbox(listbox)
    checklist_window.destroy()



# Function to compile all the processed files into one Excel sheet
def compile_files():
    if not processed_dataframes:
        messagebox.showwarning(
            "No Data", "No files processed yet. Please extract data first.")
        return

    # Add sign-in information to each DataFrame
    for df in processed_dataframes:
        for key, value in recon_data.items():
            df[key] = value

    combined_data = pd.concat(processed_dataframes, ignore_index=True)

    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save Compiled File As"
    )

    if output_file:
        try:
            combined_data.to_excel(output_file, index=False)
            hospital_log_file = os.path.splitext(
                output_file)[0] + "_Hospitals.txt"
            with open(hospital_log_file, 'w') as file:
                file.write("Extracted Hospitals:\n")
                file.write("\n".join(extracted_hospitals))
            messagebox.showinfo("Success", f"All files compiled and saved to {output_file} successfully!\n"
                                           f"Hospital names saved to {hospital_log_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# Function to prompt user for session choice
def prompt_session_choice():
    choice = messagebox.askquestion("Session Choice",
                                    "Do you want to load a previous session? (Click 'No' to start a new session)")

    if choice == 'yes':
        load_session_state()
    else:
        start_new_session()

# Function to preview the compiled data
def preview_compiled_data():
    if not processed_dataframes:
        messagebox.showwarning(
            "No Data", "No data available to preview. Please extract data first.")
        return

    combined_data = pd.concat(processed_dataframes, ignore_index=True)
    preview_data(combined_data, title="Compiled Data Preview")

# Function to handle "Enter" key press to move to the next entry
def on_enter_key(event, current_entry, next_entry):
    current_entry.focus_set()  # Move focus to the current entry
    if next_entry:
        next_entry.focus_set()  # Move focus to the next entry

# Function to bind arrow keys to navigation
def on_arrow_key(event):
    current_widget = root.focus_get()
    entry_list = [bill_date_combobox, bill_number_combobox,
                  bill_amount_combobox, entry_file_path]
    if event.keysym == 'Up':
        prev_index = entry_list.index(current_widget) - 1
        if prev_index >= 0:
            entry_list[prev_index].focus_set()
    elif event.keysym == 'Down':
        next_index = entry_list.index(current_widget) + 1
        if next_index < len(entry_list):
            entry_list[next_index].focus_set()

# Function to enable or disable the second bill number combobox
def toggle_second_bill_number():
    if second_bill_number_var.get():
        second_bill_number_combobox.configure(state="normal")
    else:
        second_bill_number_combobox.configure(state="disabled")

        # Ensure this function is used correctly
        checkbox_second_bill_number.configure(command=toggle_second_bill_number)

def check_list():
    if not extracted_sheets:
        messagebox.showinfo("No Data", "No sheets have been extracted yet.")
        return

    # Create a new Toplevel window
    check_window = ctk.CTkToplevel()
    check_window.title("Extracted Sheets List")

    # Configure columns and rows to expand
    check_window.grid_rowconfigure(0, weight=0)
    check_window.grid_rowconfigure(1, weight=1)
    check_window.grid_rowconfigure(2, weight=0)
    check_window.grid_columnconfigure(0, weight=1)
    check_window.grid_columnconfigure(1, weight=1)

    # Label for search
    label_search = ctk.CTkLabel(check_window, text="Search Sheet:")
    label_search.grid(row=0, column=0, padx=10, pady=5, sticky="e")

    # Entry for search
    entry_search = ctk.CTkEntry(check_window)
    entry_search.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

    # Listbox with scrollbar using tk.Listbox instead
    frame_listbox = ctk.CTkFrame(check_window)
    frame_listbox.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

    # Scrollbar
    scrollbar = ctk.CTkScrollbar(frame_listbox, orientation="vertical")
    scrollbar.pack(side="right", fill="y")

    # Standard tk.Listbox wrapped in CTkFrame for consistency
    listbox = tk.Listbox(frame_listbox, yscrollcommand=scrollbar.set, height=10)
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.configure(command=listbox.yview)

    # Function to populate the Listbox
    def populate_listbox(sorted_sheets, sorted_dataframes):
        listbox.delete(0, tk.END)
        for sheet in sorted_sheets:
            listbox.insert(tk.END, sheet)

    # Function to sort sheets and dataframes by sheet names in ascending order
    def sort_ascending():
        sorted_pairs = sorted(zip(extracted_sheets, processed_dataframes))
        sorted_sheets, sorted_dataframes = zip(*sorted_pairs)
        populate_listbox(sorted_sheets, sorted_dataframes)

    # Function to sort sheets and dataframes by sheet names in descending order
    def sort_descending():
        sorted_pairs = sorted(zip(extracted_sheets, processed_dataframes), reverse=True)
        sorted_sheets, sorted_dataframes = zip(*sorted_pairs)
        populate_listbox(sorted_sheets, sorted_dataframes)

    # Initially populate the Listbox with the sorted sheets
    sort_ascending()  # Default to ascending order initially

    # Function to update Listbox based on search query
    def update_listbox(*args):
        search_query = entry_search.get().lower()
        filtered_pairs = [
            (sheet, df) for sheet, df in zip(extracted_sheets, processed_dataframes)
            if search_query in sheet.lower()
        ]
        listbox.delete(0, tk.END)  # Clear the listbox before updating
        if filtered_pairs:
            filtered_sheets, filtered_dataframes = zip(*filtered_pairs)
            populate_listbox(filtered_sheets, filtered_dataframes)
        else:
            listbox.insert(tk.END, "No matching sheets found.")

    # Bind KeyRelease event to Entry to trigger live searching
    entry_search.bind("<KeyRelease>", update_listbox)

    # Function to handle double-click on Listbox item
    def on_listbox_double_click(event):
        selection = listbox.curselection()
        if selection:
            selected_sheet = listbox.get(selection[0])
            if selected_sheet == "No matching sheets found.":
                return
            try:
                index = list(extracted_sheets).index(selected_sheet)

                df = processed_dataframes[index]
            except ValueError:
                return
            preview_data(df, title=f"Preview - {selected_sheet}")

    listbox.bind("<Double-Button-1>", on_listbox_double_click)

    # Buttons for sorting
    ctk.CTkButton(check_window, text="Sort Ascending", command=sort_ascending).grid(row=2, column=0, padx=10, pady=5, sticky="ew")
    ctk.CTkButton(check_window, text="Sort Descending", command=sort_descending).grid(row=2, column=1, padx=10, pady=5, sticky="ew")

    # Make the check_window modal
    check_window.transient(root)
    check_window.grab_set()
    root.wait_window(check_window)

def on_enter_key(event, current_entry, next_entry):
    next_entry.focus_set()

def on_arrow_key(event, current_entry, direction):
    if direction == 'down':
        next_entry = find_next_entry(current_entry)
        if next_entry:
            next_entry.focus_set()
    elif direction == 'up':
        previous_entry = find_previous_entry(current_entry)
        if previous_entry:
            previous_entry.focus_set()

def find_next_entry(current_entry):
    entry_fields = [
        bill_date_combobox,
        bill_number_combobox,
        second_bill_number_combobox,
        bill_amount_combobox,
        entry_file_path
    ]
    for i, entry in enumerate(entry_fields):
        if entry == current_entry and i < len(entry_fields) - 1:
            return entry_fields[i + 1]
    return None

def find_previous_entry(current_entry):
    entry_fields = [
        bill_date_combobox,
        bill_number_combobox,
        second_bill_number_combobox,
        bill_amount_combobox,
        entry_file_path
    ]
    for i, entry in enumerate(entry_fields):
        if entry == current_entry and i > 0:
            return entry_fields[i - 1]
    return None

# Function to handle window close event
def on_close():
    global is_data_extracted, is_data_saved  # Use global to check the flags

    if is_data_extracted and not is_data_saved:  # Only prompt if there's new, unsaved data
        result = messagebox.askyesnocancel(
            "Quit", "There is unsaved data. Would you like to save the session before exiting?"
        )

        if result is None:  # If the user clicks the "X" button or cancels
            return  # Do nothing and return to the application

        if result:  # If the user clicks "Yes"
            save_session_state()  # Save session state before closing

    root.destroy()  # Close the main window regardless of the choice

def on_file_or_sheet_selected(event=None):
    global selected_file_path, sheet_df

    if not selected_file_path:  # File is not selected
        selected_file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if selected_file_path:
            entry_file_path.delete(0, ctk.END)  # Update the file path entry
            entry_file_path.insert(0, selected_file_path)

            # Populate the sheet selection combobox
            try:
                xl = pd.ExcelFile(selected_file_path)
                sheets = xl.sheet_names
                sheet_selection_combobox['values'] = sheets
                sheet_selection_combobox.set(sheets[0] if sheets else "")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to load sheets: {e}")

    sheet_name = sheet_selection_combobox.get()
    if selected_file_path and sheet_name:
        load_sheet(selected_file_path, sheet_name)
        update_column_suggestions()

def setup_initial_gui():
    # Populate the sheet selection combobox if a file was already selected
    if selected_file_path:
        try:
            xl = pd.ExcelFile(selected_file_path)
            sheets = xl.sheet_names
            sheet_selection_combobox['values'] = sheets
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheets: {e}")

    # Bind the file path entry to update when Enter is pressed
    entry_file_path.bind("<Return>", lambda event: on_file_or_sheet_selected())

    # Bind the sheet selection combobox to load data when a sheet is selected
    sheet_selection_combobox.bind("<<ComboboxSelected>>", on_file_or_sheet_selected)

    # Bind the combobox selections to update column suggestions and prevent selecting the same column twice
    bill_date_combobox.bind("<<ComboboxSelected>>", lambda event: update_column_suggestions())
    bill_number_combobox.bind("<<ComboboxSelected>>", lambda event: update_column_suggestions())
    bill_amount_combobox.bind("<<ComboboxSelected>>", lambda event: update_column_suggestions())

    # Bind the file path entry to update when Enter is pressed
    entry_file_path.bind("<Return>", lambda event: on_file_or_sheet_selected())

    # Bind the sheet selection combobox to load data when a sheet is selected
    sheet_selection_combobox.bind("<<ComboboxSelected>>", on_file_or_sheet_selected)


# Creating the GUI
root = ctk.CTk()
root.title("Bill Data Extractor and Compiler")

# Configure columns and rows to expand
for i in range(7):
    root.grid_rowconfigure(i, weight=1)
for i in range(5):
    root.grid_columnconfigure(i, weight=1)

root.option_add('*Font', 'Helvetica 14')

# Light mode background color and black text color
light_gray_bg = "#d9d9d9"
text_color = "#000000"  # Black text for light mode

# Apply default style for ttk widgets like Combobox
style = ttk.Style()
style.theme_use('clam')  # Use default light mode
style.configure("TCombobox",
                fieldbackground=light_gray_bg,  # Background of the combobox field
                background=light_gray_bg,       # Dropdown background
                foreground=text_color,          # Text color
                arrowcolor=text_color)          # Arrow color

style.map('TCombobox', fieldbackground=[('readonly', light_gray_bg)],
                         selectbackground=[('readonly', light_gray_bg)],
                         selectforeground=[('readonly', text_color)])

# Menu bar setup
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# File Menu
file_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="New Session", command=start_new_session)
file_menu.add_command(label="Load Previous Session", command=load_session_state)
file_menu.add_command(label="Save Session", command=save_session_state)

# View Menu
data_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="View", menu=data_menu)
data_menu.add_command(label="Compiled Data", command=preview_compiled_data)
data_menu.add_command(label="List of Uploads", command=check_list)

# Create "Edit" menu
data_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Edit", menu=data_menu)
data_menu.add_command(label="Remove Extracted Data", command=remove_selected_data)
data_menu.add_command(label="Edit Extracted Data", command=edit_data)



# Title Label
label_title = ctk.CTkLabel(root, text="Bill Data Extraction Tool", font=("Helvetica", 16))
label_title.grid(row=0, column=0, columnspan=5, padx=10, pady=10, sticky="ew")

# Pin Window Checkbutton
pin_var = ctk.BooleanVar()
pin_button = ctk.CTkCheckBox(root, text="Pin Window", variable=pin_var, command=toggle_pin)
pin_button.grid(row=0, column=0, padx=5, pady=5, sticky="w")

# File Path Entry
label_file_path = ctk.CTkLabel(root, text="Excel File Path:")
label_file_path.grid(row=1, column=0, padx=10, pady=10, sticky="e")
entry_file_path = tk.Entry(root, bg=light_gray_bg, fg=text_color, insertbackground=text_color)
entry_file_path.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

# Browse Button
button_browse = ctk.CTkButton(root, text="Browse", corner_radius=10, command=browse_file)
button_browse.grid(row=1, column=2, padx=10, pady=10, sticky="ew")

# Sheet Selection Label and Combobox
label_sheet_selection = ctk.CTkLabel(root, text="Select Sheet:")
label_sheet_selection.grid(row=1, column=3, padx=10, pady=10, sticky="e")
sheet_selection_combobox = ttk.Combobox(root, style="TCombobox")
sheet_selection_combobox.grid(row=1, column=4, columnspan=2, padx=10, pady=10, sticky="ew")
sheet_selection_combobox.bind("<<ComboboxSelected>>", on_sheet_selected)

# Bill Date Entry
label_bill_date = ctk.CTkLabel(root, text="Bill Date Column (Letter):")
label_bill_date.grid(row=2, column=0, padx=10, pady=10, sticky="e")
bill_date_combobox = ttk.Combobox(root, style="TCombobox")
bill_date_combobox.grid(row=2, column=1, padx=10, pady=10, sticky="ew")

# Bill Number Entry
label_bill_number = ctk.CTkLabel(root, text="Bill Number Column (Letter):")
label_bill_number.grid(row=3, column=0, padx=10, pady=10, sticky="e")
bill_number_combobox = ttk.Combobox(root, style="TCombobox")
bill_number_combobox.grid(row=3, column=1, padx=10, pady=10, sticky="ew")

# Second Bill Number Checkbox and Combobox
second_bill_number_var = ctk.BooleanVar()
checkbox_second_bill_number = ctk.CTkCheckBox(root, text="Add Column", variable=second_bill_number_var, command=toggle_second_bill_number)
checkbox_second_bill_number.grid(row=3, column=3, padx=10, pady=10, sticky="e")
second_bill_number_combobox = ttk.Combobox(root, state="disabled", values=[], style="TCombobox")
second_bill_number_combobox.grid(row=3, column=2, padx=10, pady=10, sticky="ew")

# Remove non-numeric Checkbox
remove_non_numeric = ctk.BooleanVar()
check_clean_number = ctk.CTkCheckBox(root, text="Remove Letters", variable=remove_non_numeric)
check_clean_number.grid(row=3, column=4, padx=5, pady=5, sticky="w")

# Bill Amount Entry
label_bill_amount = ctk.CTkLabel(root, text="Bill Amount Column (Letter):")
label_bill_amount.grid(row=4, column=0, padx=10, pady=10, sticky="e")
bill_amount_combobox = ttk.Combobox(root, style="TCombobox")
bill_amount_combobox.grid(row=4, column=1, padx=10, pady=10, sticky="ew")

# Extract Button
button_extract = ctk.CTkButton(root, corner_radius=10, text="Extract Data", command=start_extraction)
button_extract.grid(row=6, column=1, padx=10, pady=10, columnspan=2, sticky="ew")

# Send to database Button
button_compile = ctk.CTkButton(root, corner_radius=10, text="Add to Database", command=insert_data_to_db)
button_compile.grid(row=7, column=3, padx=10, pady=10, columnspan=2, sticky="ew")

# Compile Button
button_compile = ctk.CTkButton(root, corner_radius=10, text="Compile Files", command=compile_files)
button_compile.grid(row=6, column=3, padx=10, pady=10, columnspan=2, sticky="ew")

# Function to navigate between widgets using Enter key
def bind_nav(widget_from, widget_to):
    widget_from.bind("<Return>", lambda event: widget_to.focus_set() if widget_to else None)
    widget_from.bind("<Down>", lambda event: widget_to.focus_set() if widget_to else None)
    widget_from.bind("<Up>", lambda event: widget_to.focus_set() if widget_to else None)

# Example where the function is applied to bind widgets in sequence:
bind_nav(entry_file_path, sheet_selection_combobox)
bind_nav(sheet_selection_combobox, bill_date_combobox)
bind_nav(bill_date_combobox, bill_number_combobox)
bind_nav(bill_number_combobox, second_bill_number_combobox)
bind_nav(second_bill_number_combobox, bill_amount_combobox)
bind_nav(bill_amount_combobox, button_extract)


# Set the close protocol
root.protocol("WM_DELETE_WINDOW", on_close)

# Start the Tkinter event loop
root.mainloop()


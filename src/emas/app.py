"""Application controller and desktop window construction."""
import sys
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from tkcalendar import DateEntry
from .sessions import SessionMixin
from .extraction import ExtractionMixin
from .views import ViewMixin
from .navigation import NavigationMixin
from .database_actions import DatabaseActionsMixin

class Application(SessionMixin, ExtractionMixin, ViewMixin, NavigationMixin, DatabaseActionsMixin):

    def __init__(self):
        self.processed_dataframes = []
        self.extracted_hospitals = []
        self.recon_data = {}
        self.extracted_sheets = []
        self.is_data_extracted = False
        self.is_data_saved = True
        self.selected_file_path = None
        self.sheet_df = None

    def sign_in_window(self):
        ctk.set_default_color_theme('blue')
        sign_in = ctk.CTk()
        sign_in.title('Sign-In Window')
        sign_in.geometry('400x400')
        sign_in.resizable(False, False)

        def on_close():
            messagebox.showwarning('Exiting', 'You closed the window without submitting. Exiting the program.')
            sign_in.destroy()
            sys.exit()
        sign_in.protocol('WM_DELETE_WINDOW', on_close)
        ctk.CTkLabel(sign_in, text='Welcome to the Reconciliation Tool', font=('Helvetica', 18)).pack(pady=10)
        ctk.CTkLabel(sign_in, text='Recon By:', font=('Helvetica', 14)).pack(pady=5)
        recon_by_entry = ctk.CTkEntry(sign_in, placeholder_text='Enter your name', font=('Helvetica', 14))
        recon_by_entry.pack(pady=5)
        ctk.CTkLabel(sign_in, text='Recon Date:', font=('Helvetica', 14)).pack(pady=5)
        recon_date_entry = DateEntry(sign_in, date_pattern='y/mm/dd', font=('Helvetica', 16), width=15)
        recon_date_entry.pack(pady=5)
        recon_date_entry.set_date(datetime.today())
        ctk.CTkLabel(sign_in, text='SOA Date:', font=('Helvetica', 14)).pack(pady=5)
        soa_date_entry = DateEntry(sign_in, date_pattern='y/mm/dd', font=('Helvetica', 16), width=15)
        soa_date_entry.pack(pady=5)
        submit_button = ctk.CTkButton(sign_in, text='Submit', font=('Helvetica', 16), command=lambda: submit_action(sign_in, recon_by_entry, recon_date_entry, soa_date_entry))
        submit_button.pack(pady=20)

        def submit_action(window, recon_by, recon_date, soa_date):
            recon_by_value = recon_by.get().strip()
            recon_date_value = recon_date.get()
            soa_date_value = soa_date.get()
            if not recon_by_value:
                messagebox.showwarning('Missing Name', "You must enter a name for 'Recon By'. Exiting...")
                window.destroy()
                sys.exit()
            if recon_by_value and recon_date_value and soa_date_value:
                self.recon_data['Recon By'] = recon_by_value
                self.recon_data['Recon Date'] = recon_date_value
                self.recon_data['SOA Date'] = soa_date_value
                messagebox.showinfo('Success', 'Data saved successfully.')
                window.destroy()
            else:
                messagebox.showwarning('Missing Information', 'All fields are required.')

        def on_enter_key(event, current_entry, next_entry):
            if next_entry:
                next_entry.focus_set()
        recon_by_entry.bind('<Return>', lambda event: on_enter_key(event, recon_by_entry, recon_date_entry))
        recon_date_entry.bind('<Return>', lambda event: on_enter_key(event, recon_date_entry, soa_date_entry))
        soa_date_entry.bind('<Return>', lambda event: on_enter_key(event, soa_date_entry, submit_button))
        recon_by_entry.bind('<Down>', lambda event: recon_date_entry.focus_set())
        recon_by_entry.bind('<Up>', lambda event: recon_by_entry.focus_set())
        recon_date_entry.bind('<Down>', lambda event: soa_date_entry.focus_set())
        recon_date_entry.bind('<Up>', lambda event: recon_by_entry.focus_set())
        soa_date_entry.bind('<Down>', lambda event: submit_button.focus_set())
        soa_date_entry.bind('<Up>', lambda event: recon_date_entry.focus_set())
        sign_in.mainloop()
        if self.recon_data:
            self.main_interface()

    def main_interface(self):
        self.root = ctk.CTk()
        self.root.title('Bill Data Extractor and Compiler')
        for i in range(7):
            self.root.grid_rowconfigure(i, weight=1)
        for i in range(5):
            self.root.grid_columnconfigure(i, weight=1)
        self.root.option_add('*Font', 'Helvetica 14')
        self.light_gray_bg = '#d9d9d9'
        self.text_color = '#000000'
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TCombobox', fieldbackground=self.light_gray_bg, background=self.light_gray_bg, foreground=self.text_color, arrowcolor=self.text_color)
        self.style.map('TCombobox', fieldbackground=[('readonly', self.light_gray_bg)], selectbackground=[('readonly', self.light_gray_bg)], selectforeground=[('readonly', self.text_color)])
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label='File', menu=self.file_menu)
        self.file_menu.add_command(label='New Session', command=self.start_new_session)
        self.file_menu.add_command(label='Load Previous Session', command=self.load_session_state)
        self.file_menu.add_command(label='Save Session', command=self.save_session_state)
        self.data_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label='View', menu=self.data_menu)
        self.data_menu.add_command(label='Compiled Data', command=self.preview_compiled_data)
        self.data_menu.add_command(label='List of Uploads', command=self.check_list)
        self.data_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label='Edit', menu=self.data_menu)
        self.data_menu.add_command(label='Remove Extracted Data', command=self.remove_selected_data)
        self.data_menu.add_command(label='Edit Extracted Data', command=self.edit_data)
        self.label_title = ctk.CTkLabel(self.root, text='Bill Data Extraction Tool', font=('Helvetica', 16))
        self.label_title.grid(row=0, column=0, columnspan=5, padx=10, pady=10, sticky='ew')
        self.pin_var = ctk.BooleanVar()
        self.pin_button = ctk.CTkCheckBox(self.root, text='Pin Window', variable=self.pin_var, command=self.toggle_pin)
        self.pin_button.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.label_file_path = ctk.CTkLabel(self.root, text='Excel File Path:')
        self.label_file_path.grid(row=1, column=0, padx=10, pady=10, sticky='e')
        self.entry_file_path = tk.Entry(self.root, bg=self.light_gray_bg, fg=self.text_color, insertbackground=self.text_color)
        self.entry_file_path.grid(row=1, column=1, padx=10, pady=10, sticky='ew')
        self.button_browse = ctk.CTkButton(self.root, text='Browse', corner_radius=10, command=self.browse_file)
        self.button_browse.grid(row=1, column=2, padx=10, pady=10, sticky='ew')
        self.label_sheet_selection = ctk.CTkLabel(self.root, text='Select Sheet:')
        self.label_sheet_selection.grid(row=1, column=3, padx=10, pady=10, sticky='e')
        self.sheet_selection_combobox = ttk.Combobox(self.root, style='TCombobox')
        self.sheet_selection_combobox.grid(row=1, column=4, columnspan=2, padx=10, pady=10, sticky='ew')
        self.sheet_selection_combobox.bind('<<ComboboxSelected>>', self.on_sheet_selected)
        self.label_bill_date = ctk.CTkLabel(self.root, text='Bill Date Column (Letter):')
        self.label_bill_date.grid(row=2, column=0, padx=10, pady=10, sticky='e')
        self.bill_date_combobox = ttk.Combobox(self.root, style='TCombobox')
        self.bill_date_combobox.grid(row=2, column=1, padx=10, pady=10, sticky='ew')
        self.label_bill_number = ctk.CTkLabel(self.root, text='Bill Number Column (Letter):')
        self.label_bill_number.grid(row=3, column=0, padx=10, pady=10, sticky='e')
        self.bill_number_combobox = ttk.Combobox(self.root, style='TCombobox')
        self.bill_number_combobox.grid(row=3, column=1, padx=10, pady=10, sticky='ew')
        self.second_bill_number_var = ctk.BooleanVar()
        self.checkbox_second_bill_number = ctk.CTkCheckBox(self.root, text='Add Column', variable=self.second_bill_number_var, command=self.toggle_second_bill_number)
        self.checkbox_second_bill_number.grid(row=3, column=3, padx=10, pady=10, sticky='e')
        self.second_bill_number_combobox = ttk.Combobox(self.root, state='disabled', values=[], style='TCombobox')
        self.second_bill_number_combobox.grid(row=3, column=2, padx=10, pady=10, sticky='ew')
        self.remove_non_numeric = ctk.BooleanVar()
        self.check_clean_number = ctk.CTkCheckBox(self.root, text='Remove Letters', variable=self.remove_non_numeric)
        self.check_clean_number.grid(row=3, column=4, padx=5, pady=5, sticky='w')
        self.label_bill_amount = ctk.CTkLabel(self.root, text='Bill Amount Column (Letter):')
        self.label_bill_amount.grid(row=4, column=0, padx=10, pady=10, sticky='e')
        self.bill_amount_combobox = ttk.Combobox(self.root, style='TCombobox')
        self.bill_amount_combobox.grid(row=4, column=1, padx=10, pady=10, sticky='ew')
        self.button_extract = ctk.CTkButton(self.root, corner_radius=10, text='Extract Data', command=self.start_extraction)
        self.button_extract.grid(row=6, column=1, padx=10, pady=10, columnspan=2, sticky='ew')
        self.button_compile = ctk.CTkButton(self.root, corner_radius=10, text='Add to Database', command=self.insert_data_to_db)
        self.button_compile.grid(row=7, column=3, padx=10, pady=10, columnspan=2, sticky='ew')
        self.button_compile = ctk.CTkButton(self.root, corner_radius=10, text='Compile Files', command=self.compile_files)
        self.button_compile.grid(row=6, column=3, padx=10, pady=10, columnspan=2, sticky='ew')
        self.bind_nav(self.entry_file_path, self.sheet_selection_combobox)
        self.bind_nav(self.sheet_selection_combobox, self.bill_date_combobox)
        self.bind_nav(self.bill_date_combobox, self.bill_number_combobox)
        self.bind_nav(self.bill_number_combobox, self.second_bill_number_combobox)
        self.bind_nav(self.second_bill_number_combobox, self.bill_amount_combobox)
        self.bind_nav(self.bill_amount_combobox, self.button_extract)
        self.root.protocol('WM_DELETE_WINDOW', self.on_close)
        self.setup_initial_gui()
        self.root.mainloop()

def main():
    Application().sign_in_window()

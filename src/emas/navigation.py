"""Navigation callbacks for the application controller."""
from tkinter import messagebox
import customtkinter as ctk

class NavigationMixin:

    def change_appearance(self, mode):
        ctk.set_appearance_mode(mode)

    def toggle_pin(self):
        if self.pin_var.get():
            self.root.attributes('-topmost', True)
        else:
            self.root.attributes('-topmost', False)

    def prompt_session_choice(self):
        choice = messagebox.askquestion('Session Choice', "Do you want to load a previous session? (Click 'No' to start a new session)")
        if choice == 'yes':
            self.load_session_state()
        else:
            self.start_new_session()

    def toggle_second_bill_number(self):
        if self.second_bill_number_var.get():
            self.second_bill_number_combobox.configure(state='normal')
        else:
            self.second_bill_number_combobox.configure(state='disabled')
            self.checkbox_second_bill_number.configure(command=self.toggle_second_bill_number)

    def on_enter_key(self, event, current_entry, next_entry):
        next_entry.focus_set()

    def on_arrow_key(self, event, current_entry, direction):
        if direction == 'down':
            next_entry = self.find_next_entry(current_entry)
            if next_entry:
                next_entry.focus_set()
        elif direction == 'up':
            previous_entry = self.find_previous_entry(current_entry)
            if previous_entry:
                previous_entry.focus_set()

    def find_next_entry(self, current_entry):
        entry_fields = [self.bill_date_combobox, self.bill_number_combobox, self.second_bill_number_combobox, self.bill_amount_combobox, self.entry_file_path]
        for i, entry in enumerate(entry_fields):
            if entry == current_entry and i < len(entry_fields) - 1:
                return entry_fields[i + 1]
        return None

    def find_previous_entry(self, current_entry):
        entry_fields = [self.bill_date_combobox, self.bill_number_combobox, self.second_bill_number_combobox, self.bill_amount_combobox, self.entry_file_path]
        for i, entry in enumerate(entry_fields):
            if entry == current_entry and i > 0:
                return entry_fields[i - 1]
        return None

    def bind_nav(self, widget_from, widget_to):
        widget_from.bind('<Return>', lambda event: widget_to.focus_set() if widget_to else None)
        widget_from.bind('<Down>', lambda event: widget_to.focus_set() if widget_to else None)
        widget_from.bind('<Up>', lambda event: widget_to.focus_set() if widget_to else None)

"""Extraction callbacks for the application controller."""
import os
import re
import logging
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from .helpers import index_to_letter, create_reconref, letter_to_index, combine_bill_numbers
from .database import connection
from .claims import fetch_claim_data

class ExtractionMixin:

    def load_sheet(self, file_path, sheet_name):
        try:
            self.sheet_df = pd.read_excel(file_path, sheet_name=sheet_name)
        except Exception as e:
            messagebox.showerror('Error', f'Failed to load sheet: {e}')
            return None
        return self.sheet_df

    def update_column_suggestions(self):
        if self.sheet_df is not None:
            columns = self.sheet_df.columns
            column_letters = [index_to_letter(i) for i in range(len(columns))]
            selected_columns = [self.bill_date_combobox.get(), self.bill_number_combobox.get(), self.bill_amount_combobox.get(), self.second_bill_number_combobox.get()]
            available_columns = [letter for letter in column_letters if letter not in selected_columns]
            if not self.bill_date_combobox.get():
                self.bill_date_combobox['values'] = available_columns
            if not self.bill_number_combobox.get():
                self.bill_number_combobox['values'] = available_columns
            if not self.bill_amount_combobox.get():
                self.bill_amount_combobox['values'] = available_columns
            if not self.second_bill_number_combobox.get():
                self.second_bill_number_combobox['values'] = available_columns

    def add_extracted_data(self, sheet_name, extracted_data):
        if sheet_name in self.extracted_sheets:
            self.remove_extracted_data(sheet_name)
        self.extracted_sheets.append(sheet_name)
        self.processed_dataframes.append(extracted_data)
        providers = extracted_data.get('ProviderName', pd.Series(dtype=str))
        self.extracted_hospitals.append(', '.join(providers.dropna().astype(str).unique()))
        self.is_data_extracted = True
        self.is_data_saved = False

    def remove_extracted_data(self, sheet_name):
        if sheet_name not in self.extracted_sheets:
            return
        index = self.extracted_sheets.index(sheet_name)
        self.extracted_sheets.pop(index)
        self.processed_dataframes.pop(index)
        if index < len(self.extracted_hospitals):
            self.extracted_hospitals.pop(index)
        self.is_data_extracted = bool(self.processed_dataframes)
        self.is_data_saved = False

    def extract_data(self, bill_date_letter, bill_number_letter, bill_amount_letter, file_path, sheet_name):
        selected_columns = [bill_date_letter.upper(), bill_number_letter.upper(), bill_amount_letter.upper()]
        if self.second_bill_number_var.get():
            selected_columns.append(self.second_bill_number_combobox.get().upper())
        if len(selected_columns) != len(set(selected_columns)):
            messagebox.showerror('Selection Error', 'You have selected the same column for more than one field. Please choose different columns for each field.')
            return
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            bill_date_idx = letter_to_index(bill_date_letter)
            bill_number_idx = letter_to_index(bill_number_letter)
            bill_amount_idx = letter_to_index(bill_amount_letter)
            data_without_header = df.iloc[1:]
            extracted_data = data_without_header.iloc[:, [bill_date_idx, bill_number_idx, bill_amount_idx]].copy()
            extracted_data.columns = ['Bill Date', 'Bill Number', 'Bill Amount']
            try:
                extracted_data['Bill Date'] = pd.to_datetime(extracted_data['Bill Date'], dayfirst=True).dt.date
            except Exception as date_error:
                raise ValueError(f'Date conversion failed: {date_error}')
            extracted_data['Bill Number'] = extracted_data['Bill Number'].apply(str)
            if self.second_bill_number_var.get():
                second_bill_number_idx = letter_to_index(self.second_bill_number_combobox.get())
                extracted_data['Second Bill Number'] = data_without_header.iloc[:, second_bill_number_idx]
                extracted_data['Bill Number'] = extracted_data.apply(lambda row: combine_bill_numbers(str(row['Bill Number']), str(row['Second Bill Number'])), axis=1)
                extracted_data.drop(columns=['Second Bill Number'], inplace=True)
            if self.remove_non_numeric.get():
                extracted_data['Bill Number'] = extracted_data['Bill Number'].apply(lambda x: re.sub('[a-zA-Z]', '', str(x)))
            if sheet_name in self.extracted_sheets:
                user_choice = messagebox.askyesno('Duplicate Sheet Found', f"The sheet '{sheet_name}' has already been extracted. Do you want to replace the existing data?")
                if not user_choice:
                    messagebox.showinfo('Data Kept', f"The existing data for sheet '{sheet_name}' has been kept.")
                    return
            bill_numbers = extracted_data['Bill Number'].unique().tolist()
            sql_data = connection.pandas_dataframe(bill_numbers)
            combined_data = pd.merge(extracted_data, sql_data, how='left', left_on='Bill Number', right_on='BillNo')
            if 'ClaimID' in combined_data.columns:
                combined_data['ClaimID'] = combined_data['ClaimID'].astype(pd.Int64Dtype())
                DMS_ID = combined_data['ClaimID'].apply(fetch_claim_data)
                combined_data['DMSID'] = DMS_ID
            else:
                raise ValueError("SQL data does not contain 'ClaimID' column.")
            combined_data['Recon By'] = self.recon_data.get('Recon By', '')
            combined_data['Recon Date'] = self.recon_data.get('Recon Date', '')
            combined_data['SOA Date'] = self.recon_data.get('SOA Date', '')
            combined_data['Bill Number'] = combined_data['Bill Number'].astype(str)
            combined_data['BillNo'] = combined_data['BillNo'].astype(str)
            combined_data['BillNo'] = combined_data['BillNo'].replace('NaN', '')
            combined_data['Bill Amount'] = pd.to_numeric(combined_data['Bill Amount'], errors='coerce')
            combined_data['Bill Amount'] = combined_data['Bill Amount'].apply(lambda x: f'{x:.2f}' if pd.notna(x) else '')
            combined_data['ClaimID'] = combined_data['ClaimID'].astype(pd.Int64Dtype())
            combined_data['ProviderID'] = combined_data['ProviderID'].astype(pd.Int64Dtype())
            combined_data['Reconrefno'] = combined_data.apply(lambda row: create_reconref(row['SOA Date'], row['Recon Date']), axis=1)
            combined_data['ModifyDate'] = ''
            combined_data['ModifyBy'] = ''
            combined_data['ReconRemarks'] = ''
            combined_data['CreateDate'] = datetime.now().date()
            self.add_extracted_data(sheet_name, combined_data)
            self.is_data_extracted = True
            self.is_data_saved = False
            self.preview_data(combined_data, title='Preview')
            messagebox.showinfo('Success', f'Data extracted from {os.path.basename(file_path)} successfully!')
            return True
        except ValueError as ve:
            logging.debug('Extraction failed', exc_info=True)
            messagebox.showerror('Input Error', f'Input Error: {ve}')
        except Exception as e:
            messagebox.showerror('Error', f'An error occurred during extraction: {e}')

    def start_extraction(self):
        bill_date_letter = self.bill_date_combobox.get()
        bill_number_letter = self.bill_number_combobox.get()
        bill_amount_letter = self.bill_amount_combobox.get()
        file_path = self.selected_file_path
        sheet_name = self.sheet_selection_combobox.get()
        if not file_path:
            messagebox.showwarning('No File', 'Please select a file to extract data from.')
            return
        if not all([bill_date_letter, bill_number_letter, bill_amount_letter]):
            messagebox.showwarning('Missing Information', 'Please fill in all required fields.')
            return
        if not sheet_name:
            messagebox.showwarning('No Sheet Selected', 'Please select a sheet to extract data from.')
            return
        try:
            success = self.extract_data(bill_date_letter, bill_number_letter, bill_amount_letter, file_path, sheet_name)
            if not success:
                return
            self.bill_date_combobox.set('')
            self.bill_number_combobox.set('')
            self.bill_amount_combobox.set('')
            self.second_bill_number_combobox.set('')
            self.second_bill_number_var.set(False)
            self.remove_non_numeric.set(False)
            self.second_bill_number_combobox.config(state='disabled')
        except Exception as e:
            messagebox.showerror('Error', f'An error occurred: {e}')

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])
        if file_path:
            self.selected_file_path = file_path
            self.entry_file_path.delete(0, ctk.END)
            self.entry_file_path.insert(0, file_path)
            try:
                xl = pd.ExcelFile(file_path)
                sheets = xl.sheet_names
                self.sheet_selection_combobox['values'] = sheets
                self.sheet_selection_combobox.set(sheets[0] if sheets else '')
                if sheets:
                    self.load_sheet(file_path, sheets[0])
                    self.update_column_suggestions()
            except Exception as e:
                messagebox.showerror('Error', f'Failed to load sheets: {e}')

    def on_sheet_selected(self, event):
        sheet_name = self.sheet_selection_combobox.get()
        if sheet_name and self.selected_file_path:
            self.load_sheet(self.selected_file_path, sheet_name)
            self.update_column_suggestions()

    def on_file_or_sheet_selected(self, event=None):
        entered = self.entry_file_path.get().strip()
        if entered:
            self.selected_file_path = entered
        if not self.selected_file_path:
            self.selected_file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx *.xls')])
            if self.selected_file_path:
                self.entry_file_path.delete(0, ctk.END)
                self.entry_file_path.insert(0, self.selected_file_path)
                try:
                    xl = pd.ExcelFile(self.selected_file_path)
                    sheets = xl.sheet_names
                    self.sheet_selection_combobox['values'] = sheets
                    self.sheet_selection_combobox.set(sheets[0] if sheets else '')
                except Exception as e:
                    messagebox.showerror('Error', f'Failed to load sheets: {e}')
        sheet_name = self.sheet_selection_combobox.get()
        if self.selected_file_path and sheet_name:
            self.load_sheet(self.selected_file_path, sheet_name)
            self.update_column_suggestions()

    def setup_initial_gui(self):
        self.entry_file_path.delete(0, tk.END)
        if self.selected_file_path:
            self.entry_file_path.insert(0, self.selected_file_path)
            try:
                with pd.ExcelFile(self.selected_file_path) as workbook:
                    self.sheet_selection_combobox.configure(values=workbook.sheet_names)
            except Exception as exc:
                messagebox.showerror('Error', f'Failed to load sheets: {exc}')
        self.entry_file_path.bind('<Return>', self.on_file_or_sheet_selected)
        self.sheet_selection_combobox.bind('<<ComboboxSelected>>', self.on_file_or_sheet_selected)
        for widget in (self.bill_date_combobox, self.bill_number_combobox, self.bill_amount_combobox, self.second_bill_number_combobox):
            widget.bind('<<ComboboxSelected>>', lambda event: self.update_column_suggestions())

"""Views callbacks for the application controller."""
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, ttk
import customtkinter as ctk
from .database import connection
from .claims import fetch_claim_data

class ViewMixin:

    def preview_data(self, df, title='Data Preview'):
        preview_window = ctk.CTkToplevel(self.root)
        preview_window.title(title)
        preview_window.transient(self.root)
        preview_window.grab_set()
        preview_window.geometry('800x400')
        preview_window.resizable(True, True)
        style = ttk.Style(preview_window)
        style.theme_use('clam')
        style.configure('Treeview', background='#f0f0f0', foreground='black', rowheight=25, fieldbackground='#f0f0f0')
        style.configure('Treeview.Heading', background='#003366', foreground='white', font=('Arial', 10, 'bold'))
        style.map('Treeview', background=[('selected', '#003366')], foreground=[('selected', 'white')])
        tree = ttk.Treeview(preview_window, selectmode='browse', style='Treeview')
        tree['column'] = list(df.columns)
        tree['show'] = 'headings'
        for col in df.columns:
            tree.heading(col, text=col, anchor=ctk.CENTER)
            tree.column(col, anchor=ctk.CENTER, width=120)
        for index, row in df.iterrows():
            tree.insert('', 'end', values=list(row), tags=('oddrow' if index % 2 == 0 else 'evenrow',))
        tree.tag_configure('oddrow', background='#e6f2ff')
        tree.tag_configure('evenrow', background='#ffffff')
        scrollbar = ctk.CTkScrollbar(preview_window, orientation='vertical', command=tree.yview)
        horizontal_scrollbar = ctk.CTkScrollbar(preview_window, orientation='horizontal', command=tree.xview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.configure(xscrollcommand=horizontal_scrollbar.set)
        tree.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        horizontal_scrollbar.grid(row=1, column=0, sticky='ew')
        preview_window.grid_rowconfigure(0, weight=1)
        preview_window.grid_columnconfigure(0, weight=1)
        preview_window.protocol('WM_DELETE_WINDOW', preview_window.destroy)

    def edit_data(self):
        if not self.processed_dataframes:
            messagebox.showwarning('No Data', 'Extract data before editing.')
            return
        if isinstance(self.processed_dataframes, list) and all((isinstance(df, pd.DataFrame) for df in self.processed_dataframes)):
            df = pd.concat(self.processed_dataframes, ignore_index=True).copy()
        else:
            return
        original_bill_number = None

        def on_double_click(event, tree):
            try:
                item = tree.selection()[0]
                column = tree.identify_column(event.x)
                col = int(column.split('#')[-1]) - 1
                if col == 1:
                    start_edit(tree, item, col)
            except IndexError:
                pass

        def start_edit(tree, item, col):
            nonlocal original_bill_number
            value = tree.item(item, 'values')[col]
            original_bill_number = tree.item(item, 'values')[1]
            edit_window = ctk.CTkToplevel(tree)
            edit_window.title('Edit Bill Number')
            edit_window.geometry('200x100')
            edit_entry = ctk.CTkEntry(edit_window, width=150)
            edit_entry.insert(0, value)
            edit_entry.grid(row=0, column=0, padx=10, pady=10)
            edit_entry.focus()
            save_button = ctk.CTkButton(edit_window, text='Save', command=lambda: save_edit(tree, item, col, edit_entry.get(), edit_window))
            save_button.grid(row=1, column=0, padx=5, sticky=tk.W)
            cancel_button = ctk.CTkButton(edit_window, text='Cancel', command=edit_window.destroy)
            cancel_button.grid(row=1, column=1, padx=5, sticky=tk.E)

        def save_edit(tree, item, col, new_value, edit_window):
            values = list(tree.item(item, 'values'))
            old_billnumber = values[1]
            new_list = values
            new_list[col] = new_value
            sql_data = connection.pandas_dataframe([new_value])
            for original_df in self.processed_dataframes:
                if not (original_df['Bill Number'] == old_billnumber).any():
                    continue
                if len(sql_data) == 1:
                    new_list[3] = sql_data['ClaimID'].astype(pd.Int64Dtype()).tolist()[0]
                    new_list[4] = sql_data['BillNo'].astype(str).tolist()[0]
                    new_list[5] = sql_data['ProviderID'].astype(pd.Int64Dtype()).tolist()[0]
                    new_list[6] = sql_data['ProviderName'].tolist()[0]
                    new_list[7] = sql_data['InvoiceNo'].astype(str).tolist()[0]
                    new_list[8] = sql_data['ClaimAmount'].apply(pd.to_numeric).tolist()[0]
                    new_list[9] = sql_data['CompanyName'].tolist()[0]
                    new_list[10] = sql_data['ClientPaidDate'].apply(pd.to_datetime).tolist()[0]
                    new_list[11] = sql_data['PayOrderNo'].tolist()[0]
                    new_list[12] = sql_data['Status'].tolist()[0]
                    original_df.loc[original_df['Bill Number'] == old_billnumber, 'Bill Number'] = new_value
                    original_df.loc[original_df['Bill Number'] == new_value, 'ClaimID'] = sql_data['ClaimID'].astype(pd.Int64Dtype()).tolist()[0]
                    original_df.loc[original_df['Bill Number'] == new_value, 'BillNo'] = sql_data['BillNo'].astype(str).tolist()[0]
                    original_df.loc[original_df['Bill Number'] == new_value, 'ProviderID'] = sql_data['ProviderID'].astype(pd.Int64Dtype()).tolist()[0]
                    original_df.loc[original_df['Bill Number'] == new_value, 'ProviderName'] = sql_data['ProviderName'].tolist()[0]
                    original_df.loc[original_df['Bill Number'] == new_value, 'InvoiceNo'] = sql_data['InvoiceNo'].astype(str).tolist()[0]
                    original_df.loc[original_df['Bill Number'] == new_value, 'ClaimAmount'] = sql_data['ClaimAmount'].apply(pd.to_numeric).tolist()[0]
                    original_df.loc[original_df['Bill Number'] == new_value, 'CompanyName'] = sql_data['CompanyName'].tolist()[0]
                    original_df.loc[original_df['Bill Number'] == new_value, 'ClientPaidDate'] = sql_data['ClientPaidDate'].apply(pd.to_datetime).tolist()[0]
                    original_df.loc[original_df['Bill Number'] == new_value, 'PayOrderNo'] = sql_data['PayOrderNo'].tolist()[0]
                    original_df.loc[original_df['Bill Number'] == new_value, 'Status'] = sql_data['Status'].tolist()[0]
                    DMS_ID = original_df.loc[original_df['Bill Number'] == new_value, 'ClaimID'].apply(fetch_claim_data)
                    original_df.loc[original_df['Bill Number'] == new_value, 'DMSID'] = DMS_ID
                    new_list[13] = DMS_ID.iloc[0] if len(DMS_ID) else None
                elif len(sql_data) == 0:
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
                    original_df.loc[original_df['Bill Number'] == new_value, 'ProviderID'] = None
                    original_df.loc[original_df['Bill Number'] == new_value, 'ProviderName'] = None
                    original_df.loc[original_df['Bill Number'] == new_value, 'InvoiceNo'] = None
                    original_df.loc[original_df['Bill Number'] == new_value, 'ClaimAmount'] = None
                    original_df.loc[original_df['Bill Number'] == new_value, 'CompanyName'] = None
                    original_df.loc[original_df['Bill Number'] == new_value, 'ClientPaidDate'] = None
                    original_df.loc[original_df['Bill Number'] == new_value, 'PayOrderNo'] = None
                    original_df.loc[original_df['Bill Number'] == new_value, 'Status'] = None
                    original_df.loc[original_df['Bill Number'] == new_value, 'DMSID'] = None
                    new_list[13] = None
            self.is_data_saved = False
            tree.item(item, values=new_list)
            edit_window.destroy()

        def add_remarks(tree, remarks_entry, selected_item):
            if selected_item:
                for original_df in self.processed_dataframes:
                    remarks = remarks_entry.get()
                    bill_number = tree.item(selected_item, 'values')[1]
                    if not (original_df['Bill Number'] == bill_number).any():
                        continue
                    original_df.loc[original_df['Bill Number'] == bill_number, 'ReconRemarks'] = remarks
                    self.is_data_saved = False
                    updated_row = original_df[original_df['Bill Number'] == bill_number].iloc[0]
                    tree.item(selected_item, values=list(updated_row))
                remarks_entry.delete(0, tk.END)

        def on_row_select(event, tree, remarks_entry):
            if not tree.selection():
                return
            selected_item = tree.selection()[0]
            remarks = tree.item(selected_item, 'values')[list(df.columns).index('ReconRemarks')]
            remarks_entry.delete(0, tk.END)
            remarks_entry.insert(0, remarks)

        def save_changes_to_original():
            self.is_data_saved = False
            messagebox.showinfo('Edits Applied', 'Edits are applied to the extracted data. Save the session or compile a report to keep them.')

        def create_editable_treeview():
            window = ctk.CTkToplevel(self.root)
            window.title('Editable Treeview Example')
            frame = ctk.CTkFrame(window)
            frame.pack(fill=tk.BOTH, expand=True)
            columns = list(df.columns)
            tree = ttk.Treeview(frame, columns=columns, show='headings')
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            horizontal_scrollbar = ttk.Scrollbar(frame, orient='horizontal', command=tree.xview)
            horizontal_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            tree.configure(xscrollcommand=horizontal_scrollbar.set)
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=100)
            for index, row in df.iterrows():
                values = tuple(row.fillna(''))
                tree.insert('', tk.END, iid=index, values=values)
            remarks_entry = ctk.CTkEntry(window, width=300)
            remarks_entry.pack(padx=20, pady=10)
            remarks_button = ctk.CTkButton(window, text='Update Remarks', command=lambda: add_remarks(tree, remarks_entry, tree.selection()[0] if tree.selection() else None))
            remarks_button.pack(padx=20, pady=10)
            save_button = ctk.CTkButton(window, text='Save Changes', command=save_changes_to_original)
            save_button.pack(padx=20, pady=10)
            window.grid_rowconfigure(0, weight=1)
            window.grid_columnconfigure(0, weight=1)
            tree.bind('<Double-1>', lambda event: on_double_click(event, tree))
            tree.bind('<<TreeviewSelect>>', lambda event: on_row_select(event, tree, remarks_entry))
        create_editable_treeview()

    def remove_selected_data(self):
        if not self.extracted_sheets:
            messagebox.showwarning('No Data', 'No sheets available to show.')
            return
        checklist_window = ctk.CTkToplevel(self.root)
        checklist_window.title('Sheet Checklist')

        def update_listbox(search_term=''):
            listbox.delete(0, ctk.END)
            for sheet in self.extracted_sheets:
                if search_term.lower() in sheet.lower():
                    listbox.insert(ctk.END, sheet)

        def on_search(event):
            search_term = search_entry.get()
            update_listbox(search_term)
        search_frame = ctk.CTkFrame(checklist_window)
        search_frame.pack(padx=10, pady=10)
        search_label = ctk.CTkLabel(search_frame, text='Search:')
        search_label.pack(side=ctk.LEFT)
        search_entry = ctk.CTkEntry(search_frame)
        search_entry.pack(side=ctk.LEFT, fill=ctk.X, expand=True)
        search_entry.bind('<KeyRelease>', on_search)
        listbox = Listbox(checklist_window, selectmode='multiple', height=10, width=50)
        listbox.pack(padx=10, pady=10)
        update_listbox()
        ctk.CTkButton(checklist_window, text='Remove Selected', command=lambda: self.remove_selected(listbox, checklist_window)).pack(pady=5)
        ctk.CTkButton(checklist_window, text='Close', command=checklist_window.destroy).pack(pady=10)

    def update_listbox(self, listbox, search_term=''):
        listbox.delete(0, ctk.END)
        for sheet in self.extracted_sheets:
            if search_term.lower() in sheet.lower():
                listbox.insert(ctk.END, sheet)

    def remove_selected(self, listbox, checklist_window):
        selected = [listbox.get(i) for i in listbox.curselection()]
        if not selected:
            messagebox.showwarning('No Selection', 'Please select sheets to remove.')
            return
        for sheet in selected:
            self.remove_extracted_data(sheet)
        checklist_window.destroy()

    def compile_files(self):
        if not self.processed_dataframes:
            messagebox.showwarning('No Data', 'No files processed yet. Please extract data first.')
            return
        for df in self.processed_dataframes:
            for key, value in self.recon_data.items():
                df[key] = value
        combined_data = pd.concat(self.processed_dataframes, ignore_index=True)
        output_file = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')], title='Save Compiled File As')
        if output_file:
            try:
                combined_data.to_excel(output_file, index=False)
                hospital_log_file = os.path.splitext(output_file)[0] + '_Hospitals.txt'
                with open(hospital_log_file, 'w') as file:
                    file.write('Extracted Hospitals:\n')
                    file.write('\n'.join(self.extracted_hospitals))
                messagebox.showinfo('Success', f'All files compiled and saved to {output_file} successfully!\nHospital names saved to {hospital_log_file}')
            except Exception as e:
                messagebox.showerror('Error', f'An error occurred: {e}')

    def preview_compiled_data(self):
        if not self.processed_dataframes:
            messagebox.showwarning('No Data', 'No data available to preview. Please extract data first.')
            return
        combined_data = pd.concat(self.processed_dataframes, ignore_index=True)
        self.preview_data(combined_data, title='Compiled Data Preview')

    def check_list(self):
        if not self.extracted_sheets:
            messagebox.showinfo('No Data', 'No sheets have been extracted yet.')
            return
        check_window = ctk.CTkToplevel()
        check_window.title('Extracted Sheets List')
        check_window.grid_rowconfigure(0, weight=0)
        check_window.grid_rowconfigure(1, weight=1)
        check_window.grid_rowconfigure(2, weight=0)
        check_window.grid_columnconfigure(0, weight=1)
        check_window.grid_columnconfigure(1, weight=1)
        label_search = ctk.CTkLabel(check_window, text='Search Sheet:')
        label_search.grid(row=0, column=0, padx=10, pady=5, sticky='e')
        entry_search = ctk.CTkEntry(check_window)
        entry_search.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        frame_listbox = ctk.CTkFrame(check_window)
        frame_listbox.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky='nsew')
        scrollbar = ctk.CTkScrollbar(frame_listbox, orientation='vertical')
        scrollbar.pack(side='right', fill='y')
        listbox = tk.Listbox(frame_listbox, yscrollcommand=scrollbar.set, height=10)
        listbox.pack(side='left', fill='both', expand=True)
        scrollbar.configure(command=listbox.yview)

        def populate_listbox(sorted_sheets, sorted_dataframes):
            listbox.delete(0, tk.END)
            for sheet in sorted_sheets:
                listbox.insert(tk.END, sheet)

        def sort_ascending():
            sorted_pairs = sorted(zip(self.extracted_sheets, self.processed_dataframes))
            sorted_sheets, sorted_dataframes = zip(*sorted_pairs)
            populate_listbox(sorted_sheets, sorted_dataframes)

        def sort_descending():
            sorted_pairs = sorted(zip(self.extracted_sheets, self.processed_dataframes), reverse=True)
            sorted_sheets, sorted_dataframes = zip(*sorted_pairs)
            populate_listbox(sorted_sheets, sorted_dataframes)
        sort_ascending()

        def update_listbox(*args):
            search_query = entry_search.get().lower()
            filtered_pairs = [(sheet, df) for sheet, df in zip(self.extracted_sheets, self.processed_dataframes) if search_query in sheet.lower()]
            listbox.delete(0, tk.END)
            if filtered_pairs:
                filtered_sheets, filtered_dataframes = zip(*filtered_pairs)
                populate_listbox(filtered_sheets, filtered_dataframes)
            else:
                listbox.insert(tk.END, 'No matching sheets found.')
        entry_search.bind('<KeyRelease>', update_listbox)

        def on_listbox_double_click(event):
            selection = listbox.curselection()
            if selection:
                selected_sheet = listbox.get(selection[0])
                if selected_sheet == 'No matching sheets found.':
                    return
                try:
                    index = list(self.extracted_sheets).index(selected_sheet)
                    df = self.processed_dataframes[index]
                except ValueError:
                    return
                self.preview_data(df, title=f'Preview - {selected_sheet}')
        listbox.bind('<Double-Button-1>', on_listbox_double_click)
        ctk.CTkButton(check_window, text='Sort Ascending', command=sort_ascending).grid(row=2, column=0, padx=10, pady=5, sticky='ew')
        ctk.CTkButton(check_window, text='Sort Descending', command=sort_descending).grid(row=2, column=1, padx=10, pady=5, sticky='ew')
        check_window.transient(self.root)
        check_window.grab_set()
        self.root.wait_window(check_window)

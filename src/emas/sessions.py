"""Sessions callbacks for the application controller."""
import pickle
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class SessionMixin:

    def save_to_master_file(self, master_file_path):
        try:
            df = pd.read_excel(master_file_path)
            if 'Recon By' not in df.columns:
                df['Recon By'] = pd.NA
            if 'Recon Date' not in df.columns:
                df['Recon Date'] = pd.NA
            if 'SOA Date' not in df.columns:
                df['SOA Date'] = pd.NA
            new_row = {'Recon By': self.recon_data['Recon By'], 'Recon Date': self.recon_data['Recon Date'], 'SOA Date': self.recon_data['SOA Date']}
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            df.to_excel(master_file_path, index=False)
        except Exception as e:
            messagebox.showerror('Save Error', str(e))

    def load_session_state(self):
        session_file = filedialog.askopenfilename(title='Select Session File', filetypes=[('Pickle files', '*.pkl')])
        if session_file:
            try:
                with open(session_file, 'rb') as f:
                    state = pickle.load(f)
                    self.processed_dataframes = state.get('processed_dataframes', [])
                    self.extracted_hospitals = state.get('extracted_hospitals', [])
                    self.recon_data = state.get('recon_data', {})
                    self.extracted_sheets = list(state.get('extracted_sheets', []))
                    if len(self.extracted_sheets) != len(self.processed_dataframes):
                        raise ValueError('Legacy session has combined sheets; re-extract the original workbook.')
                    self.selected_file_path = state.get('selected_file_path', None)
                    self.sheet_df = state.get('sheet_df', None)
                    self.is_data_extracted = state.get('is_data_extracted', False)
                    self.is_data_saved = state.get('is_data_saved', True)
                if not self.processed_dataframes and (not self.extracted_hospitals) and (not self.extracted_sheets):
                    messagebox.showerror('No Data', 'The selected session file contains no data. Starting a new session.')
                    self.start_new_session()
                else:
                    self.setup_initial_gui()
                    self.is_data_saved = True
                    messagebox.showinfo('Session Loaded', 'Session data loaded successfully from the selected file.')
            except Exception as e:
                messagebox.showerror('Error', f'Failed to load session: {e}')
        else:
            messagebox.showinfo('No Selection', 'No session file selected. Load session canceled.')

    def save_session_state(self):
        session_file = filedialog.asksaveasfilename(title='Save Session As', defaultextension='.pkl', filetypes=[('Pickle files', '*.pkl')])
        if session_file:
            try:
                state = {key: getattr(self, key) for key in ('processed_dataframes', 'extracted_hospitals', 'extracted_sheets', 'recon_data', 'selected_file_path', 'sheet_df', 'is_data_extracted')}
                state['is_data_saved'] = True
                with open(session_file, 'wb') as f:
                    pickle.dump(state, f)
                messagebox.showinfo('Save Session', 'Session state saved successfully!')
                self.is_data_saved = True
                return True
            except Exception as e:
                messagebox.showerror('Error', f'Failed to save session: {e}')
        else:
            messagebox.showinfo('No Selection', 'Session save canceled.')

    def start_new_session(self):
        if not messagebox.askyesno('Confirm New Session', 'Clear the current extracted data?'):
            return
        self.processed_dataframes = []
        self.extracted_hospitals = []
        self.extracted_sheets = []
        self.selected_file_path = None
        self.sheet_df = None
        self.is_data_extracted = False
        self.is_data_saved = True
        if hasattr(self, 'entry_file_path'):
            self.entry_file_path.delete(0, tk.END)
            self.sheet_selection_combobox.configure(values=[])
            self.sheet_selection_combobox.set('')
        messagebox.showinfo('New Session', 'Starting a new session.')

    def on_close(self):
        if self.is_data_extracted and (not self.is_data_saved):
            result = messagebox.askyesnocancel('Quit', 'There is unsaved data. Would you like to save the session before exiting?')
            if result is None:
                return
            if result:
                if not self.save_session_state():
                    return
        self.root.destroy()

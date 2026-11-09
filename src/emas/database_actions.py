"""Database persistence callbacks for the application controller."""
from decimal import Decimal
import pandas as pd
from tkinter import messagebox
from .database import connection, DatabaseManager

def normalized(value, default=''):
    return default if pd.isna(value) else str(value)

class DatabaseActionsMixin:

    def insert_data_to_db(self):
        if not self.processed_dataframes:
            messagebox.showwarning('No Data', 'Please extract data first.')
            return
        conn = None
        try:
            conn = DatabaseManager.create_connection()
            for frame in self.processed_dataframes:
                rows = DatabaseManager.prepare_dataframe(frame.copy()).to_dict('records')
                for row in rows:
                    bill_number = normalized(row['Bill Number'])
                    if not bill_number:
                        continue
                    existing = connection.table_recon(bill_number)
                    if existing.empty:
                        DatabaseManager.insert_data(conn, row)
                        continue
                    stored = existing.iloc[0]
                    changes = {}
                    amount = Decimal(str(row['Bill Amount'])).quantize(Decimal('0.00'))
                    if Decimal(str(stored['BillAmount'])).quantize(Decimal('0.00')) != amount:
                        changes['BillAmount'] = str(amount)
                    mapping = {'ProviderID': 'ProviderID', 'ClaimID': 'ClaimID', 'DMSID': 'DMSID', 'Recon Date': 'ReconDate', 'ReconRemarks': 'ReconRemarks', 'Status': 'ReconStatus'}
                    for source, destination in mapping.items():
                        value = normalized(row[source], '0' if source in ('ProviderID', 'ClaimID') else '')
                        old = normalized(stored[destination])
                        if destination == 'ReconDate':
                            value = str(pd.to_datetime(value).date())
                            old = str(pd.to_datetime(old).date()) if old else ''
                        if old != value:
                            changes[destination] = value
                    if changes:
                        changes['ModifyBy'] = normalized(row['Recon By'])
                        assignments = ', '.join((f'[{name}] = ?' for name in changes))
                        query = 'UPDATE tbl_SOARecon SET ' + assignments + ', ModifyDate = GETDATE() WHERE BillNo = ?'
                        DatabaseManager.update_data(conn, query, tuple(changes.values()) + (bill_number,))
            messagebox.showinfo('Success', 'Database operations completed successfully.')
        except Exception as exc:
            messagebox.showerror('Database Error', str(exc))
        finally:
            if conn is not None:
                DatabaseManager.close_connection(conn)

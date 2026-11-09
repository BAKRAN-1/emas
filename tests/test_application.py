"""Regression checks using synthetic data; no production services are contacted."""
import importlib
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch
import pandas as pd

from src.emas.app import Application
from src.emas import claims, database, helpers, sessions, extraction, database_actions


class ApplicationTests(unittest.TestCase):
    def setUp(self):
        self.app = Application()

    def test_import_does_not_launch_ui_or_connect(self):
        with patch('customtkinter.CTk') as gui, patch.object(database, 'connect_database') as connect:
            importlib.reload(importlib.import_module('src.emas.app'))
            gui.assert_not_called()
            connect.assert_not_called()

    def test_excel_columns_and_dates(self):
        for index in (0, 25, 26, 51, 52, 701):
            self.assertEqual(helpers.letter_to_index(helpers.index_to_letter(index)), index)
        with self.assertRaises(ValueError):
            helpers.letter_to_index('A1')
        with patch.object(helpers, 'generate_random_string', return_value='ABC'):
            self.assertEqual(helpers.create_reconref('2026/01/02', '2026-03-04'), '02010403ABC')

    def test_sheet_data_remains_aligned_after_remove(self):
        first = pd.DataFrame({'ProviderName': ['Synthetic A']})
        second = pd.DataFrame({'ProviderName': ['Synthetic B']})
        self.app.add_extracted_data('A', first)
        self.app.add_extracted_data('B', second)
        self.app.remove_extracted_data('A')
        self.assertEqual(self.app.extracted_sheets, ['B'])
        pd.testing.assert_frame_equal(self.app.processed_dataframes[0], second)
        self.assertEqual(self.app.extracted_hospitals, ['Synthetic B'])
        self.assertFalse(self.app.is_data_saved)

    def test_session_round_trip_keeps_metadata_and_workbook(self):
        self.app.add_extracted_data('Demo', pd.DataFrame({'Bill Number': ['DEMO-1']}))
        self.app.recon_data = {'Recon By': 'Demo User', 'SOA Date': '2026-01-01'}
        self.app.selected_file_path = 'synthetic.xlsx'
        self.app.setup_initial_gui = MagicMock()
        with tempfile.TemporaryDirectory() as directory:
            path = str(Path(directory)/'session.pkl')
            with patch.object(sessions.filedialog, 'asksaveasfilename', return_value=path), patch.object(sessions, 'messagebox'):
                self.assertTrue(self.app.save_session_state())
            self.app.recon_data = {}
            with patch.object(sessions.filedialog, 'askopenfilename', return_value=path), patch.object(sessions, 'messagebox'):
                self.app.load_session_state()
            self.assertEqual(self.app.recon_data['Recon By'], 'Demo User')
            self.assertEqual(self.app.selected_file_path, 'synthetic.xlsx')
            self.assertTrue(self.app.is_data_extracted)
            self.assertTrue(self.app.is_data_saved)

    def test_cancel_save_on_close_keeps_window_open(self):
        self.app.is_data_extracted = True
        self.app.is_data_saved = False
        self.app.root = MagicMock()
        with patch.object(sessions.messagebox, 'askyesnocancel', return_value=True), patch.object(sessions.filedialog, 'asksaveasfilename', return_value=''), patch.object(sessions.messagebox, 'showinfo'):
            self.app.on_close()
        self.app.root.destroy.assert_not_called()

    def test_api_uses_configured_key_and_timeout(self):
        response = MagicMock()
        response.json.return_value = {'list': [{'id': 'synthetic-id'}]}
        with patch.dict(os.environ, {'EMAS_API_BASE_URL': 'https://example.invalid/api/v1', 'EMAS_API_KEY': 'test-only'}), patch.object(claims.requests, 'get', return_value=response) as get:
            self.assertEqual(claims.fetch_claim_data(123), 'synthetic-id')
            self.assertEqual(get.call_args.kwargs['timeout'], 30)
            self.assertEqual(get.call_args.kwargs['params']['where[0][value]'], '123')
        with patch.object(claims.requests, 'get') as get:
            self.assertIsNone(claims.fetch_claim_data(pd.NA))
            get.assert_not_called()

    def test_bill_lookup_is_parameterized_and_closes_connection(self):
        conn = MagicMock()
        conn.cursor.return_value.description = [('BillNo',)]
        conn.cursor.return_value.fetchall.return_value = [('DEMO',)]
        with patch.object(database, 'connect_database', return_value=conn):
            result = database.connection.table_recon("DEMO' OR 1=1 --")
        conn.cursor.return_value.execute.assert_called_once_with('SELECT * FROM tbl_SOARecon WHERE BillNo = ?', ("DEMO' OR 1=1 --",))
        conn.close.assert_called_once()
        self.assertEqual(result.iloc[0]['BillNo'], 'DEMO')

    def test_database_error_rolls_back_and_propagates(self):
        conn = MagicMock()
        conn.cursor.return_value.__enter__.return_value.execute.side_effect = RuntimeError('synthetic failure')
        with self.assertRaises(RuntimeError):
            database.DatabaseManager.update_data(conn, 'UPDATE demo SET value=?', ('test',))
        conn.rollback.assert_called_once()
        conn.commit.assert_not_called()

    def configure_extraction(self):
        self.app.second_bill_number_var = MagicMock()
        self.app.second_bill_number_var.get.return_value = False
        self.app.remove_non_numeric = MagicMock()
        self.app.remove_non_numeric.get.return_value = False
        self.app.recon_data = {'Recon By': 'Demo', 'Recon Date': '2026-01-01', 'SOA Date': '2026-01-01'}
        self.app.preview_data = MagicMock()

    def test_workbook_extraction_with_fake_sql_and_api(self):
        self.configure_extraction()
        sql = pd.DataFrame({'BillNo': ['DEMO-1'], 'ClaimID': [1], 'ProviderID': [2], 'ProviderName': ['Synthetic Provider']})
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory)/'demo.xlsx'
            pd.DataFrame({'Date': ['01/01/2026'], 'Number': ['DEMO-1'], 'Amount': [100]}).to_excel(workbook, index=False)
            with patch.object(extraction.connection, 'pandas_dataframe', return_value=sql), patch.object(extraction, 'fetch_claim_data', new=lambda claim: 'demo-id'), patch.object(extraction, 'messagebox') as messages:
                self.assertTrue(self.app.extract_data('A','B','C',workbook,'Sheet1'), str(messages.showerror.call_args))
                messages.showerror.assert_not_called()
        self.assertEqual(self.app.extracted_sheets, ['Sheet1'])
        self.assertEqual(self.app.processed_dataframes[0].iloc[0]['DMSID'], 'demo-id')

    def test_failed_replacement_retains_existing_sheet(self):
        self.configure_extraction()
        original = pd.DataFrame({'Bill Number': ['KEEP']})
        self.app.add_extracted_data('Sheet1', original)
        raw = pd.DataFrame([['Date','Number','Amount'], ['01/01/2026','DEMO-1',100]])
        with patch.object(extraction.pd, 'read_excel', return_value=raw), patch.object(extraction.connection, 'pandas_dataframe', side_effect=RuntimeError('offline')), patch.object(extraction, 'messagebox'):
            self.app.extract_data('A','B','C','fake.xlsx','Sheet1')
        self.assertEqual(self.app.extracted_sheets, ['Sheet1'])
        pd.testing.assert_frame_equal(self.app.processed_dataframes[0], original)

    def test_ui_construction_resolves_callbacks_without_opening_windows(self):
        module = importlib.import_module('src.emas.app')
        with patch.object(module, 'ctk'), patch.object(module, 'tk'), patch.object(module, 'ttk'):
            self.app.main_interface()
            self.app.root.mainloop.assert_called_once()

    def test_database_action_closes_connection_after_failure(self):
        self.app.processed_dataframes = [pd.DataFrame({'anything': ['synthetic']})]
        conn = MagicMock()
        with patch.object(database_actions.DatabaseManager, 'create_connection', return_value=conn), patch.object(database_actions.DatabaseManager, 'prepare_dataframe', side_effect=ValueError('synthetic error')), patch.object(database_actions, 'messagebox') as messages:
            self.app.insert_data_to_db()
        conn.close.assert_called_once()
        messages.showinfo.assert_not_called()
        messages.showerror.assert_called_once()

    def test_database_update_includes_multiple_changed_fields(self):
        row = {'Bill Number': 'DEMO-1', 'Bill Amount': '200.00', 'ProviderID': 2,
               'ClaimID': 3, 'DMSID': 'new-id', 'Recon Date': '2026/01/01',
               'ReconRemarks': 'demo remark', 'Status': 'Pending', 'Recon By': 'Demo'}
        self.app.processed_dataframes = [pd.DataFrame([row])]
        old = pd.DataFrame([{'BillAmount': 100, 'ProviderID': 1, 'ClaimID': 3,
                             'DMSID': 'old-id', 'ReconDate': '2026-01-01',
                             'ReconRemarks': '', 'ReconStatus': 'Pending'}])
        conn = MagicMock()
        with patch.object(database_actions.DatabaseManager, 'create_connection', return_value=conn), patch.object(database_actions.DatabaseManager, 'prepare_dataframe', side_effect=lambda frame: frame), patch.object(database_actions.connection, 'table_recon', return_value=old), patch.object(database_actions.DatabaseManager, 'update_data') as update, patch.object(database_actions, 'messagebox'):
            self.app.insert_data_to_db()
        query = update.call_args.args[1]
        for field in ('BillAmount', 'ProviderID', 'DMSID', 'ReconRemarks'):
            self.assertIn('[' + field + '] = ?', query)
        self.assertEqual(update.call_args.args[2][-1], 'DEMO-1')
        conn.close.assert_called_once()


if __name__ == '__main__':
    unittest.main()


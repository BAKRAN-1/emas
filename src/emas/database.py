"""SQL Server access. Connections are opened only when an action needs them."""
import json
import logging
from decimal import Decimal
from datetime import datetime
import pandas as pd
from .config import database_config_path

def connect_database(config_file=None):
    import pyodbc
    with open(config_file or database_config_path(), encoding='utf-8') as stream:
        config = json.load(stream)['db connection']

    def quote(value):
        return '{' + str(value).replace('}', '}}') + '}'
    values = {'DRIVER': config.get('driver', 'SQL Server'), 'SERVER': config['server'], 'DATABASE': config['database'], 'UID': config['username'], 'PWD': config['password']}
    return pyodbc.connect(';'.join((key + '=' + quote(value) for key, value in values.items())))

class connection:

    @staticmethod
    def _query(query, params=()):
        conn = connect_database()
        try:
            cursor = conn.cursor()
            try:
                cursor.execute(query, params)
                rows = cursor.fetchall()
                return pd.DataFrame.from_records(rows, columns=[column[0] for column in cursor.description])
            finally:
                cursor.close()
        finally:
            conn.close()

    @staticmethod
    def pandas_dataframe(bill_numbers):
        if not bill_numbers:
            return pd.DataFrame(columns=['BillNo', 'ClaimID', 'ProviderID'])
        placeholders = ','.join(('?' for _ in bill_numbers))
        return connection._query('SELECT * FROM VW_SOAReconInfo WHERE BillNo IN (' + placeholders + ')', bill_numbers)

    @staticmethod
    def table_recon(bill_number):
        return connection._query('SELECT * FROM tbl_SOARecon WHERE BillNo = ?', (bill_number,))

    @staticmethod
    def convert_data(data):
        return [{key: str(value) if isinstance(value, (Decimal, datetime)) else value for key, value in row.items()} for row in data]

class DatabaseManager:

    @staticmethod
    def create_connection(config_file=None):
        return connect_database(config_file)

    @staticmethod
    def prepare_dataframe(df):
        df['Bill Amount'] = pd.to_numeric(df['Bill Amount'], errors='coerce')
        df['ClaimAmount'] = pd.to_numeric(df['ClaimAmount'], errors='coerce')
        df['ClaimID'] = pd.to_numeric(df['ClaimID'], errors='coerce').astype(pd.Int64Dtype())
        df['ProviderID'] = pd.to_numeric(df['ProviderID'], errors='coerce').astype(pd.Int64Dtype())
        varchar_fields = ['Bill Number', 'BillNo', 'ProviderName', 'InvoiceNo', 'CompanyName', 'PayOrderNo', 'Status', 'Recon By', 'Reconrefno']
        for field in varchar_fields:
            df[field] = df[field].astype(str)
        return df

    @staticmethod
    def insert_data(connection, row):
        if connection is None:
            raise RuntimeError('Connection is not established.')
        try:
            with connection.cursor() as cursor:
                logging.debug('Inserting reconciliation row')
                cursor.execute('INSERT INTO tbl_SOARecon ( CreateDate, BillNo,\n                    BillAmount, ReconReferenceNo, ReconDate, SOADate,  ProviderID, ClaimID,\n                    DMSID, ReconBy, ModifyDate, ModifyBy, ReconRemarks,\n                    ReconStatus\n                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)', (str(row['CreateDate']), str(row['Bill Number']) if str(row['Bill Number']) != 'nan' else '', str(row['Bill Amount']), str(row['Reconrefno']), str(row['Recon Date']).replace('/', '-'), str(row['SOA Date']).replace('/', '-'), str(row['ProviderID']) if str(row['ProviderID']) != 'None' else '0', str(row['ClaimID']) if str(row['ClaimID']) != 'None' else '0', str(row['DMSID']), str(row['Recon By']), str(row['ModifyDate']), str(row['ModifyBy']), str(row['ReconRemarks']), str(row['Status']) if str(row['Status']) != 'nan' else ''))
            connection.commit()
            logging.info('Data inserted successfully.')
        except Exception as e:
            logging.error(f'Error inserting data: {e}')
            connection.rollback()
            raise

    @staticmethod
    def update_data(connection, query, updatedata):
        if connection is None:
            raise RuntimeError('Connection is not established.')
        try:
            with connection.cursor() as cursor:
                cursor.execute(query, updatedata)
            connection.commit()
            logging.info('Data updated successfully.')
        except Exception as e:
            logging.error(f'Error updating data: {e}')
            connection.rollback()
            raise

    @staticmethod
    def close_connection(connection):
        if connection:
            connection.close()
            logging.info('Connection closed.')

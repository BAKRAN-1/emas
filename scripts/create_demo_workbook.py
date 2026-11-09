"""Generate an Excel workbook from the committed synthetic sample."""
from pathlib import Path
import pandas as pd


def main():
    directory = Path(__file__).resolve().parents[1] / 'data' / 'sample'
    frame = pd.read_csv(directory / 'invoices.csv')
    frame.to_excel(directory / 'invoices.xlsx', index=False, sheet_name='Demo')
    print('Created:', directory / 'invoices.xlsx')


if __name__ == '__main__':
    main()

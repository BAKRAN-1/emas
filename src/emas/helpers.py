"""Pure formatting, identifier and Excel-column helpers."""
import random
import string
import re
import pandas as pd

def generate_random_string(length=3):
    letters = string.ascii_uppercase
    return ''.join((random.choice(letters) for i in range(length)))

def create_reconref(soa_date, recon_date):
    if pd.isna(soa_date) or pd.isna(recon_date):
        return ''
    soa = pd.to_datetime(soa_date, yearfirst=True)
    recon = pd.to_datetime(recon_date, yearfirst=True)
    day_month_soa = soa.strftime('%d%m')
    day_month_recon = recon.strftime('%d%m')
    random_str = generate_random_string()
    base_key = f'{day_month_soa}{day_month_recon}{random_str}'
    return base_key[:11]

def letter_to_index(col_letter):
    col_letter = str(col_letter).strip().upper()
    if not re.fullmatch('[A-Z]+', col_letter):
        raise ValueError('Use an Excel column letter such as A or AA.')
    index = 0
    for char in col_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def format_bill_amount(amount):
    """
    Format the bill amount to ensure it has no more than 2 decimal points.
    """
    try:
        formatted_amount = round(float(amount), 2)
        return f'{formatted_amount:.2f}'
    except ValueError:
        return '0.00'

def remove_letters_from_number(bill_number):
    """
    Remove all letters from a bill number but keep non-numeric characters.
    """
    return re.sub('[a-zA-Z]', '', bill_number)

def remove_numbers_after_decimal(bill_number):
    """
    Remove all numbers after the decimal point in a bill number.
    """
    return re.sub('\\.\\d+', '', bill_number)

def combine_bill_numbers(bill_number_1, bill_number_2, remove_letters=False):
    """
    Combine two bill numbers based on the following rules:
    1. If bill_number_1's numeric value is '1', combine with a leading '0' between.
    2. If bill_number_1's numeric value is '2', combine with a space.
    3. If remove_letters is True, remove letters from both bill numbers before combining.
    4. Remove numbers after the decimal points from both bill numbers before combining.
    """
    bill_number_1 = remove_numbers_after_decimal(bill_number_1)
    bill_number_2 = remove_numbers_after_decimal(bill_number_2)
    if remove_letters:
        bill_number_1 = remove_letters_from_number(bill_number_1)
        bill_number_2 = remove_letters_from_number(bill_number_2)
    bill_number_1_value = remove_letters_from_number(bill_number_1)
    if bill_number_1_value == '1':
        combined = f'{bill_number_1}0{bill_number_2}'
    elif bill_number_1_value == '2':
        combined = f'{bill_number_1} {bill_number_2}'
    else:
        combined = f'{bill_number_1}/{bill_number_2}'
    return combined

def index_to_letter(index):
    if index < 0:
        raise ValueError('Column index must be non-negative.')
    result = ''
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result

"""
This code can help you to send automatic messages in Whatsapp.

Provide a csv or xls* file containing phone numbers in a column to send an
equal message to all of them.

Make sure that the area/state code is informed in your
numbers, otherwise the message will not be delivered.

Note: The code takes some time to send the message because of the
pywhatkit external library delay. If you are going to send a lot of
messages, it will be better to 'leave' the code running because it
will take some time.
"""

import warnings
import pandas as pd
import re
import pywhatkit as kit
import os
import pyautogui
from datetime import datetime as dt
from time import sleep

warnings.filterwarnings('ignore')


def warning_print(string):
    print('\033[1;31;40m' + string + '\033[0m')  # Colors the terminal in red


def get_file():
    file = input(r'Please, insert the .xls* or .csv file path that contains the phone numbers: ')
    _, file_extension = os.path.splitext(file)

    while not (os.path.exists(file)) or (file_extension[0:4] not in ['.csv', '.xls']):

        if not os.path.exists(file):
            warning_print(f'This file path does not exist in your system: "{file}". Try again.')
        else:
            warning_print(f'You must select a .csv or .xls* file. Try again.')

        file = input(r'Please, insert the .xls* or .csv file path that contains the phone numbers: ')
        _, file_extension = os.path.splitext(file)

    return file, file_extension


def clean_number(number):
    return re.sub(r'\D', '', number)  # Removes all extra symbols in the number, like "(", ")", "-", "+"


def format_phone(number, country_code=''):
    phone = str(number)
    phone = clean_number(phone)

    if country_code:
        country_code = clean_number(country_code)
        len_country_code = len(country_code)
        if phone[:len_country_code] != country_code:
            phone = country_code + phone

    phone = '+' + phone

    return phone


def get_numbers_from_file(file_name, file_extension):
    if file_extension == 'csv':
        df = pd.read_csv(file_name)
    else:
        df = pd.read_excel(file_name)

    df.columns = df.columns.str.upper()
    columns = df.columns.to_list()

    phone_column = input('Name of the phone column: ').upper()
    while phone_column not in columns:
        warning_print(f'There is no column "{phone_column}" in the file. Try again.')
        print(f'Columns found in your file:\n{columns}\n')
        phone_column = input('Name of the phone column: ').upper()
    print()
    country_code = input('Sometimes there are numbers without the country code.\nInform a number to '
                         'be used as default for the country code or press enter to skip: ')
    phones = df[phone_column].dropna().apply(format_phone, country_code=country_code).to_list()
    return phones


def get_message():
    message = input('Message:\n')

    while len(message) < 3:
        warning_print('You need to type at least 3 characters in your message')
        message = input('Message:\n')

    print('Are you sure you want to send this message to every phone?\n'
          f'"{message}"')

    confirm = input('1 to confirm, 0 to type the message again: ')
    while confirm not in ['0', '1']:
        confirm = input('1 to confirm, 0 to type the message again: ')

    confirm = int(confirm)  # Only converted to int here to prevent typing alphabetic values in input

    if not confirm:
        return get_message()
    else:
        return message


def send_message(phones, message, verbose=False):
    while len(phones) >= 1:
        hour = dt.now().time().hour
        minute = dt.now().time().minute + 2
        if verbose:
            print(f'Sending message to {phones[0]}...', end=' ')
        try:    # In case the there is no Whatsapp associated with the number
            kit.sendwhatmsg(phone_no=phones[0], message=message, time_hour=hour, time_min=minute, wait_time=10)
            sleep(5)
            pyautogui.hotkey('ctrl', 'w')
            print('Sent!')
        except Exception:
            print(f'An error has occurred and the message was not sent to {phones[0]}')
            pass
        del phones[0]





if __name__ == '__main__':
    print('--- Whatsapp Automatic Messenger ---')
    print('Welcome!')
    print('*Before you start, make sure you are logged in https://web.whatsapp.com/\n')

    file, file_extension = get_file()
    numbers = get_numbers_from_file(file_name=file, file_extension=file_extension)
    print(f'Your message will be sent to these following numbers:\n{numbers}')
    print()
    message = get_message()
    print(message)
    send_message(phones=numbers, message=message, verbose=True)
    print('Finish!')

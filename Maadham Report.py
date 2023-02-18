# -*- coding: utf-8 -*-
"""
Created on Wed Feb 15 15:26:51 2023

@author: Naveenkumar
"""

from openpyxl import load_workbook
import PySimpleGUI as R
from datetime import datetime


R.theme('DarkAmber')

layout = [[R.Text('Name'),R.Push(), R.Input(key='NAME')],
          [R.Text('Part Number'),R.Push(), R.Input(key='Part Number')],
          [R.Text('Address'),R.Push(), R.Input(key='Address')],
          [R.Text('TEL:'),R.Push(), R.Input(key='NUMBER')],
          [R.Button('Submit'), R.Button('Close')]]

window = R.Window('Data Entry', layout, element_justification='center')

while True:
    event, values = window.read()
    if event == R.WIN_CLOSED or event == 'Close':
        break
    if event == 'Submit':
        try:
            wb = load_workbook('Monthly_Reports.xlsx')
            sheet = wb['Sheet1']
            ID = len(sheet['ID']) + 1
            time_stamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            data = [ID, values['NAME'], values['Part Number'],values['Address'],values['NUMBER'], time_stamp]

            sheet.append(data)

            wb.save('Monthly_Reports.xlsx')

            window['NAME'].update(value='')
            window['Part Number'].update(value='')
            window['Address'].update(value='')
            window['NUMBER'].update(value='')
            window['NAME'].set_focus()

            R.popup('Success', 'Data Saved')
        except PermissionError:
            R.popup('File in use', 'File is being used by another User.\nPlease try again later.')
        


window.close()
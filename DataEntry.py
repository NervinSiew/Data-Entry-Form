from pickle import FALSE
import PySimpleGUI as sg
import pandas as pd

# Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = 'Load List Testing Data Entry.xlsx'
df = pd.read_excel(EXCEL_FILE)

# Each list within the Layout List represents a column in the GUI
layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Product Category', size=(15,1)), sg.InputText(key='Product Category')], # Size=(15,1) means it is 15 characters wide and 1 character tall
    [sg.Text('Product Name', size=(15,1)), sg.InputText(key='Product Name')],
    [sg.Text('Serial No.', size=(15,1)), sg.InputText(key='Serial No.')],
    [sg.Text('Product Manufactured Date', size=(15,1)), sg.InputText(key='Product Manufactured Date')],
    [sg.Text('Qty Available ', size=(15,1)), sg.InputText(key='Qty Available ')],
    [sg.Text('Product Not Available', size=(15,1)), sg.Combo(['FALSE', 'TRUE'], key='Product Not Available')],  
    [sg.Text('Company Loan to', size=(15,1)), sg.InputText(key='Company Loan to')],
    [sg.Text('Loan Out Date', size=(15,1)), sg.InputText(key='Loan Out Date')],
    [sg.Text('Colleague in Charge', size=(15,1)), sg.InputText(key='Colleague in Charge')],
    [sg.Text('Remarks', size=(15,1)), sg.InputText(key='Remarks')],

    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event == 'Clear':
        clear_input()

    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index = FALSE)
        sg.popup('Data Saved!')
        clear_input()
window.close()
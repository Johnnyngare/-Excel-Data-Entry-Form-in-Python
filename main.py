import PySimpleGUI as sg
import pandas as pd

sg.theme('DarkTeal9')

EXCEL_FILE = 'Data.xlsx'

# Try to read the existing Excel file, or create an empty DataFrame if it doesn't exist
try:
    existing_data = pd.read_excel(EXCEL_FILE)
except FileNotFoundError:
    existing_data = pd.DataFrame(columns=['Name', 'Favorite Colour', 'German', 'Spanish', 'English', 'No_of_Children'])

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15, 1)), sg.InputText(key='Name')],
     [sg.Text('City', size=(15, 1)), sg.InputText(key='City')],
    [sg.Text('Favorite Colour', size=(15, 1)), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Colour')],
    [sg.Text('I speak', size=(15, 1)),
     sg.Checkbox('German', key='German'),
     sg.Checkbox('Spanish', key='Spanish'),
     sg.Checkbox('English', key='English')
     ],
    [sg.Text('No. of Children', size=(15, 1)), sg.Spin([i for i in range(0, 16)], key='No_of_Children')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

# Display the window
window = sg.Window('Simple data entry form', layout)

def clear_input():
 for key in values:
    window[key]('')
 return None 

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
       clear_input()
    if event == 'Submit':
        existing_data = existing_data._append(values, ignore_index=True)
        existing_data.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved')
        clear_input()

window.close()


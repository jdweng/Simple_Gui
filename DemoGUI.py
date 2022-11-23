import pandas as pd
import PySimpleGUI as sg
import openpyxl


#sg.theme_previewer()
sg.theme('LightBlue6')
#sg.popup('gooey gui')

layout = [
    [sg.Text('My Info')],
    [sg.Text('Name',size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('Age',size=(10,1)), sg.InputText(key='Age')],
    [sg.Text('Where will you spend most of Break?',size=(30,1)),
     sg.Combo(['EMU', 'Home', 'Other'], key = 'Location')],
    [sg.Text('Plans for break',size=(15,1)),
    sg.Checkbox('Relax',key='Relax'),
    sg.Checkbox('Study',key='Study'),
    sg.Checkbox('Other',key='Other')],
    [sg.Text('Current excitement for break', size=(15,1)),
         sg.Spin([i for i in range (0,10)], initial_value=0,key='Excitement')],
    [[sg.Text('On a scale from 1-100, how much do you wish it was Christmas instead?')],
      sg.Slider(range=(1,100),
         default_value=50,
         size=(20,15),
         orientation='horizontal',
         font=('Helvetica', 12))],
    [sg.Submit(),sg.Exit()]
    ]

window = sg.Window('Simple data form',layout)
#window.read()

try:
    df = pd.read_excel('classForm.xlsx')
except:    
    writer = pd.ExcelWriter('classForm.xlsx',engine='xlsxwriter')
    writer.close()
    df = pd.read_excel('classForm.xlsx')


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        new_record=pd.DataFrame(values, index=[0])
        df=pd.concat([df, new_record],ignore_index=True)
        df.to_excel('classForm.xlsx',index=False)
        sg.popup('Data Saved!')

print(df)

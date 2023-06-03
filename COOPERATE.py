import PySimpleGUI as sg
import pandas as pd
import os
import xlwings as xw
import openpyxl as ox

def update():
    files=[os.path.abspath(os.path.join(os.getcwd(), p)) for p in os.listdir(os.getcwd()) if p.endswith(('xlsx'))]
    for i in files:
        if i[-11:]=="report.xlsx":
               if os.path.isfile(i):
                    files.remove(i)
                    break


    combined=pd.DataFrame()
    for file in files:
          file = pd.read_excel(file,header=None)
          combined=pd.concat([combined, file]) 
    wb=ox.load_workbook(i)
    for ir in range (0,len(combined)):
          for ic in range (0,len(combined.iloc[ir])):
               wb["Sheet1"].cell(1+ir,1+ic).value=combined.iloc[ir][ic]
    wb.save(i)
    text_elem = window['-text-']
    text_elem.update("Results: {}".format('Complete'))

layout = [[sg.Button('Create a new report about clients',enable_events=True, key='-FUNCTION-', font='Helvetica 16')],
        [sg.Text('Results:', size=(25, 1), key='-text-', font='Helvetica 16')]]

window = sg.Window('Reporting', layout, size=(350,100))

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == '-FUNCTION-':
        update()
vba_book = xw.Book("report.xlsx")


window.close()

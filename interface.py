

# from DDkgta.HelpProfkom.HelpProfkomBot.HelpProfkonBot import send_count_doclad
# from DDkgta.HelpProfkom.HelpProfkomBot.HelpProfkonBot import send_people_info
# from DDkgta.HelpProfkom.HelpProfkomBot.HelpProfkonBot import get_victory_fakyltet, send_mesto_quanlity_indicators
import PySimpleGUI as sg
import traceback

from typing import Text
from PySimpleGUI.PySimpleGUI import InputOptionMenu, WIN_CLOSED
from HelpProfkonBot import *
from termcolor import colored

sg.theme('DarkAmber')
a = ["МТФ", "ФАиЭ", "ФЭиМ", "ЭМК"]
layout = [
    [sg.Text('Название секции'), sg.InputText(key='section')],
    [sg.Text('Кафедра'), sg.InputText(key='kafedra_section'), sg.InputOptionMenu(a,'Факультет', size=(10,20), key="section_fakyltet")],
    [sg.Text('Число студентов принявших участие в подготовке секции'), sg.InputText()],
    [[sg.Text('МТФ'), sg.InputText(key='МТФ', size=(5,20))],],
    [[sg.Text('АиЭ'), sg.InputText(key='ФАиЭ', size=(5,20))],],
    [[sg.Text('ЭиМ'), sg.InputText(key='ФЭиМ', size=(5,20))],],
    [[sg.Text('ЭМК'), sg.InputText(key='ЭМК', size=(5,20))],],
    [[sg.Text('Призовые места:')],],
    [[sg.Text('1 Место: ')],],
    [[sg.Text('ФИО'), sg.InputText(key='1mesto_fio', size=(30,20)), 
      sg.Text('Группа'), sg.InputText(key='1mesto_group', size=(10,20)), 
      sg.Text('Факультет'), sg.InputOptionMenu(a,'Факультет', size=(10,20), key="1mesto_fakyltet")
    ]],
    [[sg.Text("Название работы"), sg.InputText(key='1mesto_nameWork', size=(30,20)), sg.Text('Пол'),sg.Checkbox('M', key='1mesto_Male'), sg.Checkbox('Ж', key='1mesto_She'),sg.Submit('+ еще 1е место')]],

    [[sg.Text('----------------------------------------------------------------------------------------------------------------------')]],

    [[sg.Text('2 Место: ')],],
    [[sg.Text('ФИО'), sg.InputText(key='2mesto_fio', size=(30,20)), 
      sg.Text('Группа'), sg.InputText(key='2mesto_group', size=(10,20)), 
      sg.Text('Факультет'), sg.InputOptionMenu(a,'Факультет', size=(10,20), key="2mesto_fakyltet")
    ]],
    [[sg.Text("Название работы"), sg.InputText(key='2mesto_nameWork', size=(30,20)), sg.Text('Пол'),sg.Checkbox('M', key='2mesto_Male'), sg.Checkbox('Ж', key='2mesto_She'), sg.Submit('+ еще 2е место')]],
    
    [[sg.Text('----------------------------------------------------------------------------------------------------------------------')]],
    
    [[sg.Text('3 Место: ')],],
    [[sg.Text('ФИО'), sg.InputText(key='3mesto_fio', size=(30,20)), 
      sg.Text('Группа'), sg.InputText(key='3mesto_group', size=(10,20)), 
      sg.Text('Факультет'), sg.InputOptionMenu(a,'Факультет', size=(10,20), key="3mesto_fakyltet")
    ]],
    [[sg.Text("Название работы"), sg.InputText(key='3mesto_nameWork', size=(30,20)), sg.Text('Пол'), sg.Checkbox('M', key='3mesto_Male'), sg.Checkbox('Ж', key='3mesto_She'), sg.Submit('+ еще 3е место')]],
    
    [sg.Output(size=(100, 20))],
    [sg.Submit(), sg.Cancel(), sg.Submit('Test'), sg.Submit('Победители')]  
]  

win2 = sg
def get_layout():
    victoris = get_victory_fakyltet()

    layout2 = [ 
        [sg.Text('1 место:    '), sg.Text(f'{victoris[0][0]}') ],
        [sg.Text('2 место:    '), sg.Text(f'{victoris[1][0]}') ],
        [sg.Text('3 место:    '), sg.Text(f'{victoris[2][0]}') ],
        [sg.Text('4 место:    '), sg.Text(f'{victoris[3][0]}') ],
    ]
    return layout2
window2 = win2.Window('Победители', get_layout(), size=(200,100))

def send_for_mesto(mesto: int) :
    send_mesto_inSection(values['section'], values['section_fakyltet'], mesto=mesto)
            
    send_ratio_section_inMesto_and_fakyltet(
        section = values['section'],
        fakyltet= values['section_fakyltet'],
        ratio   = get_ves_ratio_section(section= values['section']),
        mesto   = mesto 
    )

    send_people_info(
        fio=values[f'{mesto}mesto_fio'],
        fakyltet=values[f'{mesto}mesto_fakyltet'],
        kafedra=values['kafedra_section'],
        groups=values[f'{mesto}mesto_group'],
        mesto=mesto,
        section=values['section'],
        work=values[f'{mesto}mesto_nameWork'],
        sex= 'М' if values[f'{mesto}mesto_Male'] else 'Ж'
    )


window = sg.Window('Кассиопея', layout)

# TODO сделать вывод в таблицу + грамоты или просто вывод  мест факультетов
while True:   
    try:                          # The Event Loop
        event, values = window.read()
        print(event, values) #debug
        # print(colored('[OK]', 'green'))
        # print(values['section'], values['section_fakyltet'], values['1mesto_fakyltet'])
        if event == 'Submit':
            fak = []
            fak.append(values['ФАиЭ'])
            fak.append(values['МТФ'])
            fak.append(values['ФЭиМ'])
            fak.append(values['ЭМК'])
    
            send_count_doclad(
                section = values['section'],
                fakyltet= values['section_fakyltet'],
                count   = fak,
                kafedra = values['kafedra_section'])
            send_mesto_quantitative_indicators()
            send_mesto_quanlity_indicators()

            if values['1mesto_fio'] != '':
                send_for_mesto(mesto=1)
             
            if values['2mesto_fio'] != '':
                send_for_mesto(mesto=2)
             
            if values['3mesto_fio'] != '':
                send_for_mesto(mesto=3)

        if event == "Победители":
            event, values = window2.read()
        # if event == 'Test':
            # window['3mesto_fio'].set_tooltip('New Tooltip')
            # window['3mesto_fio'].update('')
            
        if event in ( 'Exit', 'Cancel', sg.WIN_CLOSED ):
            break
    except Exception as e:
        print(e, traceback.format_exc())

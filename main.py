import PySimpleGUI as sg
import And
import Max
import TPK

def main():
    layout = [
        [sg.Button('Данные о мигрантах', key='And', font='Arial 18')],
        [sg.Button('Данные о браках', key='Max', font='Arial 18')],
        [sg.Button('Данные о рынке жилья', key='TPK', font='Arial 18')]
    ]
    window = sg.Window('Главная форма', layout, size=(400, 250), finalize=True)

    while True:
        event, _ = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        if event == 'And':
            window.hide()
            And.run()
            window.un_hide()
        elif event == 'Max':
            sg.popup('...')
        elif event == 'TPK':
            sg.popup('...')
    window.close()

if __name__ == '__main__':
    main()
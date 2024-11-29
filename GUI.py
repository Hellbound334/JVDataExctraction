import PySimpleGUI as sg


layout = [[sg.Text("Choose an action")],
          [sg.Button('Data extract')],
          [sg.Button('Power and IV calculation')],
          [sg.Button('End program session')] ]

window = sg.Window(title='J-V charasteristics extraction program', layout=layout)

while True:
    event, values = window.read()
    if event == 'Data extract':
        import data_extraction_program
    elif event == 'Power and IV calculation':
        import plotting_power
    elif event == 'End program session' or event == sg.WIN_CLOSED:
        break

window.close()
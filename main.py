#import docx
import os
import PySimpleGUI as sg

# sg.theme('DarkTeal2')
# layout = [[sg.T("")], [sg.Text("Choose a folder: "), sg.Input(key="-IN2-" ,change_submits=True), sg.FolderBrowse(key="-IN-")],[sg.Button("Submit")]]
# window = sg.Window('My File Browser', layout, size=(600,150))
# while True:
#     event, values = window.read()
#     print(values["-IN2-"])
#     if event == sg.WIN_CLOSED or event=="Exit":
#         break
#     elif event == "Submit":
#         print(values["-IN-"])


def select_folder():
    sg.theme('DarkTeal2')
    layout = [[sg.T("")],
              [sg.Text("Choose a folder: "), sg.Input(key="-IN2-", change_submits=True), sg.FolderBrowse(key="-IN-")],
              [sg.Button("Submit")]]
    window = sg.Window('My File Browser', layout, size=(600, 150))
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "Exit":
            break
        elif event == "Submit":
            print(values["-IN-"])
    return values


folder = select_folder()
print(folder)







# def select_folder():
#     left_col = [[sg.Text('Folder'), sg.In(size=(25, 1), enable_events=True, key='-FOLDER-'), sg.FolderBrowse()]]
#     layout = [sg.Column(left_col, element_justification='c')]
#     window = sg.Window('Multiple Format Image Viewer', layout, resizable=True)
#     while True:
#         event, values = window.read()
#         if event in (sg.WIN_CLOSED, 'Exit'):
#             break
#         if event == '-FOLDER-':
#             folder = values['-FOLDER-']
#     return folder
#
#
# def list_of_files():
#
#     filelist = os.listdir(filepath)
#
#
# folder = select_folder()
# print (folder)
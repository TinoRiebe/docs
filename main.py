from docx import Document
import os
import PySimpleGUI as sg


def select_folder():
    sg.theme('DarkTeal2')
    # combobox = sg.InputCombo(values=['', ], size=(10, 1), key='lang')

    layout = [[sg.T("")],
              [sg.Text("Choose a folder: "), sg.In(default_text='E:/', key="-IN2-", change_submits=True), sg.FolderBrowse(key="-IN-")],
              [sg.Text('search'), sg.In(default_text='test', size=(10,1), key='search', change_submits=True), sg.Checkbox('cAse SenSItive?:', default=False, key="case")],

              [sg.Text('replace'), sg.In(default_text='TEST', size=(10,1),key='replace', change_submits=True)],
              [sg.Button("Submit")]]
    window = sg.Window('My File Browser', layout, size=(600, 150))
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "Exit":
            break
        elif event == "Submit":

            selected_folder = (values["-IN-"])
            files_in_folder = os.listdir(selected_folder)
            filelist = []
            for file in files_in_folder:

                if 'doc' in (file.split('.')[-1]):
                    # print(file.split('.')[0])
                    filelist.append(file)
            looking_for = values['search']
            replace_with = values['replace']
            case = values['case']
            window.close()
    return selected_folder, filelist, looking_for, replace_with, case


def manipulate_doc(doc_folder, doc_files, looking, replace, case_sensitive):
    document = Document(os.path.join(doc_folder, doc_files))
    for paragraph in document.paragraphs:
        inline = paragraph.runs
        for i in range(len(inline)):
            text = inline[i].text
            sep = ' '
            phrasen = inline[i].text.split()
            phrasen_new = []
            cnt = 0
            if not case_sensitive:
                for k in range(len(phrasen)):
                    if phrasen[k].lower() == looking:
                        phrasen[k] = replace
                        cnt = 1
                if cnt == 1:
                    phrasen_new = sep.join(phrasen)
                    text = text.replace(text, phrasen_new)
                    inline[i].text = text
            else:
                for k in range(len(phrasen)):
                    if phrasen[k] == looking:
                        phrasen[k] = replace
                        cnt = 1
                if cnt == 1:
                    phrasen_new = sep.join(phrasen)
                    text = text.replace(text, phrasen_new)
                    inline[i].text = text
    document.save(os.path.join(folder, 'probe1.docx'))


if __name__ == '__main__':
    try:
        folder, files, looking_for, replace_with, case = select_folder()
        manipulate_doc(folder, files[0], looking_for, replace_with, case)
    except UnboundLocalError:
        print('stop')


    #folder = 'e:/'
    #files = 'probe.docx'
    #print(os.path.join(folder, files))
    #manipulate_doc(folder, files)

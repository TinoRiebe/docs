from docx import Document
import os
import PySimpleGUI as sg


def select_folder():
    sg.theme('DarkTeal2')
    # combobox = sg.InputCombo(values=['', ], size=(10, 1), key='lang')

    layout = [[sg.T("")],
              [sg.Text("Choose a folder: "), sg.In(default_text='E:/', key="-IN2-", change_submits=True),
               sg.FolderBrowse(key="-IN-")],
              [sg.Text('search'), sg.In(default_text='test', size=(10, 1), key='search', change_submits=True),
               sg.Checkbox('cAse SenSItive?:', default=False, key="case")],
              [sg.Text('replace'), sg.In(default_text='TEST', size=(10, 1), key='replace', change_submits=True)],
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
            print(filelist)
    return selected_folder, filelist, looking_for, replace_with, case


def manipulate_doc(doc_folder, doc_files, looking, replace, case_sensitive):
    for j in range(len(doc_files)):
        document = Document(os.path.join(doc_folder, doc_files[j]))
        for paragraph in document.paragraphs:
            inline = paragraph.runs
            for i in range(len(inline)):
                text = inline[i].text
                sep = ' '
                phrases = inline[i].text.split()
                cnt = 0
                if not case_sensitive:
                    for k in range(len(phrases)):
                        if phrases[k].lower() == looking:
                            phrases[k] = replace
                            cnt = 1
                    if cnt == 1:
                        phrases_new = sep.join(phrases)
                        text = text.replace(text, phrases_new)
                        inline[i].text = text
                else:
                    for k in range(len(phrases)):
                        if phrases[k] == looking:
                            phrases[k] = replace
                            cnt = 1
                    if cnt == 1:
                        phrases_new = sep.join(phrases)
                        text = text.replace(text, phrases_new)
                        inline[i].text = text
        doc_name = doc_files[j].split('.')[0] + '_new.docx'
        document.save(os.path.join(folder, doc_name))


if __name__ == '__main__':
    try:
        folder, files, searching, replacing, camelcase = select_folder()

        # def convertToListOfLists(toConvert):
        #     listOfLists = []
        #     for t in toConvert:
        #         temp = []
        #         temp.append(t.datenbestand)
        #         temp.append(t.aktenzeichen)
        #         temp.append(t.markendarstellung)
        #         temp.append(t.aktenzustand)
        #         listOfLists.append(temp)
        #     print(listOfLists)
        #     return listOfLists


        manipulate_doc(folder, files, searching, replacing, camelcase)
    except UnboundLocalError:
        print('stop')

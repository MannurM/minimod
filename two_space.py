import os
import docx
import docx2txt
from docx import Document

path_folder = 'C:\\Users\\User\\PycharmProjects\\minimod\\Upload_folder'  # input()
# path_folder = '/home/mannur/PycharmProjects/minimod/Upload_folder'

# symbol = '/'  #
symbol = '\\'
print('Yfxfkb')
for file_name in os.listdir(path_folder):
    print('1')
    if file_name[-5:] != '.docx':
        continue

    print(file_name)
    list_new = []
    doc = Document(path_folder + symbol + file_name)
    all_paragr = doc.paragraphs

    for par in all_paragr:
        # run = par.add_run()
        text = par.text
        new_text = text

        list_index = []
        # text = text.lstrip()

        for index, txt in enumerate(text):
            if text[index].isalpha() or text[index].isdigit():
                # print('exit')
                break
            elif text[:1] == ' ':
                text = text.lstrip()
        par.text = text

    doc.save(path_folder + symbol + file_name)
import docx
import docx2txt
import os

path_folder = 'C:\\Users\\User\\PycharmProjects\\minimod\\Upload_folder'
for file_name in os.listdir(path_folder):
    rezult_text = docx2txt.process(path_folder + '\\' + file_name)
    new_file_name = path_folder + '\\' + 'new_' + '.txt'
    print(rezult_text)
    doc_new = open(new_file_name, 'w')
    doc_new.write(rezult_text)
    print('end!!!!!')
    doc_new.close()




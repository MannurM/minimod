# utf-8
# план задание
# получить файл из папки, прочитать его, найти строку с номером приказа, заменить строку, сохранить файл


import os
import shutil
import zipfile

import docx


def add_path_folder():
    """ввести путь до папки с файлами, прочитать папку, в цикле взять файл, прочитать его по параграфам"""
    print('введите путь до папки с файлами')
    path_folder = 'C:\\Users\\User\\PycharmProjects\\minimod\\Upload_folder'  # input()
    if path_folder:
        return path_folder


def search_replace(doc, search_string, replace_string):
    pass
    # doc.add_tables

    # for paragraph in doc.paragraphs:
    #     if paragraph == replace_string:
    #         break
    #     elif paragraph == search_string:
    #         paragraph = replace_string
    # return doc


def input_data():
    """ввод данных для изменения в файле"""
    print('введите текст поиска')
    search_string = input()
    print('введите текст для замены')
    replace_string = input()
    return search_string, replace_string


def unzip_archive():
    extract_dir = 'Upload_folder'
    file_name = '_instr.zip'
    # with os.scandir(extract_dir) as scan:
    #     return next(scan, None) is None
    if len(os.listdir(extract_dir)) > 0:
        shutil.rmtree(extract_dir)
        os.mkdir(extract_dir)
    print('1')
    with zipfile.ZipFile(file_name) as zf:
        zf.extractall(extract_dir)
    return


def read_files(path_folder):
    """"переименование файлов в латиницу по одному"""
    folder = path_folder
    folder_base = os.getcwd()
    list_old = []
    for file_name in os.listdir(folder):
        count_name = ''
        count_simbol = 0
        for simbol in file_name:
            if count_simbol == 3:
                break
            if simbol.isdigit():
                count_name += simbol
                count_simbol += 1
        file_rename = 'instr' + '_' + count_name + '.docx'
        if file_rename not in os.listdir(folder):
            os.chdir(folder)
            os.rename(file_name, file_rename)
            os.chdir(folder_base)
        else:
            list_old.append(file_name)
    # print(list_old)
    return


def main():
    print(1)
    # search_string, replace_string = input_data()
    path_folder = add_path_folder()
    os.chdir(path_folder)
    read_files(path_folder)
    styles = []
    for file_name in os.listdir(path_folder):
        if file_name == 'instr_1.docx':
            print('first')
            doc_1 = docx.Document(path_folder + '\\' + file_name)
            for paragraph in doc_1.paragraphs:
                styles.append(paragraph.style)
    for file_name in os.listdir(path_folder):
        print('3', file_name)
        doc = docx.Document(path_folder + '\\' + file_name)
        doc_new = docx.Document()
        anchor = '0'
        for paragraph in doc.paragraphs:
            p_text = paragraph.text
            if p_text[:10] == 'Инструкция' or p_text[:10] == 'ИНСТРУКЦИЯ':
                table = doc_new.add_table(rows=4, cols=1)
                # получаем ячейку таблицы
                cell = table.cell(0, 0)
                # записываем в ячейку данные
                cell.text = 'Государственное бюджетное общеобразовательное учреждение'
                cell = table.cell(1, 0)
                # записываем в ячейку данные
                cell.text = 'Самарской области средняя общеобразовательная школа'
                cell = table.cell(2, 0)
                cell.text = '«Образовательный  центр имени В.Н. Татищева» с. Челно-Вершины'
                cell = table.cell(3, 0)
                cell.text = 'муниципального района Челно-Вершинский Самарской области'

                table_1 = doc_new.add_table(rows=3, cols=2)
                cell = table_1.cell(0, 0)
                cell.text = 'Согласовано'
                cell = table_1.cell(0, 1)
                cell.text = 'Утверждаю'
                cell = table_1.cell(1, 0)
                cell.text = 'Председатель профкома'
                cell = table_1.cell(2, 0)
                cell.text = '______________Сергеева Н.А.'
                cell = table_1.cell(1, 1)
                cell.text = 'Директор школы'
                cell = table_1.cell(2, 1)
                cell.text = '______________Моисеева Н.В.'
                doc_new.add_paragraph()
                table_2 = doc_new.add_table(rows=1, cols=2)
                cell = table_2.cell(0, 1)
                cell.text = 'Приказ № 142-од от 01.03.2022'
                doc_new.add_paragraph()
                print('7')
                anchor = '1'
            if anchor == '1':
                print(file_name, anchor)
                doc_new.add_paragraph(p_text)
                print('8')

        # for i in range(len(doc.paragraphs)):
        #     doc.paragraphs[i].style = styles[i]

        # doc = search_replace(doc, search_string, replace_string)
        doc_new.save('new_' + file_name)
        # получаем первую таблицу в документе
        # table = doc.tables[0]

        # читаем данные из таблицы
        # for row in table.rows:
        #     string = ''
        #     for cell in row.cells:
        #         string = string + cell.text + ' '
        #     print(string)


if __name__ == '__main__':
    print('0')
    main()


#  Каждый файл  чтобы добавить таблицу нужно переписать?
#  сначала создать новый файл
#  потом вставить в него элементы в той последовательности которой нужно
#  и сохранить под старым именем?? или новым ??
# TODO работа со стилями исходного документа или вообще нужен 1 шаблонный документ?
#  и как привести все к обному виду единообразию


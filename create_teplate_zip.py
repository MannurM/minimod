# utf-8
# скрипт по созданию шаблона файлов инструкций docx



# TODO делать ли Web interface
# TODO FLASK?

# запрос значений полей и сохранение в файл
# распаковка  папки зип архив из своих файлов
#   чтение файла из папки
#   поиск мест вставки замены
#   вставка данных
#   сохранение
# выход зип архив с изменениями

import requests
import zipfile
import os


def input_data():
    """ввод данных для изменения в файле"""
    print('введите текст поиска')
    # name_org = input()
    list_input, list_replace = [], []

    name_org = 'Государственное бюджетное общеобразовательное учреждение Самарской области средняя общеобразовательная школа «Образовательный  центр имени В.Н. Татищева» с. Челно-Вершины муниципального района Челно-Вершинский Самарской области'
    name_prof = 'Н.А. Сергеева'
    name_director = 'Н.В. Моисеева'
    name_spec = 'Специалист по охране труда_____________ М.М. Зайдуллин'
    list_input.append(name_org)
    list_input.append(name_prof)
    list_input.append(name_director)
    list_input.append(name_spec)
    print('введите текст изменения!')
    new_name_org = '{{name_org}}'
    new_name_prof = '{{name_prof}}'
    new_name_director = '{{name_director}}'
    new_name_spec = '{{name_spec}}'
    list_replace.append(new_name_org)
    list_replace.append(new_name_prof)
    list_replace.append(new_name_director)
    list_replace.append(new_name_spec)
    return list_input, list_replace


def unzip_archive():
    # TODO Сделать функцию для ввода пути к архиву и имени архива?
    # extract_dir
    # file_name
    # return extract_dir, file_name
    extract_dir = 'Upload_folder'
    file_name = '_instruct_2022.zip'
    with zipfile.ZipFile(file_name) as zf:
        zf.extractall(extract_dir)
    return


def check_upload_folder():
    """проверить файлы в папке на соответствие расширению .docx"""
    pass


def read_files():
    """"переименование файлов в латиницу по одному"""
    folder = 'Upload_folder'
    folder_base = os.getcwd()
    count = 1
    for file in os.listdir(folder):
        if count >= 1:
            file_rename = 'instr' + '_' + str(count) + '.docx'
            os.chdir(folder)
            os.rename(file, file_rename)
            os.chdir(folder_base)
        count += 1
    return


def replace_data(list_input, list_replace):
    """поиск текста на замену в файле"""

    # в цикле по файлам
    #   открыть файл , прочитать файл в буфер
    #   в цикле по условию
    #      взять первой значение
    #      найти место замены,
    #      сделать замену
    #
    # сохранить файл


def zip_template():
    path = ''  # '/home/docs-python/script/sql-script/'
    file_dir = os.listdir(path)
    with zipfile.ZipFile('test.zip', mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        for file in file_dir:
            add_file = os.path.join(path, file)
            zf.write(add_file)


def main():
    unzip_archive()
    read_files()
    list_input, list_replace = input_data()
    replace_data(list_input, list_replace)


if __name__ == '__main__':
    main()

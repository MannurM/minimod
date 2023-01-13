# utf-8
# скрипт  - форматирование файла по образцу(другому файлу) docx

# получение шаблона
# распаковка шаблона на составляющие, индексирование шаблона
# распаковка  папки зип архив из своих файлов или чтение папки с файлами
# чтение файла из папки
#   индексирование файла
#   сравнение с шаблоном
#   изменение файла по шаблону
#   сохранение файла с новый названием
# завершение работы.

import docx
import os
import os.path


def insert_template():
    # Запросить путь до шаблона
    # сохранить шаблон в переменную docx_template
    # передать шаблон дальше
    print('Введите путь до шаблона')  # скопировать сделать кнопку или функцию
    path_template = input()
    # заключить в try except
    # функция проверки файла по пути
    # функция проверки расширения шаблона
    # чтение doc или docx файла, сохранение в docx_templаte

    try:
        os.path.isfile(path_template)
        if path_template[:-4] == '.doc' or path_template[:-5] == '.docx':
            docx_template = open(path_template)
        else:
            print('Ваш файл не имеет правильного расширения .doc млм .docx')
            raise IOError
    except IOError as e:
        print(u'не удалось открыть файл')
        print('Проверьте ваш файл и повторите')
    else:
        return docx_template


def unpack_template(docx_template):
    with open(docx_template, 'r') as file_template:
        # открыть файл шаблон построчно и сделать из него шаблон для каждого типа надписи
        line_list = []
        dict_styles = {}
        for line in file_template:
            print(line)
            line_list.append(line.strip())


            # TODO можно ли так распаковать docx
            # получить стиль строки?? или лучше сначала прочитать еликом файл и конвертировать в,  только потом
            # забирать стили каждого параграфа и индексировать их.



def insert_files():
    pass


def indexing_file():
    pass


def compare_file():
    pass


def change_file():
    pass


def save_file():
    pass


def read_files():
    indexing_file()
    compare_file()
    change_file()
    save_file()


def quit_mod():
    pass


def main():
    docx_template = insert_template()
    unpack_template(docx_template)
    insert_files()
    read_files()
    quit_mod()


if __name__ == '__main__':
    main()



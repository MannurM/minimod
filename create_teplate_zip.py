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


import docx
import zipfile
import os


def input_data():
    """ввод данных для изменения в файле"""
    print('введите текст поиска')
    # name_org = input()
    d_input, d_replace = {}, {}

    d_input['name_org'] = 'Государственное бюджетное общеобразовательное учреждение Самарской области средняя общеобразовательная школа «Образовательный  центр имени В.Н. Татищева» с. Челно-Вершины муниципального района Челно-Вершинский Самарской области'
    d_input['name_prof'] = 'Н.А. Сергеева'
    d_input['name_director'] = 'Н.В. Моисеева'
    d_input['name_spec'] = 'Специалист по охране труда_____________ М.М. Зайдуллин  '

    print('введите текст изменения!')
    d_replace['name_org'] = '{{name_org}}'
    d_replace['name_prof'] = '{{name_prof}}'
    d_replace['name_director'] = '{{name_director}}'
    d_replace['name_spec'] = '{{name_spec}}'

    return d_input, d_replace


def unzip_archive():
    extract_dir = 'Upload_folder'
    file_name = '_instr_.zip'
    with zipfile.ZipFile(file_name) as zf:
        zf.extractall(extract_dir)
    return


def replace_data(d_input, d_replace):
    """поиск текста на замену в файле"""

    # в цикле по файлам
    #   открыть файл , прочитать файл в буфер
    #   в цикле по условию
    #      взять первой значение
    #      найти место замены,
    #      сделать замену
    #
    # сохранить файл
    folder = 'Upload_folder'
    folder_base = os.getcwd()
    count = 1

    print('1', os.getcwd())
    for file in os.listdir(folder):
        os.chdir(folder)
        os.getcwd()
        print(os.getcwd())
        file_path = os.path.abspath(file)
        print(file,  file_path)

        doc = docx.Document(file_path)
        all_paras = doc.paragraphs
        len(all_paras)
        print('len', len(all_paras))
        for key, value in d_input.items():
            print('key', key)
            for a_paras in all_paras:
                print('a_paras', a_paras.text)
                print(type(a_paras.text), type(value))
                print(len(a_paras.text), len(value))
                if a_paras.text == value:
                    print('d_replace', d_replace[key])
                    # a_paras.text = d_replace[key]
                    # TODO нужны два файла в одной директорри
                    #  иначе постоянно прыгать за файлами в разные директории
                    # TODO нужно очистить тектс из параграфа для возможного сравнения
        doc.save
    return




def zip_template():
    path = 'Upload_folder'  # '/home/docs-python/script/sql-script/'
    file_dir = os.listdir(path)
    with zipfile.ZipFile('_instr_new.zip.zip', mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        for file in file_dir:
            add_file = os.path.join(path, file)
            zf.write(add_file)


def main():
    # unzip_archive()
    d_input, d_replace = input_data()
    replace_data(d_input, d_replace)


if __name__ == '__main__':
    main()

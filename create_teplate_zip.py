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

    d_input['name_org'] = 'Государственное бюджетное общеобразовательное учреждение ' \
                          'Самарской области средняя общеобразовательная школа ' \
                          '«Образовательный  центр имени В.Н. Татищева» с. Челно-Вершины ' \
                          'муниципального района Челно-Вершинский Самарской области'
    d_input['name_org_var1'] = 'Государственное бюджетное общеобразовательное учреждение '
    d_input['name_org_var2'] = 'Государственное бюджетное общеобразовательное учреждение'
    d_input['name_org_var3'] = 'Государственное бюджетное общеобразовательное учреждение Самарской области'
    d_input['name_org_var4'] = 'Государственное бюджетное общеобразовательное учреждение Самарской области '

    d_input['name_prof'] = 'Н.А. Сергеева'
    d_input['name_director'] = 'Н.В. Моисеева'
    d_input['name_spec'] = 'Специалист  по охране труда_____________ М.М. Зайдуллин'
    d_input['name_pos_prof'] = "Председатель профкома"
    d_input['name_pos_direct'] = 'Директор школы'
    d_input['name_prof_var1'] = 'Н.А.Сергеева'
    d_input['name_director_var1'] = 'Н.В.Моисеева'

    print('введите текст изменения!')
    d_replace['name_org'] = '{{name_org}}'
    d_input['name_org_var1'] = '{{name_org}}'
    d_input['name_org_var2'] = '{{name_org}}'
    d_input['name_org_var3'] = '{{name_org}}'
    d_input['name_org_var4'] = '{{name_org}}'
    d_replace['name_prof'] = '{{name_prof}}'
    d_replace['name_director'] = '{{name_director}}'
    d_replace['name_spec'] = '{{name_spec}}'
    d_replace['name_pos_prof'] = '{{name_pos_prof}}'
    d_replace['name_pos_direct'] = '{{name_pos_direct}}'
    d_replace['name_prof_var1'] = '{{name_prof}}'
    d_replace['name_director_var1'] = '{{name_director}}'

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
        # print(os.getcwd())
        file_path = os.path.abspath(file)
        print(file,  file_path)

        doc = docx.Document(file_path)
        all_paras = doc.paragraphs

        for key, value in d_input.items():
            # print('key and value---', key, value)
            for paragr in all_paras:
                inline = paragr.runs
                # print(inline, len(inline))
                if value in paragr.text:
                    print('Sucsess!')
                    print('OLD', paragr.text)
                    paragr.text = d_replace[key]
                    print("NEW", paragr.text)
                # else:
                #     print(value, '-Fail!')
                # Loop added to work with runs (strings with same style)
                # for i in range(len(inline)):
                #     if 'old text' in inline[i].text:
                #         text = inline[i].text.replace('old text', 'new text')
                #         inline[i].text = text


                os.chdir(folder_base)
                # TODO нужны два файла в одной директорри
                #  иначе постоянно прыгать за файлами в разные директории
                # TODO нужно очистить тектс из параграфа для возможного сравнения

            # for table in doc.tables:
            #     for cell in table.cells:
            #         for paragr in cell.paragraphs:
            #             if value in paragr.text:
            #                 paragr.text = d_replace[key]
                        
        doc.save(file_path)
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

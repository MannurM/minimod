# utf-8
# план задание
# получить файл из папки, прочитать его, найти строку с номером приказа, заменить строку, сохранить файл


import os
import shutil
import zipfile

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt


def add_path_folder():
    """ввести путь до папки с файлами, прочитать папку, в цикле взять файл, прочитать его по параграфам"""
    print('введите путь до папки с файлами')
    path_folder = 'C:\\Users\\User\\PycharmProjects\\minimod\\Upload_folder'  # input()
    if path_folder:
        return path_folder


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


def compile_file(path_folder):
    def cell_format():
        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        cell.paragraphs[0].runs[0].font.name = 'Times New Roman'
        cell.paragraphs[0].runs[0].font.size = Pt(12)
        cell.paragraphs[0].runs[0].font.bold = True
        return

    for file_name in os.listdir(path_folder):
        print('3', file_name)
        doc = docx.Document(path_folder + '\\' + file_name)
        doc_new = docx.Document()
        anchor = '0'

        for paragraph in doc.paragraphs:
            p_text = paragraph.text
            # print(p_text[-1:])

            if p_text[-1:] == '^|':
                p_text[-1:] = '^p'
                print('Замена мягкого абзаца')
            if p_text[:10] == 'Инструкция' or p_text[:10] == 'ИНСТРУКЦИЯ':
                table = doc_new.add_table(rows=4, cols=1)
                table.alignment = WD_TABLE_ALIGNMENT.LEFT
                # получаем ячейку таблицы
                cell = table.cell(0, 0)
                # записываем в ячейку данные
                cell.text = 'Государственное бюджетное общеобразовательное учреждение'
                cell_format()
                cell = table.cell(1, 0)
                # записываем в ячейку данные
                cell.text = 'Самарской области средняя общеобразовательная школа'
                cell_format()
                cell = table.cell(2, 0)
                cell.text = '«Образовательный  центр имени В.Н. Татищева» с. Челно-Вершины'
                cell_format()
                cell = table.cell(3, 0)
                cell.text = 'муниципального района Челно-Вершинский Самарской области'
                cell_format()
                table_1 = doc_new.add_table(rows=3, cols=2)
                cell = table_1.cell(0, 0)
                cell.text = 'Согласовано'
                cell_format()
                cell = table_1.cell(0, 1)
                cell.text = 'Утверждаю'
                cell_format()
                cell = table_1.cell(1, 0)
                cell.text = 'Председатель профкома'
                cell_format()
                cell = table_1.cell(2, 0)
                cell.text = '______________Сергеева Н.А.'
                cell = table_1.cell(1, 1)
                cell.text = 'Директор школы'
                cell_format()
                cell = table_1.cell(2, 1)
                cell.text = '______________Моисеева Н.В.'
                cell_format()
                doc_new.add_paragraph()
                table_2 = doc_new.add_table(rows=1, cols=2)
                cell = table_2.cell(0, 1)
                cell.text = 'Приказ № 142-од от 01.03.2022'
                cell_format()
                doc_new.add_paragraph()
                print('7')
                anchor = '1'
            if anchor == '1':
                para = doc_new.add_paragraph()
                para.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                # para.alignment = 3  # выравниевание по ширине
                para.paragraph_format.line_spacing = 1.0
                # Как унифицировать псоледнюю строку абзаца чтобы не было больших пробелов на строке
                # заменить разрыв строки на абзац
                para_row = paragraph.runs
                for row in para_row:
                    para_row = para.add_run(row.text)
                    # Font data
                    # изменить шрифт на Times Nеw Roman,
                    para_row.style.name = row.style.name
                    # Size and name of font data
                    para_row.font.name = 'Times New Roman'
                    para_row.font.size = row.font.size
                    # Bold data
                    para_row.bold = row.bold
                    # Italic data
                    para_row.italic = row.italic
                    # Underline data
                    para_row.underline = row.underline
                    # Color data
                    para_row.font.color.rgb = row.font.color.rgb

                list_format_center = ["ИОТ", "Инс", "ИНС", 'по ', '1. ', '2. ', '3. ', '4. ', '5. ', '6. ', '7. ',
                                      '8. ', '9. ', '10. ', '11. ', '12. ',
                                      'I. ', 'II.', 'III', 'IV.', 'V. ', 'VI.', 'VII', 'IX.', 'X. ', 'XI.', 'XII',
                                      'XV.', 'XVI']
                if p_text[:3] in list_format_center:
                    para.alignment = 1  # Центрирование заголовков по центру
        doc_new.save('new_' + file_name)


def main():
    print('1')
    path_folder = add_path_folder()
    os.chdir(path_folder)
    # read_files(path_folder)
    compile_file(path_folder)


if __name__ == '__main__':
    print('0')
    main()

# TODO Нужно расписать максимальнов возможные ошибки для выявления и коррекции
# TODO Нужно сделать единый шаблон для инструкции и на основании его привести все  инструкции к единообразию!
# TODO  Удалить лишние интервалы?

def delete_paragraph(paragraph):  # Удалить пустой абзац
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None
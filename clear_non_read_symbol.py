# utf-8
import docx2txt
import os


def clear_non_read_symbol(path_folder):
    symbol_s = '\\' # '/'
    for file_name in os.listdir(path_folder):
        if file_name[-5:] != '.docx':
            continue
        base_file_name = file_name[:-5]
        new_file_name = path_folder + symbol_s + base_file_name + '.txt'  # новый текстовый файл для записи результата
        doc_new = open(new_file_name, 'w+')  # Открываем файл с дозаписью

        rezult_text = docx2txt.process(path_folder + symbol_s + file_name)  # Извлекаем текст в формате txt
        rezult_text_list = rezult_text.splitlines()  # читаем по строчно
        list_rezult = []
        for i, val in enumerate(rezult_text_list):
            if val != '':  # Удаляем пустые строки
                list_rezult.append(val)  # создаем из текста список
        for string in list_rezult:
            list_rezult_new = string.split(sep=' ')  # Делим текст каждого абзаца по пробелам на слова и группы слов
            sep_data = '\xa0'  # Строка с нечитаемыми символами для удаления из текста
            for index, s in enumerate(list_rezult_new):
                if sep_data in s:  # Поиск нечитаемого символа
                    s = list_rezult_new.pop(index)  # выделение нечитаемого символа из группы слов
                    s_list = s.split(sep=sep_data)  # разделение группы слов по нечитаемому символу
                    new_s = ' '.join(s_list)  # обратная сборка группы слов в качестве разделителя пробел
                    list_rezult_new.insert(index, new_s)  # вставка на свою позицию
            new_string = ' '.join(list_rezult_new)  # обратная сборка абзаца в качестве разделителя пробел
            doc_new.write(new_string + '\n')  # Запись в файл измененного абзаца
        doc_new.close()
    return

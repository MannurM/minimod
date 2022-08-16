# utf-8
import shutil
import zipfile
import os


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


def read_files():
    """"переименование файлов в латиницу по одному"""
    folder = 'Upload_folder'
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
    print(list_old)
    return list_old


# def delete_old_file(list_old):
#     folder = 'Upload_folder'
#     print(list_old)
#     for the_file in os.listdir(folder):
#         if the_file in list_old:
#             os.remove(the_file)
#             print('DELETE!')
#         else:
#             print(the_file)
#     return


def zip_template():
    path = 'Upload_folder'  # '/home/docs-python/script/sql-script/'
    file_dir = os.listdir(path)
    with zipfile.ZipFile('_instr_new.zip.zip', mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        for file in file_dir:
            add_file = os.path.join(path, file)
            zf.write(add_file)


def main():
    unzip_archive()
    list_old = read_files()
    # delete_old_file(list_old)


if __name__ == '__main__':
    main()

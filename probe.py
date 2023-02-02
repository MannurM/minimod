import os

path_folder = 'C:\\Users\\User\\PycharmProjects\\minimod\\Upload_folder'  # input()
# path_folder = '/home/mannur/PycharmProjects/minimod/Upload_folder'

# symbol = '/'  #
symbol = '\\'

for file_name in os.listdir(path_folder):
    if file_name[-5:] == '.docx':
        continue
    list_new = []
    with open(path_folder + symbol + file_name, 'r') as f:
        list_f = f.readlines()
        # print(type(list_f))
        for lf in list_f:
            if lf[-1] != '.\n':
                new = (lf[:-1] + '\n')
                list_new.append(new)
        print(list_new)
    with open(path_folder + symbol + file_name, 'w') as f:
        for new in list_new:
            f.write(new)

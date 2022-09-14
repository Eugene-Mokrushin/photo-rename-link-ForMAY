import os
import shutil
import openpyxl
import re

main_folder = os.getcwd()

list_of_required_folders = ['Фото основное', 'E-COM РУССКИЙ', 'Фото вложения', '4']


def get_list(folders_only=False):
    result = os.listdir()
    if folders_only:
        result = [f for f in result if os.path.isdir(f)]
    return result


all_sku_list = get_list(True)


def get_xlsx_file():
    global xlsx_file
    arr = os.listdir(main_folder)
    for everything in arr:
        if '.xlsx' in everything:
            xlsx_file = everything
    return xlsx_file



EAN_UPC = openpyxl.load_workbook(get_xlsx_file())
current_sheet = EAN_UPC['Sheet1']

list_of_all_EAN = []
for i in range(1, 150):
    list_of_all_EAN.append(current_sheet[f'A{i}'].value)
list_of_all_EAN = list(filter(None, list_of_all_EAN))

counter = 1
why = 0
count_SKU = 0
for sku in all_sku_list:
    if count_SKU >= len(list_of_all_EAN):
        break
    current_EAN = list_of_all_EAN[count_SKU]
    for folder in list_of_required_folders:
        for address, dirs, files in os.walk(f'{main_folder}/{sku}'):
            if os.path.basename(os.path.normpath(address)) == folder:
                for file in files:
                    base_name, ext = os.path.splitext(file)
                    old_name = os.path.join(address, file)
                    new_name = os.path.join(address, f'{current_EAN}_{counter}{ext}')
                    os.rename(old_name, new_name)
                    counter += 1
    counter = 1
    count_SKU += 1


# copy and delete
def remove_folder(path):
    # check if folder exists
    if os.path.exists(path):
        # remove if exists
        shutil.rmtree(path)


def copy_file(name, new_name):
    if os.path.isdir(name):
        try:
            shutil.copytree(name, new_name)
        except FileExistsError:
            print('Такая папка уже есть')
    else:
        shutil.copy(name, new_name)


for sku in all_sku_list:
    for (address, dirs, files) in os.walk(f"{main_folder}/{sku}"):
        if os.path.basename(os.path.normpath(address)) in list_of_required_folders:
            for file in files:
                full_photo_path = os.path.join(address, file)
                if '.jpeg' in full_photo_path:
                    copy_file(full_photo_path, f"{main_folder}/{sku}")

for sku in all_sku_list:
    for (address, dirs, files) in os.walk(f"{main_folder}/{sku}"):
        for dir in dirs:
            alt_full = os.path.join(address, dir)
            if os.path.isdir(alt_full):
                remove_folder(alt_full)

file_check = re.compile('_0')

for EAN in list_of_all_EAN:
    for sku in all_sku_list:
        for address, dirs, files in os.walk(f'{main_folder}/{sku}'):
            for file in files:
                base_name, ext = os.path.splitext(file)
                if file_check.search(base_name) != None:
                    old_name = os.path.join(address, file)
                    base_name = base_name[:-2]
                    new_name = os.path.join(address, f'{base_name}{ext}')
                    os.rename(old_name, new_name)

counter_EAN = 0
for sku in all_sku_list:
    os.rename(f'{main_folder}/{sku}', f'{main_folder}/{list_of_all_EAN[counter_EAN]}')
    counter_EAN += 1

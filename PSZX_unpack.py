import zipfile
import os
import shutil
import glob

#Makro do rozpakowania plików PSZX oraz ich bibliotek
path='D:\sys+root_check\Kernlinie\\'
file_pattern=f'{path}*.pszx'
files=glob.glob(file_pattern)
print(files)

i=0
for file in files:
    file_basename=os.path.splitext(os.path.basename(file))[0]
    folder_name=file_basename
    if not os.path.exists(f'{path}{folder_name}'):
        os.mkdir(f'{path}{folder_name}')
    with zipfile.ZipFile(file,'r') as zip_ref:
        zip_ref.extractall(f'{path}{folder_name}')
    i = i + 1
    print(f'{i} {file} PSZX rozpakowany')
    with zipfile.ZipFile(f'{path}{folder_name}\Library.zip','r') as zip_ref:
        zip_ref.extractall(f'{path}{folder_name}')
    print(f'{i} {file} Biblioteka rozpakowana')

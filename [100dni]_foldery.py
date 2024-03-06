import os
import shutil
import glob

folder='source'

file_patern=f'{folder}\*.*'
files=glob.glob(file_patern)
extensions=set()
for file in files:
    file_basename,file_extension=os.path.splitext(file)
    extensions.add(file_extension)
print(extensions)

for ext in extensions:
    nazwa_folderu=ext.replace('.','')
    if not os.path.exists(f'{folder}\{nazwa_folderu}'):
        os.mkdir(f'{folder}\{nazwa_folderu}')
    file_patern=f'{folder}\*{ext}'
    files=glob.glob(file_patern)
    for file in files:
        file_name=file.replace(f'{folder}\\','')
        print(file_name)
        shutil.move(file,f'{folder}\{nazwa_folderu}\{file_name}')

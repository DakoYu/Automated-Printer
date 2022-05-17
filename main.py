'''
Python Scripts to automate printing files
'''

import os, glob, shutil
import pandas as pd
from datetime import date

path = os.path.join('D:\\', 'Python')

# Making sure the path do exist

year = date.today().year
month = date.today().month

dir = os.path.join('D:\\', 'Hinga', str(year), str(month))

# Create the path
if not os.path.exists(dir):
    os.makedirs(dir)

files = glob.glob(f'{dir}/*.xlsx')
for file in files:
    df = pd.read_excel(file).to_string()
    name = 'misc'
    if '哈哈' in df:
        print('Yes')
    

    os.startfile(file, 'print')

    move_dir = os.path.join(dir, name)
    print(move_dir)
    if not os.path.exists(move_dir):
        os.makedirs(move_dir)

    # shutil.move(file, move_dir)
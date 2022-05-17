'''
Python Scripts to automate printing files
'''

import os, glob, shutil
import pandas as pd
from time import sleep
from datetime import date

path = os.path.join('D:\\', 'Python')

# Making sure the path do exist

year = date.today().year
month = date.today().month

dir = os.path.join('D:\\', '兴雅', str(year), str(month))

# Create the path
if not os.path.exists(dir):
    os.makedirs(dir)

test_dir = os.path.join('D:\\', '兴雅')

files = glob.glob(os.path.join(test_dir, '**/*'), recursive=True)
for file in files:
    if os.path.isfile(file):
        df = pd.read_excel(file).to_string()
        name = 'misc'
        if '富 怡 雅' in df:
            name = '富怡雅'
        elif '锦昌' in df or '锦骏' in df:
            name = '锦昌'
        elif '新卓' in df:
            name = '新卓'
        


        # os.startfile(file, 'print')

        move_dir = os.path.join(dir, name)
        print(move_dir)
        if not os.path.exists(move_dir):
            os.makedirs(move_dir)
        sleep(5)
        shutil.move(file, move_dir)
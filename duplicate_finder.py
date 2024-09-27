import os
import pathlib
import openpyxl
import pandas as pd

file_path = pathlib.Path().resolve()
main_list = []
inside_list = os.listdir(file_path)

for i,x in enumerate(inside_list):
    if not x.endswith('.py'):
        playlist_path = str(file_path)+'\\'+ x
        play_list = os.listdir(playlist_path)
        main_list.extend(play_list)

res = set([duplicate for duplicate in main_list if main_list.count(duplicate)>1])
print('There are', len(res) , 'duplicates')

df = pd.DataFrame(res, columns=['File Name'])
excel_file = 'duplicates.xlsx'
df.to_excel(excel_file, index=False)
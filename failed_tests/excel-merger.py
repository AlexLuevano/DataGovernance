import pandas as pd
import xlsxwriter
from tkinter import *
from tkinter import filedialog

window = Tk()
files = filedialog.askopenfilenames(parent = window, title = 'Escoge archivos')
files_str = str(files)
files_list = list(files)
window.destroy()
window.mainloop()


df = pd.DataFrame()
writer = pd.ExcelWriter('excel_merged.xlsx' ,engine='xlsxwriter')
print(files_list)

for f in files_list:
    if f[-4:] == '.csv':
        data = pd.read_csv(f, index_col=False, header=0, na_filter = False)
        df = df.append(data)
    else:
        data = pd.read_excel(f,sheet_name='Entities', index_col=False, header=1, na_filter = False)
        df = df.append(data)

print('All files read and merged')
df.to_excel(writer,sheet_name='Entities',index = False, header = True)
writer.save()
print('File saved')
#writer.close()
print('File closed')
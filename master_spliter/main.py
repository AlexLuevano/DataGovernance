import pandas as pd
import xlsxwriter
from tkinter import *
from tkinter import filedialog
import os
import sys


window = Tk()

def browse_files():
    global filename
    filename = filedialog.askopenfilename(initialdir = "/Downloads", title = "Selecciona un archivo:", filetypes = (("all files","*.*"),(".xlsx Files","*.xlsx")))
    if filename:
            l1 = Label(window, text = "File path: " + filename).pack()
    else:
            print('No seleccionaste ningún archivo.')
    window.destroy()


def main():
    #window = Tk()
    window.title("Master per APREF")
    window.geometry("200x100")
    label_file_explorer = Label(window, text = "Master per APREF", width = 100, height = 4).pack()
    button_explore = Button(window, text = "Buscar archivo", command = browse_files).pack()
    window.mainloop()

    df = pd.read_excel(f'{filename}', index_col = False,na_filter= False)
    #writer = pd.ExcelWriter('excel_merged.xlsx' ,engine='xlsxwriter')
    aprefs = list(df['Ap Ref'].unique())

    if os.path.exists('C:\\temporary'):
        for num in aprefs:
                temp_df = df[df['Ap Ref'] == num]
                temp_df.to_excel(f'C:\\temporary\\{num}.xlsx',index = False, header = True)
        print(r'File succesfully split. Please review files at C:\\temporary. Finishing execution.')
        exit()
    else:
        print('C:\\temporary directory does not exist. Do you want to create it? Y/N?')
        answer = input()
        if answer == 'Y' or 'y':
            os.mkdir(r'C:\\temporary')
            for num in aprefs:
                temp_df = df[df['Ap Ref'] == num]
                temp_df.to_excel(f'C:\\temporary\\{num}.xlsx',index = False, header = True)
            print(r'File succesfully split. Please review files at C:\\temporary')
            exit()
        elif answer == 'N' or 'n':
            print('Ok. Finishing execution')
            exit()
        else:
            print('Invalid answer. Finishing execution')
            exit()

if __name__ == '__main__':
    main()


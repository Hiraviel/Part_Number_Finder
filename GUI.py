import os
from os.path import join

from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
from xlrd import open_workbook

# = Main Window =

root = Tk()
root.geometry('800x499')
root.resizable(height=FALSE, width=FALSE)
root.configure(bg='#c6cad1')
root.title('Part Number Finder v1')
root.option_add('*Dialog.msg.font', 'Arial 12')


# = Functions =

def clear_files(directory):

    for (dirname, dirs, files) in os.walk(directory):
        for filename in files:
            if filename.startswith('~'):
                thefile = os.path.join(dirname, filename)
                os.remove(thefile)
    for (dirname, dirs, files) in os.walk(directory):
        for filename in files:
            if filename.startswith('$'):
                thefile = os.path.join(dirname, filename)
                os.remove(thefile)


def search_folder():
    a = StringVar()

    status = folder_location.get()
    root.filename = filedialog.askdirectory()

    if root.filename != status:
        folder_entry.delete(0, END)
        folder_location.set(root.filename)
    else:
        folder_location.set(status)


def search_np():
    Xtensions = ('.xlsx', '.xls')
    i = 0

    directory = folder_location.get()
    Part_Number = part_number.get()
    file_save = part_number.get().upper() + '.txt'

    result_locations.delete('1.0', END)
    clear_files(directory)

    np_txt = open(file_save, "w")
    for (dirname, dirs, files) in os.walk(directory):
        for filename in files:
            if filename.endswith(Xtensions):
                Part_Number_A = os.path.join(dirname, filename)
                book = open_workbook(Part_Number_A)
                for sheet in book.sheets():
                    for rowidx in range(sheet.nrows):
                        row = sheet.row(rowidx)
                        for colidx, cell in enumerate(row):
                            if cell.value == Part_Number.upper():
                                i += 1
                                np_txt.write(Part_Number_A + '\n')
                                result_locations.insert('1.0', '{} > {} \n'.format(i, Part_Number_A))

    np_txt.close()
    tkinter.messagebox.showinfo('Resultados', 'Se encontaron {} elementos'.format(i))

# = Variables =

part_number = StringVar()
folder_location = StringVar()
color = '#c64123'

# = Frames =

f1 = Frame(root, width=800, height=20, bg=color)
f1.pack(side=TOP)

f2 = Frame(root, width=800, height=20, bg=color)
f2.pack(side=TOP)

f3 = Frame(root, width=800, height=20, bg=color)
f3.pack(side=TOP)

f4 = Frame(root, width=800, height=40, bg='black')
f4.pack(side=TOP)

f5 = Frame(root, width=800, height=20, bg='#1959c1')
f5.pack(side=TOP)


# = Labels =

part_number_label = Label(f1, text='Introduce el número de parte:', width=22, height=1, bg=color, fg='white', font=("Arial", 12))
part_number_label.pack(side=LEFT, padx=3, pady=3)

part_number_entry = Entry(f1, text='Número de parte', textvariable=part_number, width=100)
part_number_entry.pack(side=LEFT, padx=3, pady=3)


folder_label = Label(f2, text='Folder donde buscar:', width=16, height=1, bg=color, fg='white', font=("Arial", 12))
folder_label.pack(side=LEFT, padx=3, pady=3)

folder_entry = Entry(f2, text='Número de parte', textvariable=folder_location, width=89)
folder_entry.pack(side=LEFT, padx=3, pady=3)

result_locations = Text(f4, font=("Arial", 10), bg='black', fg='white', width=114)
result_locations.pack(side=TOP)

powered_by_label = Label(f5, text='Powered by Python' + ' '*150, bg='#1959c1', fg='#eff0f2', width=80)
powered_by_label.pack(side=LEFT)

credits_label = Label(f5, text=' '*25 + 'Developed by Rafael Escoto', bg='#1959c1', fg='#eff0f2', width=33)
credits_label.pack(side=RIGHT)

# = Buttons =

folder_button = Button(f2, text='Folder', width=15, command=lambda: search_folder())
folder_button.pack(side=RIGHT, padx=3, pady=3)


search_button = Button(f3, text='Buscar', width=115, command=lambda: search_np())
search_button.pack(side=LEFT, padx=3, pady=3)

root.mainloop()
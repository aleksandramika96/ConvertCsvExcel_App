# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
#libraries
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import csv 

#frame
root= tk.Tk()

canvas1 = tk.Canvas(root, width = 500, height = 400, bg = 'lightgrey', relief = 'sunken') #relief - specifies the type of border
canvas1.pack()

label1 = tk.Label(root, text='File Conversion Tool!', bg = 'lightgrey') #lightsteelblue2
label1.config(font=('Calibri Light', 20)) #helvetica
canvas1.create_window(250, 60, window=label1)

label2 = tk.Label(root, text='Delimiter', bg = 'lightgrey') #lightsteelblue2
label2.config(font=('Calibri Light', 12)) #helvetica
canvas1.create_window(250, 160, window=label2)

entry1 = tk.Entry (root) #delimiter
canvas1.create_window(250, 180, window=entry1) 


def getCSV ():
    global read_file
    delimiter = str(entry1.get())
    
    import_file_path = filedialog.askopenfilename()
    #read_file = pd.read_csv (import_file_path)
    
    with open(import_file_path, 'r', encoding = 'utf8') as file:
        csv_read = csv.reader(file, delimiter = delimiter, quoting=csv.QUOTE_NONE)    
        read_file = pd.DataFrame(csv_read)   
    
browseButton_CSV = tk.Button(text="      Import CSV File     ", command=getCSV, bg='green', fg='white', font=('Calibri Light', 12, 'bold'))
canvas1.create_window(250, 230, window=browseButton_CSV)


def convertToExcel ():
    global read_file
    
    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    read_file.to_excel (export_file_path, index = None, header=False)
    return tk.messagebox.showinfo(title = 'Status', message = 'File was saved.')

saveAsButton_Excel = tk.Button(text='Convert CSV to Excel', command=convertToExcel, bg='green', fg='white', font=('Calibri Light', 12, 'bold'))
canvas1.create_window(250, 280, window=saveAsButton_Excel)

def exitApplication():
    MsgBox = tk.messagebox.askquestion ('Exit Application','Are you sure you want to exit the application',icon = 'warning')
    if MsgBox == 'yes':
       root.destroy()
     
exitButton = tk.Button (root, text='       Exit Application     ',command=exitApplication, bg='brown', fg='white', font=('Calibri Light', 12, 'bold'))
canvas1.create_window(250, 330, window=exitButton)

root.mainloop()
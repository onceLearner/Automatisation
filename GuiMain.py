# !/usr/bin/python3
from tkinter import *
from tkinter.filedialog import askopenfile
from MainScript import Guru

parent = Tk()
parent.geometry("800x400")

GoodMorning=Label(parent,text="     Hello , click bellow  to load your script\n \n \n ").grid(row=0,column=0)
btn = Button(parent, text ='load excel sheet ', command = lambda:open_file()).grid(row=2,column=0)

filename="no file "

def open_file():
    file = askopenfile(mode ='r', filetypes =[('Python Files', '*.xlsx')])
    filename=file.name
    print(file.name)
    name = Label(parent, text="your file directory : \n\n"+filename) .grid(row=5, column=0)
    btnRun = Button(parent, text='run Script', command=lambda: Guru(filename)).grid(row=9, column=1)






parent.mainloop()
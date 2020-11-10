# !/usr/bin/python3
from tkinter import *
from tkinter.filedialog import askopenfile
from MainScript import Guru
import openpyxl












parent = Tk()
parent.geometry("1150x600")


GoodMorning=Label(parent,text="     Hello , click bellow  to load your script ").grid(row=0,column=3)
btnLoad = Button(parent, text ='load excel sheet ', command = lambda:open_file(),foreground="white",background="gray").grid(row=4,column=3)


for elt in [7,9,11]: # spacing the elements
    parent.grid_rowconfigure(elt, minsize=15)
parent.grid_rowconfigure(5, minsize=40)
################################ Select bleed
isBleed= [
"Bleed",
"NoBleed",
] #etc

bleedVar = StringVar(parent)
bleedVar.set(isBleed[0]) # default value


selectBleed = OptionMenu(parent, bleedVar, *isBleed).grid(row=6,column=3)

########################################### Select adult
isAdult= [
"Not Only Adult",
"Yes Only Adult",
] #etc

adultVar = StringVar(parent)
adultVar.set(isAdult[0]) # default value


selectAdult = OptionMenu(parent, adultVar, *isAdult).grid(row=8,column=3)





######################################### Select  color
isColor= [
"NoColor",
"Color",
] #etc

colorVar = StringVar(parent)
colorVar.set(isColor[0]) # default value


selectColor = OptionMenu(parent, colorVar, *isColor).grid(row=10,column=3)




######################################## Select size


isSize= [

"6x9",
"8.5x11"
]#etc

sizeVar = StringVar(parent)
sizeVar.set(isSize[0]) # default value


selectSize = OptionMenu(parent, sizeVar, *isSize).grid(row=12,column=3)


###############################################


parent.grid_columnconfigure(0, minsize=80)
parent.grid_rowconfigure(1, minsize=70)
filename="no file "




def open_file():
    file = askopenfile(mode ='r', filetypes =[('Python Files', '*.xlsx')])
    filename=file.name

    workbook = openpyxl.load_workbook(filename, data_only=True)
    sheet = workbook.active


    Label(parent, text="start cell",width=10).grid(row=4, column=6)
    Label(parent, text="stop cell(included)").grid(row=6, column=6)


    start_index = Entry(parent)
    start_index.insert(0,2)
    start_index.grid(row=4, column=7)


    lbl1=Label(parent, text=sheet["A"+start_index.get()].value,font=("Courier", 7))
    lbl1.grid(row=5, column=7)

    def start_book_evt(event):
        lbl1.config(text=sheet["A"+start_index.get()].value)


    start_index.bind("<Return>",start_book_evt)
    stop_index = Entry(parent)
    stop_index.grid(row=6, column=7)
    cell=int(start_index.get())
    maxRows=1
    while sheet["A" + str(cell)].value:
        cell+=1
        maxRows+=1
    stop_index.insert(0,maxRows)
    lbl2=Label(parent, text=sheet["A" + str(maxRows)].value,font=("Courier", 7))
    lbl2.grid(row=7, column=7)

    def stop_book_evt(event):
        lbl2.config(text=sheet["A" + stop_index.get()].value)

    stop_index.bind("<Return>", stop_book_evt)


    fileNameLbl = Label(parent, text="your file directory : ").grid(row=10, column=7,sticky=W)
    name = Label(parent, text=filename,font=("Sans-Serif", 8)) .grid(row=11, column=7)
    btnRun = Button(parent, text='run Script', command=lambda: Guru(filename,int(start_index.get()),int(stop_index.get()),bleedVar.get(),adultVar.get(),colorVar.get(),sizeVar.get())).grid(row=14, column=7)




parent.mainloop()
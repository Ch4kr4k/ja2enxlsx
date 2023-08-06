from tkinter import *
import msoffcrypto
from jaenfunc.jfunc import unlock
from jaenfunc.jfunc import translate_to_english
from tkinter import filedialog
import pandas as pd
import io
import os

file_name= ''
converted_file_name='decrypted_final.xlsx'
temp = io.BytesIO()
password = 0


def fileselector():
    global file_name
    file_name = filedialog.askopenfilename()


def conv():
    global password
    password = entry.get()
    if entry.get() == "password":
        input = file_name
    else:
        unlock(file_name,password)
        input = 'out.xlsx'
    df = pd.ExcelFile(input)
    sheet_names = df.sheet_names
    writer = pd.ExcelWriter("translated.xlsx")
    for i in sheet_names:
        sheetname = i
        data_frame = pd.read_excel(input,sheetname)
        if data_frame is not None:
            translated_df = data_frame.applymap(translate_to_english)
        translated_df.to_excel(writer, sheet_name=sheetname)
    writer.close()
    if input == 'out.xlsx':
        os.remove('out.xlsx')



def sub():
    global password
    password = entry.get()
    entry.config(state=DISABLED)
    return password


def dell():
    entry.delete(0,END)


def backspace():
    entry.delete(len(entry.get()) - 1, END)

    
window = Tk()
window.title("ja2en")
label = Label(window,
              text="Jap 2 Eng",
              font=('Arial', 16),
              relief=SUNKEN,
              padx=20,
              pady=20)

open_file = Button(text="Chose a file",
                   font=('Arial', 14),
                   command=fileselector
                   )
passtext = StringVar()
passtext.set("password")

entry = Entry(window,
              font=("Arial", 14),
              textvariable=passtext,
              )
label_leave=Label(window,
            text="just select file and press conv if there is no password",
            font=("Arial",12))

conv_button = Button(text="Convert now",
                   font=('Arial', 14),
                   command=conv
                   )

label.pack()
open_file.pack()
entry.pack()
label_leave.pack()
conv_button.pack()
window.geometry("500x500")
window.mainloop()

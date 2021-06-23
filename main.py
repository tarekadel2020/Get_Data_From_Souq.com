from tkinter import messagebox , filedialog
import tkinter as tk
from openpyxl import load_workbook
from bs4 import BeautifulSoup
import requests
from tkinter import *
import re




# https://www.youtube.com/watch?v=7YS6YDQKFh0

def souq():
    try:
        wb = load_workbook(Enter_Path.get())
    except:
        messagebox.showinfo("Erorr", "Please Enter Path")

    ws = wb.active

    try:
        zz = int(Enter_Cell.get())+1
        z(zz,ws , wb)
    except:
        messagebox.showinfo("Erorr", "Please Enter End Cell")

#print(ws['A20'].value)


def z(x , ws , wb) :
    for cal in range(2,x):
        if ws["A"+ str(cal)].value == None:
            print("error")
        else:
            print(ws["A"+ str(cal)].value)
            zz = ws["A" + str(cal)].value
            r = requests.get(zz)
            soup = BeautifulSoup(r.content, "html.parser")
            for souq_title in soup.findAll('head') :
                print(souq_title.text.strip())
                ws["B" + str(cal)].value = souq_title.text.strip()
                try:
                    wb.save(Enter_Path.get())
                except:
                    messagebox.showinfo("Erorr", "Please Close The Excel")

            for souq_name in soup.findAll('h1'):
                print(souq_name.text)
                ws["C" + str(cal)].value = souq_name.text
                try:
                    wb.save(Enter_Path.get())
                except:
                    messagebox.showinfo("Erorr", "Please Close The Excel")

            for souq_price in soup.findAll('div', class_='columns large-8 medium-8 small-7') :
                print(souq_price.text.strip())
                ws["D" + str(cal)].value = souq_price.text.strip()
                try:
                    wb.save(Enter_Path.get())
                except:
                    messagebox.showinfo("Erorr", "Please Close The Excel")

            for souq_qte in soup.findAll('div', class_='unit-labels'):
                print(souq_qte.text.strip())
                ws["E" + str(cal)].value = souq_qte.text.strip()
                try:
                    wb.save(Enter_Path.get())
                except:
                    messagebox.showinfo("Erorr", "Please Close The Excel")

            for url_image in soup.findAll('img', attrs={'src':re.compile('jpg')}):
                print(url_image.get('src'))
                ws["F" + str(cal)].value = url_image.get('src')
                try:
                    wb.save(Enter_Path.get())
                    count['text'] = str(x - 2)
                except:
                    messagebox.showinfo("Erorr", "Please Close The Excel")
    messagebox.showinfo("Done", "Done")







def select():
    root = tk.Tk()
    root.withdraw()
    file = str(filedialog.askopenfilename())
    Enter_Path.insert(0,file)

x = ""
top =Tk()
top.title("Souq.com")
top.minsize(500,200)

Path = Label(text="Path")
Path.pack()

Enter_Path = Entry(top , width=70)
Enter_Path.pack()

but_1 = Button(text="Select" , command=select)
but_1.pack()



Num = Label(text="Cell")
Num.pack()

Enter_Cell = Entry(top , width=10)
Enter_Cell.pack()

but = Button(text="RUN" , command=souq)
but.pack()

count = Label()
count.pack()

top.mainloop()

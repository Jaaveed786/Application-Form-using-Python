from openpyxl import Workbook,load_workbook

import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import pandas as pd

def get_student():
    sub_root=tk.Tk()
    sub_root.title('Student Details')
    #sub_root.iconbitmap('resource\foldername') is used for another folder
    sub_root.geometry('900x500')
    sub_root.config(bg='grey')
    sub_root.resizable(0,0)

    frm1=LabelFrame(sub_root,text='Search',font=('',13),bg='lightblue',fg='black')
    frm2=LabelFrame(sub_root, text='Student details', font=('', 13), bg='lightblue', fg='black')

    l1=Label(frm1,text='Name of the student : ',font=('',12),bg='lightblue')
    t1=Entry(frm1,font=('',12))
    l2 = Label(frm2, text='Name of the student : ', font=('', 12), bg='lightblue')
    l3 = Label(frm2, text='Father name : ', font=('', 12), bg='lightblue')
    l4 = Label(frm2, text='Course name : ', font=('', 12), bg='lightblue')
    l5 = Label(frm2, text='Fees : ', font=('', 12), bg='lightblue')
    l6 = Label(frm2, text='Total fees : ', font=('', 12), bg='lightblue')
    l7 = Label(frm2, text='Balance : ', font=('', 12), bg='lightblue')
    l8 = Label(frm2, text='Started date : ', font=('', 12), bg='lightblue')
    l9 = Label(frm2, text='Ended date : ', font=('', 12), bg='lightblue')
    l10 = Label(frm2, text='Certificate issued : ', font=('', 12), bg='lightblue')

    t2=Entry(frm2,font=('',12))
    t3 = Entry(frm2, font=('', 12))
    t4 = Entry(frm2, font=('', 12))
    t5 = Entry(frm2, font=('', 12))
    t6 = Entry(frm2, font=('', 12))
    t7 = Entry(frm2, font=('', 12))
    t8 = Entry(frm2, font=('', 12))
    t9 = Entry(frm2, font=('', 12))
    t10 = Entry(frm2, font=('', 12))

    def clr():
        #t1.delete(0,END)
        t2.delete(0,END)
        t3.delete(0, END)
        t4.delete(0, END)

    def get():
        f=t1.get()
        data=pd.read_excel('data.xlsx',sheet_name=['Sheet2','Sheet3'])
        sdata=data['Sheet3']

        if f in sdata.values:
            x=sdata.loc[(sdata['Name of the student']==f)]  #to get only one record of given student
            t2.insert(0,x.iloc[0,0]+' '+x.iloc[0,1])
            t3.insert(0, x.iloc[0, 2])
            t4.insert(0, x.iloc[0, 3])
            t5.insert(0, x.iloc[0, 4])
            #t6.insert(0, x.iloc[0, 5])
            #t7.insert(0, x.iloc[0, 6])
            #t8.insert(0, x.iloc[0, 7])
            #t9.insert(0, x.iloc[0, 9])
            t10.insert(0, x.iloc[0,9])

        else:
            messagebox.showerror('Error','Please enter name correctly')
            sub_root.lift()

    bt1=Button(frm1,text='Get',command=get)
    frm1.pack(padx=10,pady=10)
    frm2.pack(padx=10,pady=10)

    l1.grid(row=0,column=0,padx=10,pady=10)
    t1.grid(row=0,column=1,padx=10,pady=10)
    bt1.grid(row=0,column=2,padx=10,pady=10)

    l2.grid(row=0,column=0,padx=10,pady=10)
    t2.grid(row=0,column=1,padx=10,pady=10)
    l3.grid(row=0,column=2,padx=10,pady=10)
    t3.grid(row=0,column=3,padx=10,pady=10)

    l4.grid(row=1,column=0,padx=10,pady=10)
    t4.grid(row=1,column=1,padx=10,pady=10)
    l5.grid(row=1,column=2,padx=10,pady=10)
    t5.grid(row=1,column=3,padx=10,pady=10)

    l6.grid(row=2,column=0,padx=10,pady=10)
    t6.grid(row=2,column=1,padx=10,pady=10)
    l7.grid(row=2,column=2,padx=10,pady=10)
    t7.grid(row=2,column=3,padx=10,pady=10)

    l8.grid(row=3,column=0,padx=10,pady=10)
    t8.grid(row=3,column=1,padx=10,pady=10)
    l9.grid(row=3,column=2,padx=10,pady=10)
    t9.grid(row=3,column=3,padx=10,pady=10)

    l10.grid(row=4,column=1,padx=10,pady=10)
    t10.grid(row=4,column=2,padx=10,pady=10)

    sub_root.mainloop()

#get_student()
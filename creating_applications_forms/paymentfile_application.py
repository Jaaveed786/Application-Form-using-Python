import tkinter
from openpyxl import Workbook,load_workbook
from tkinter import *
from tkinter import messagebox
import pandas as pd
import datetime

date=datetime.datetime.now()
dt=date.strftime('%d/%m/%y')
formatStr='%d/%m/%y'

pd.set_option('display.max_columns',500)   #500 to 1000 is used for space in b/w columns
pd.set_option('display.max_rows',500)
pd.set_option('display.width',500)

def pay_fees():
    sub_root=tkinter.Tk()
    sub_root.title('Fees payments')
    sub_root.geometry('700x400+800+100')

    l1=Label(sub_root,text='Name of the student : ')
    t1=Entry(sub_root)

    t2=Text(sub_root,height=50,width=200,font=('',13))
    l2=Label(sub_root,text='Name of the student:',font=('',10))
    l3=Label(sub_root, text='Father name :', font=('', 10))
    l4=Label(sub_root, text='Fees paid :', font=('', 10))
    l5 = Label(sub_root, text='Paid date :', font=('', 10))
    t3=Entry(sub_root,font=('',10))
    t4=Entry(sub_root,font=('', 10))
    t5=Entry(sub_root,font=('', 10))
    t6=Entry(sub_root,font=('', 10))
    t7=Entry(sub_root,font=('', 10))

    def get():
        f=t1.get()
        data=pd.read_excel('data.xlsx')
        if f in data.values:
            x=data.loc[(data['Name of the student']==f)]    #to get only one record of given student
            t4.insert(0,x.iloc[0,0])
            t5.insert(0,x.iloc[0,1])
            t7.insert(0,dt)
        else:
            messagebox.showerror('Error','Please enter correct name')

    def save():
        wb=load_workbook('data.xlsx')
        ws=wb['Sheet2']
        sn_var=t4.get()
        fn_var=t5.get()
        fp_var=int(t6.get())
        dt_var=t7.get()

        ws.append([sn_var,fn_var,fp_var,dt_var])
        wb.save('data.xlsx')
        messagebox.showinfo('Message','Data successfully saved')

    bt1=Button(sub_root,text='Get',command=get)
    bt2 = Button(sub_root, text='Save', command=save)

    l1.place(x=150,y=10)
    t1.place(x=280,y=10)
    bt1.place(x=450,y=10)
    l2.place(x=50, y=50)
    t4.place(x=160, y=50)
    l3.place(x=50, y=80)
    t5.place(x=160, y=80)
    l4.place(x=50,y=110)
    t6.place(x=160,y=110)
    l5.place(x=50,y=140)
    t7.place(x=160,y=140)
    bt2.place(x=160,y=180)

    sub_root.mainloop()



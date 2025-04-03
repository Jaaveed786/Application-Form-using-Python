from openpyxl import Workbook,load_workbook

import tkinter as tk
from tkinter import *
from tkinter import ttk
import pandas as pd
import paymentfile_application
import datetime
from tkinter import messagebox

import reading_mulptisheets_form

date=datetime.datetime.now()
dt=date.strftime('%d/%m/%y')
formatStr='%d/%m/%y'

pd.set_option('display.max_columns',500)
pd.set_option('display.max_rows',500)
pd.set_option('display.width',500)

root=tk.Tk()
root.title('  APPLICATION FORM  ')
root.geometry('1000x800')
course=['Ms Office',
        'Tally prime',
        'C Language',
        'Python',
        'Adv Python',
        'Java',
        'Adv Java',
        'DTP',
        'Photoshop',
        'Web design']

cframe=LabelFrame(root,text='Course details',fg='blue',bg='lightblue',font=('',12,'bold'))
cframe.pack()

lb1=Label(root,text='APPLICATION FORM',bg='green',fg='white',font=('',16,),bd=5,relief=GROOVE)
lb1.pack(fill=X)

lb2=Label(cframe,text='Select Course :',bg='lightblue',font=('',12))
lb2.grid(row=0,column=0)

cn=ttk.Combobox(cframe,width=20,values=course,font=('',12))
cn.grid(row=0,column=1)

lb3=Label(cframe,text='Course fees :',bg='lightblue',font=('',12))
lb3.grid(row=0,column=2)

cf=Entry(cframe,bd=5,font=('',12))
cf.grid(row=0,column=3,padx=5,pady=10)

sframe=LabelFrame(root,text='Student Name',fg='blue',bg='lightblue',font=('',12,'bold'))
sframe.pack()

def get_fees():
    global cr_var
    global cf_var
    cr_var=cn.get()
    cf.delete(0,END)
    if(cr_var=='Ms office'):
        cf.insert(0,1800)
        cf_var=int(cf.get())
    elif(cr_var=='Tally prime' or cr_var=='Python' or cr_var=='Java'):
        cf.insert(0,3500)
        cf_var=int(cf.get())
    elif(cr_var=='C Language'):
        cf.insert(0,3000)
        cf_var=int(cf.get())
    elif(cr_var=='Adv Python' or cr_var=='Adv Java' or cr_var=='Web design'):
        cf.insert(0,4500)
        cf_var=int(cf.get())
    elif(cr_var=='DTP'):
        cf.insert(0,4000)
        cf_var=int(cf.get())
    elif(cr_var=='Photoshop'):
        cf.insert(0,2000)
    else:
        cf.insert(0,00)

bt1=Button(cframe,text='Get',command=get_fees,bg='gold',font=('',12),padx=10,pady=5)
bt1.grid(row=0,column=4,padx=5,pady=10)

lb4=Label(sframe,text='Surname :',bg='lightblue',font=('',12),padx=10,pady=5)
surn=Entry(sframe,bd=5,width=30)

lb5=Label(sframe,text='Enter Student name :',bg='lightblue',font=('',12),padx=10,pady=5)
sn=Entry(sframe,bd=5,width=30)

lb6=Label(sframe,text='Enter Father Name :',bg='Lightblue',font=('',12),padx=10,pady=5)
fn=Entry(sframe,bd=5,font=('',12))
sn_var=sn.get()
fn_var=fn.get()

lb7=Label(sframe,text='Date of join :',bg='lightblue',font=('',12),padx=10,pady=5)
dj=Entry(sframe,bd=5,font=('',12))
dj.insert(0,dt)

lb8=Label(sframe,text='Qualification ',bg='lightblue',font=('',12),padx=10,pady=5)
qn=Entry(sframe,bd=5,font=('',12))

lb9=Label(sframe,text='Mobile number :',bg='lightblue',font=('',12),padx=10,pady=5)
mn=Entry(sframe,bd=5,font=('',12))

lb10=Label(sframe,text='Whatsapp number :',bg='lightblue',font=('',12),padx=10,pady=5)
wn=Entry(sframe,bd=5,font=('',12))

lb11=Label(sframe,text='Address :',bg='lightblue',font=('',12),padx=10,pady=5)
adrs=Entry(sframe,bd=5,font=('',12))

lb12=Label(sframe,text='Advance paid :',bg='lightblue',font=('',12),padx=10,pady=5)
adv=Entry(sframe,bd=5,font=('',12))

lb4.grid(row=0,column=0,padx=5,pady=10)
surn.grid(row=0,column=1,padx=5,pady=10)
lb5.grid(row=0,column=2,padx=5,pady=10)
sn.grid(row=0,column=3,padx=5,pady=10)

lb6.grid(row=1,column=0,padx=5,pady=10)
fn.grid(row=1,column=1,padx=5,pady=10)
lb7.grid(row=1,column=2,padx=5,pady=10)
dj.grid(row=1,column=3,padx=5,pady=10)

lb8.grid(row=2,column=0,padx=5,pady=10)
qn.grid(row=2,column=1,padx=5,pady=10)
lb9.grid(row=2,column=2,padx=5,pady=10)
mn.grid(row=2,column=3,padx=5,pady=10)

lb10.grid(row=3,column=0,padx=5,pady=10)
wn.grid(row=3,column=1,padx=5,pady=10)
lb11.grid(row=3,column=2,padx=5,pady=10)
adrs.grid(row=3,column=3,padx=5,pady=10)

lb12.grid(row=4,column=1,padx=5,pady=10)
adv.grid(row=4,column=3,padx=5,pady=10)


def submit():
        from datetime import datetime
        wb=load_workbook('data.xlsx')
        ws1=wb['Sheet1']
        ws2=wb['Sheet2']
        sur_var=surn.get()
        sn_var=sn.get()
        fn_var=fn.get()
        dj_var=dj.get()
        doj=datetime.strptime(dj_var,'%d/%m/%y')
        qn_var=qn.get()
        mn_var=mn.get()
        wn_var=wn.get()
        adrs_var=adrs.get()
        adv_var=adv.get()

        ws1.append([cr_var,cf_var,sur_var,sn_var,fn_var,doj,qn_var,mn_var,wn_var,adrs_var])
        ws2.append([sn_var,fn_var,adv_var,doj])
        wb.save('data.xlsx')
        messagebox.showinfo(title='Message',message='Data saved successfuly ')

def clr():
        cn.delete(0,END)
        cf.delete(0, END)
        surn.delete(0, END)
        sn.delete(0, END)
        fn.delete(0, END)
        qn.delete(0,END)
        mn.delete(0, END)
        wn.delete(0, END)
        adrs.delete(0,END)
        adv.delete(0,END)

def show_all():
        r=tk.TK()
        r.geometry('1100X700')
        tx.Text(r,height=50,width=200,font=('',12))
        tx.config(bg='lightgray')
        tx.place(x=20,y=10)
        r.state('zoomed')

bframe=LabelFrame(root,bg='black')
bframe.pack(fill=X)
bt2=Button(bframe,text='SAVE',width=20,command=submit)
bt3=Button(bframe,text='Search',width=20,command=reading_mulptisheets_form.get_student)
lb3.grid(row=0,column=4,padx=20,pady=10)
b_clr=Button(bframe,text='NEXT',width=20,command=clr)
b_show=Button(bframe,text='Pay Fees',width=20,command=paymentfile_application.pay_fees)

bt2.grid(row=0,column=0,padx=20,pady=10)
b_clr.grid(row=0,column=1,padx=20,pady=10)
#b_pay.grid(row=0,column=2,padx=20,pady=10)
b_show.grid(row=0,column=3,padx=20,pady=10)

root.mainloop()

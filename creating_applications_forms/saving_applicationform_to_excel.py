
from openpyxl import Workbook,load_workbook

import tkinter as tk
from tkinter import *
from tkinter import ttk

root=tk.Tk()
root.title('  APPLICATION FORM  ')

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

root.geometry('900x900')
lb1=Label(root,text='APPLICATION FORM',font=('',20))
    
lb2=Label(root,text='Select course : ')
cn=ttk.Combobox(root,width=40,values=course)

lb3=Label(root,text='Course fees : ')
cf=Entry(root,width=22)

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
        
bu1=Button(root,text='Get fees',command=get_fees)



def submit():
    wb=load_workbook('data.xlsx')  ##wb=workbook
    ws=wb.active                   ##ws=worksheet 
    e1_var=e1.get()
    e2_var=e2.get()
    e3_var=e3.get()
    e4_var=e4.get()
    e5_var=e5.get()
    e6_var=e6.get()
    e7_var=e7.get('1.0','end-1c')   #to get multiple lines

    ws.append([e1_var,e2_var,e3_var,cr_var,cf_var,e4_var,e5_var,e6_var,e7_var])
    wb.save('data.xlsx')
    
bu2=Button(root,text='  Submit form  ',command=submit)



lb4=Label(root,text='Enter Student name : ')
e1=Entry(root,width=35)
e1_var=e1.get()
lb5=Label(root,text='Enter Father name : ')
e2=Entry(root,width=35)
e2_var=e2.get()
lb6=Label(root,text='Date of join : ')
e3=Entry(root,width=35)
lb7=Label(root,text='Qualification : ')
e4=Entry(root,width=35)
lb8=Label(root,text='Mobile number : ')
e5=Entry(root,width=35)
lb9=Label(root,text='Whatsapp number : ')
e6=Entry(root,width=35)
lb10=Label(root,text='Address : ')
e7=Text(root,width=35,height=6)


lb1.place(x=270,y=30)
lb2.place(x=50,y=80)
cn.place(x=160,y=82)
lb3.place(x=450,y=80)
cf.place(x=560,y=82)
bu1.place(x=720,y=78)

lb4.place(x=50,y=130)
e1.place(x=180,y=132)
lb5.place(x=50,y=180)
e2.place(x=180,y=185)
lb6.place(x=50,y=230)
e3.place(x=180,y=232)
lb7.place(x=50,y=280)
e4.place(x=180,y=282)
lb8.place(x=50,y=330)
e5.place(x=180,y=332)
lb9.place(x=50,y=380)
e6.place(x=180,y=382)
lb10.place(x=50,y=430)
e7.place(x=180,y=432)

bu2.place(x=400,y=570)

root.mainloop()

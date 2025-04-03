import tkinter as tk
from tkinter import *
from tkinter import ttk, messagebox

root = tk.Tk()
root.title('APPLICATION FORM')
root.geometry('700x700')
root.configure(bg="#f4f4f4")  # Light grey background

course = [
    'Ms Office', 'Tally prime', 'C Language', 'Python', 'Adv Python',
    'Java', 'Adv Java', 'DTP', 'Photoshop', 'Web design'
]

# Heading
lb1 = Label(root, text='APPLICATION FORM', font=('Arial', 24, 'bold'), bg="#f4f4f4", fg="#2c3e50")
lb1.pack(pady=20)

# Course Selection
frame1 = Frame(root, bg="#f4f4f4")
frame1.pack(pady=5)
lb2 = Label(frame1, text='Select Course:', font=('Arial', 12, 'bold'), bg="#f4f4f4", fg="#34495e")
lb2.grid(row=0, column=0, padx=10, pady=5)
cn = ttk.Combobox(frame1, width=30, values=course, font=('Arial', 12))
cn.grid(row=0, column=1, padx=10, pady=5)
bu1 = Button(frame1, text='Get Fees', command=lambda: get_fees(), font=('Arial', 12), bg="#2980b9", fg="white")
bu1.grid(row=0, column=2, padx=10, pady=5)

lb3 = Label(frame1, text='Course Fees:', font=('Arial', 12, 'bold'), bg="#f4f4f4", fg="#34495e")
lb3.grid(row=1, column=0, padx=10, pady=5)
cf = Entry(frame1, width=25, font=('Arial', 12))
cf.grid(row=1, column=1, padx=10, pady=5)


# Function to get fees
def get_fees():
    cr_var = cn.get()
    cf.delete(0, END)
    fees_dict = {
        'Ms Office': 1800, 'Tally prime': 3500, 'C Language': 3000,
        'Python': 3500, 'Adv Python': 4500, 'Java': 3500,
        'Adv Java': 4500, 'DTP': 4000, 'Photoshop': 2000, 'Web design': 4500
    }
    cf.insert(0, fees_dict.get(cr_var, "0"))


# Student Details
frame2 = Frame(root, bg="#f4f4f4")
frame2.pack(pady=10)

fields = [
    "Enter Student Name:", "Enter Father Name:", "Date of Join:",
    "Qualification:", "Mobile Number:", "Whatsapp Number:", "Address:"
]

entries = []
for i, text in enumerate(fields):
    lb = Label(frame2, text=text, font=('Arial', 12, 'bold'), bg="#f4f4f4", fg="#34495e")
    lb.grid(row=i, column=0, padx=10, pady=5, sticky=W)

    if text == "Address:":
        e = Text(frame2, width=30, height=4, font=('Arial', 12))
    else:
        e = Entry(frame2, width=30, font=('Arial', 12))

    e.grid(row=i, column=1, padx=10, pady=5)
    entries.append(e)


# Submit Function
def submit_form():
    for i, e in enumerate(entries):
        if isinstance(e, Text):
            if not e.get("1.0", END).strip():
                messagebox.showwarning("Validation Error", f"{fields[i]} cannot be empty!")
                return
        else:
            if not e.get().strip():
                messagebox.showwarning("Validation Error", f"{fields[i]} cannot be empty!")
                return
    messagebox.showinfo("Success", "Form Submitted Successfully!")


# Submit Button
bu2 = Button(root, text='Submit Form', command=submit_form, font=('Arial', 14, 'bold'), bg="#27ae60", fg="white",
             padx=10, pady=5)
bu2.pack(pady=20)

root.mainloop()

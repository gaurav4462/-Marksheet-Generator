from tkinter import *
from tkinter import font
from tkinter.font import BOLD
import win32api
import os


window = Tk()

window.title('Mark_sheet')

def print_marsheet():
    a = int(user1_value.get())
    b = int(user2_value.get())
    c = int(user3_value.get()) 
    d = int(user4_value.get())
    e = int(user5_value.get())
    p = int(user1_value.get())+int(user2_value.get())+int(user3_value.get())+int(user4_value.get())+int(user5_value.get())
    g = (int(p)/500)*100
    h = grade(g)
    i = grade(a)
    j = grade(b)
    k = grade(c)
    l = grade(d)
    m = grade(e)
    with open("marksheet.txt","w") as f:
        f.write("   RAJASTHAN TECHNICAL UNIVERSITY KOTA",)
        f.write("\n")
        f.write(f"\nRoll No.       : {x7.get()}")
        f.write(f"\nName           : {x6.get()}")
        f.write(f"\nFather's Name  : {x8.get()}")
        f.write(f"\nMother's Name  : {x9.get()}")
        f.write(f"\nBranch         : {x10.get()}")
        f.write("\n                                ")
        f.write("\nSR NO. Course   Totle   Mark  Grade")
        f.write(f"\n 1.    DBMS      100     {a}    {i}")
        f.write(f"\n 2.    TOC       100     {b}    {j}")
        f.write(f"\n 3.    COA       100     {c}    {k}")
        f.write(f"\n 4.    Physics   100     {d}    {l}")
        f.write(f"\n 5.    Chemistry 100     {e}    {m}")
        f.write("\n")
        f.write(f"\nGrand Totle      :  {p}")
        f.write(f"\nTotle Grade      :  {h}")
        f.write(f"\nTotle Percentage :  {g}%")
        f.write("\n")
        f.write("\nYear Of Exam     No. Of Subject")
        f.write(f"\n  May 2022          6  ")
        f.write("\n")
        f.write("\nResult Declared   Prof. Dhindra mathur")
        f.write(f"\n  {x12.get()}  \tController Of Examination")




        Pdf = win32api.ShellExecute(0,"print","marksheet.txt",None,".",0)

        os.startfile(f"{Pdf}.pdf")





def new_window():
    root = Toplevel(window)
    root.title("Print MarkSheet")
    root.geometry("1200x700")



    z1 = int(user1_value.get())
    z2 = int(user2_value.get())
    z3 = int(user3_value.get()) 
    z4 = int(user4_value.get())
    z5 = int(user5_value.get())
    
    z6=grade(z1)
    z7=grade(z2)
    z8=grade(z3)
    z9=grade(z4)
    z10=grade(z5)

    z11 = int(user1_value.get())+int(user2_value.get())+int(user3_value.get())+int(user4_value.get())+int(user5_value.get())
    z12 = (int(z11)/500)*100
    z13=grade(z12)


    f1 = Frame(root, bd=5)
    f1.place(relwidth=1, relheight=0.2)

    f5 = Frame(root, bd=5)
    f5.place(rely = 0.2,relwidth=0.1, relheight=0.5)
    f6 = Frame(root, bd=5)
    f6.place(relx=0.1, rely = 0.2,relwidth=0.4, relheight=0.5)
    f9 = Frame(root, bd=5)
    f9.place(relx=0.5, rely = 0.2,relwidth=0.125, relheight=0.5)
    f7 = Frame(root,  bd=5)
    f7.place(relx=0.625,rely = 0.2,relwidth=0.125, relheight=0.5)
    f8 = Frame(root, bd=5)
    f8.place(relx=0.75,rely = 0.2,relwidth=0.125, relheight=0.5)
    f3 = Frame(root, bd=5)
    f3.place(rely = 0.7,relwidth=1, relheight=0.2)
    f4 = Frame(root,  bd=5)
    f4.place(rely = 0.9,relwidth=1, relheight=0.1)

    l1 =Label(f1,text="RAJASTHAN TECHNICAL UNIVERSITY KOTA")
    l1.configure(font=(BOLD, 25))
    l1.place(relwidth=1, relheight=0.2)
    l2 =Label(f1,text="Roll No.             :",font=BOLD)
    l2.place(rely=0.25, relwidth=0.13, relheight=0.175)
    l3 =Label(f1,text="Name                :",font=BOLD)
    l3.place(rely=0.45, relwidth=0.13, relheight=0.175)
    l4 =Label(f1,text="Father`s Name  :",font=BOLD)
    l4.place(rely=0.65, relwidth=0.13, relheight=0.175)
    l5 =Label(f1,text="Mother`s Name  :",font=BOLD)
    l5.place(rely=0.85, relwidth=0.13, relheight=0.175)


    l6 =Label(f1,text=x7.get(),font=BOLD)
    l6.place(relx=0.2,rely=0.25, relwidth=0.2, relheight=0.175)
    l7 =Label(f1,text=x6.get(),font=BOLD)
    l7.place(relx=0.2,rely=0.45, relwidth=0.2, relheight=0.175)
    l8 =Label(f1,text=x8.get(),font=BOLD)
    l8.place(relx=0.2,rely=0.65, relwidth=0.2, relheight=0.175)
    l9 =Label(f1,text=x9.get(),font=BOLD)
    l9.place(relx=0.2,rely=0.85, relwidth=0.2, relheight=0.175)

    l8 =Label(f1,text="Branch             :",font=BOLD)
    l8.place(relx=0.5,rely=0.85, relwidth=0.13, relheight=0.175)
    l9 =Label(f1,text=x10.get(),font=BOLD)
    l9.place(relx=0.675,rely=0.85, relwidth=0.1, relheight=0.175)


    l10 =Label(f5,text="SR. NO",font=BOLD)
    l10.place(rely=0.05, relwidth=1, relheight=0.15)
    l11 =Label(f6,text="Course Title",font=BOLD)
    l11.place(rely=0.05, relwidth=1, relheight=0.15)
    o1 =Label(f9,text="Max Marks",font=BOLD)
    o1.place(rely=0.05, relwidth=1, relheight=0.15)
    l12 =Label(f7,text="Marks",font=BOLD)
    l12.place(rely=0.05, relwidth=1, relheight=0.15)
    l13 =Label(f8,text="Grade",font=BOLD)
    l13.place(rely=0.05, relwidth=1, relheight=0.15)


    l14 =Label(f5,text="1. ",font=BOLD)
    l14.place(rely=0.2, relwidth=1, relheight=0.125)
    l15 =Label(f5,text="2. ",font=BOLD)
    l15.place(rely=0.35, relwidth=1, relheight=0.125)
    l16 =Label(f5,text="3. ",font=BOLD)
    l16.place(rely=0.475, relwidth=1, relheight=0.125)
    l17 =Label(f5,text="4. ",font=BOLD)
    l17.place(rely=0.625, relwidth=1, relheight=0.125)
    l18 =Label(f5,text="5. ",font=BOLD)
    l18.place(rely=0.775, relwidth=1, relheight=0.125)

    l19 =Label(f6,text="    DBMS    ",font=BOLD)
    l19.place(rely=0.2, relwidth=1, relheight=0.125)
    l20 =Label(f6,text="    TOC       ",font=BOLD)
    l20.place(rely=0.35, relwidth=1, relheight=0.125)
    l21 =Label(f6,text="    COA      ",font=BOLD)
    l21.place(rely=0.475, relwidth=1, relheight=0.125)
    l22 =Label(f6,text="    Physics     ",font=BOLD)
    l22.place(rely=0.625, relwidth=1, relheight=0.125)
    l23 =Label(f6,text="    Chemistry   ",font=BOLD)
    l23.place(rely=0.775, relwidth=1, relheight=0.125)


    l24 =Label(f7,text=user1_value.get(),font=BOLD)
    l24.place(rely=0.2, relwidth=1, relheight=0.125)
    l25 =Label(f7,text=user2_value.get(),font=BOLD)
    l25.place(rely=0.35, relwidth=1, relheight=0.125)
    l26 =Label(f7,text=user3_value.get(),font=BOLD)
    l26.place(rely=0.475, relwidth=1, relheight=0.125)
    l27 =Label(f7,text=user4_value.get(),font=BOLD)
    l27.place(rely=0.625, relwidth=1, relheight=0.125)
    l28 =Label(f7,text=user5_value.get(),font=BOLD)
    l28.place(rely=0.775, relwidth=1, relheight=0.125)

    l29 =Label(f9,text="100",font=BOLD)
    l29.place(rely=0.2, relwidth=1, relheight=0.125)
    l30 =Label(f9,text="100",font=BOLD)
    l30.place(rely=0.35, relwidth=1, relheight=0.125)
    l31 =Label(f9,text="100",font=BOLD)
    l31.place(rely=0.475, relwidth=1, relheight=0.125)
    l32 =Label(f9,text="100",font=BOLD)
    l32.place(rely=0.625, relwidth=1, relheight=0.125)
    l33 =Label(f9,text="100",font=BOLD)
    l33.place(rely=0.775, relwidth=1, relheight=0.125)

    l34 =Label(f8,text=z6,font=BOLD)
    l34.place(rely=0.2, relwidth=1, relheight=0.125)
    l35 =Label(f8,text=z7,font=BOLD)
    l35.place(rely=0.35, relwidth=1, relheight=0.125)
    l36 =Label(f8,text=z8,font=BOLD)
    l36.place(rely=0.475, relwidth=1, relheight=0.125)
    l37 =Label(f8,text=z9,font=BOLD)
    l37.place(rely=0.625, relwidth=1, relheight=0.125)
    l38 =Label(f8,text=z10,font=BOLD)
    l38.place(rely=0.775, relwidth=1, relheight=0.125)

    l39 =Label(f3,text="Grand Total     :",font=BOLD)
    l39.place( relwidth=0.3, relheight=0.15)
    l40 =Label(f3,text=z11,font=BOLD)
    l40.place(relx=0.3 , relwidth=0.1, relheight=0.15)

    l39 =Label(f3,text="Total Grade    :",font=BOLD)
    l39.place(rely=0.25,relwidth=0.3, relheight=0.15)
    l40 =Label(f3,text=z13,font=BOLD)
    l40.place(relx=0.3 ,rely=0.25 ,relwidth=0.1, relheight=0.15)

    l41 =Label(f3,text="Month and Year of Examination ",font=BOLD)
    l41.place( rely=0.45,relwidth=0.3, relheight=0.15)
    l42 =Label(f3,text="May 2022 ",font=BOLD)
    l42.place( rely=0.7,relwidth=0.2, relheight=0.15)

    l43 =Label(f3,text="No. of Subject Offered",font=BOLD)
    l43.place(relx=0.4, rely=0.45,relwidth=0.2, relheight=0.15)
    l44 =Label(f3,text="   6   ",font=BOLD)
    l44.place( relx=0.4,rely=0.7,relwidth=0.2, relheight=0.15)
    l45 =Label(f3,text="Percentage",font=BOLD)
    l45.place(relx=0.7, rely=0.45,relwidth=0.2, relheight=0.15)
    l46 =Label(f3,text=z12,font=BOLD)
    l46.place( relx=0.7,rely=0.7,relwidth=0.15, relheight=0.15)
    l47 =Label(f3,text="%",font=BOLD)
    l47.place( relx=0.8,rely=0.7,relwidth=0.1, relheight=0.15)

    l48 =Label(f4,text="Result Declared on",font=BOLD)
    l48.place(relx=0.1,relwidth=0.2, relheight=0.45)
    l49 =Label(f4,text=x12.get(),font=BOLD)
    l49.place(relx=0.1,rely=0.5,relwidth=0.2, relheight=0.45)

    l50 =Label(f4,text="Prof. Dhindra Mathur",font=BOLD)
    l50.place(relx= 0.4,relwidth=0.3, relheight=0.45)
    l51 =Label(f4,text="Controller Of Examination",font=BOLD)
    l51.place(relx=0.4,rely=0.5,relwidth=0.3, relheight=0.45)



    b3=Button(f4, text="Print",bg="green",command=print_marsheet)
    b3.place(relx= 0.8,rely=0.7,relwidth=0.1,relheight=0.45)



    window.mainloop()







def grade(number):
    if(number>90):
        return "AA"

    elif(number<=90 and number>=80):
         return "AB"

    elif(number<80 and number>=60): 
         return "BB" 

    else:
         return "FF"

    

def Mark_sheet():
    a = int(user1_value.get())
    b = int(user2_value.get())
    c = int(user3_value.get()) 
    d = int(user4_value.get())
    e = int(user5_value.get())
    f = int(user1_value.get())+int(user2_value.get())+int(user3_value.get())+int(user4_value.get())+int(user5_value.get())
    g = (int(f)/500)*100
    h = grade(g)
    i = grade(a)
    j = grade(b)
    k = grade(c)
    l = grade(d)
    m = grade(e)

    t6.delete("1.0", END)
    t6.insert(END, a)

    t7.delete("1.0", END)
    t7.insert(END, b)

    t8.delete("1.0", END)
    t8.insert(END, c)
    
    t9.delete("1.0", END)
    t9.insert(END, d)
    
    t10.delete("1.0", END)
    t10.insert(END, e)
    
    a2.delete("1.0", END)
    a2.insert(END, f)
    
    a4.delete("1.0", END)
    a4.insert(END, g)

    a6.delete("1.0", END)
    a6.insert(END, h)

    t11.delete("1.0", END)
    t11.insert(END, i)
    
    t12.delete("1.0", END)
    t12.insert(END, j)
    
    t13.delete("1.0", END)
    t13.insert(END, k)
    
    t14.delete("1.0", END)
    t14.insert(END, l)
    
    t15.delete("1.0", END)
    t15.insert(END, m)
    
    



e1 = Label(window, text="Enter Your Number In DBMS       ")
e2 = Label(window, text="Enter Your Number In TOC         ")
e3 = Label(window, text="Enter Your Number In COA          ")
e4 = Label(window, text="Enter Your Number In Physics    ")
e5 = Label(window, text="Enter Your Number In Chemistry")
user1_value = StringVar()
user2_value = StringVar()
user3_value = StringVar()
user4_value = StringVar()
user5_value = StringVar()
e6 = Entry(window, textvariable=user1_value)
e7 = Entry(window, textvariable=user2_value)
e8 = Entry(window, textvariable=user3_value)
e9 = Entry(window, textvariable=user4_value)
e10 =Entry(window, textvariable=user5_value)

t1 = Label(window, text="DBMS      ")
t2 = Label(window, text="TOC         ")
t3 = Label(window, text="COA        ")
t4 = Label(window, text="Physics   ")
t5 = Label(window, text="Chemistry")
t6 = Text(window, height=0.1, width = 20)
t7 = Text(window, height=0.1, width = 20)
t8 = Text(window, height=0.1, width = 20)
t9 = Text(window, height=0.1, width = 20)
t10 = Text(window, height=0.1, width = 20)
t11 = Text(window, height=0.1, width = 20)
t12 = Text(window, height=0.1, width = 20)
t13 = Text(window, height=0.1, width = 20)
t14 = Text(window, height=0.1, width = 20)
t15 = Text(window, height=0.1, width = 20)

t16 = Label(window, text="Subject")
t17 = Label(window, text="Marks")
t18 = Label(window, text="Grade")

a1 = Label(window, text="Totle marks        ")
a2 = Text(window, height=0.1, width = 20)
a3 = Label(window, text="Totle Percenttage")
a4 = Text(window, height=0.1, width = 20)
a5 = Label(window, text="Totle Grade         ")
a6 = Text(window, height=0.1, width = 20)

x1 = Label(window, text="Enter Your Name           ")
x2 = Label(window, text="Enter Your Roll No.         ")
x6 = Entry(window)
x7 = Entry(window)
x3 = Label(window, text="Enter Your Father Name")
x4 = Label(window, text="Enter Your Mother Name")
x8 = Entry(window)
x9 = Entry(window)
x5 = Label(window, text="Enter Your Branch            ")
x10 = Entry(window)
x11 = Label(window, text="Enter Date                        ")
x12 = Entry(window)
x13 = Label(window, text="")

x14 = Label(window, text="--------------------------")
x15 = Label(window, text="--------------------------")
x16 = Label(window, text="--------------------------")
x17 = Label(window, text="--------------------------")


b1 = Button(window, text="Click", command=Mark_sheet)
b2 = Button(window, text="Generate MarkSheet" ,command=new_window)

x1.grid(row=0, column=0)
x6.grid(row=0, column=1)
x2.grid(row=0, column=2)
x7.grid(row=0, column=3)

x3.grid(row=1, column=0)
x8.grid(row=1, column=1)
x4.grid(row=1, column=2)
x9.grid(row=1, column=3)

x11.grid(row=2, column=0)
x12.grid(row=2, column=1)
x5.grid(row=2, column=2)
x10.grid(row=2, column=3)

x13.grid(row=3, column=1)

e1.grid(row=4, column=1)
e6.grid(row=4, column=2)

e2.grid(row=5, column=1)
e7.grid(row=5, column=2)

e3.grid(row=6, column=1)
e8.grid(row=6, column=2)

e4.grid(row=7, column=1)
e9.grid(row=7, column=2)

e5.grid(row=8, column=1)
e10.grid(row=8, column=2)

b1.grid(row=9, column=3)

x14.grid(row=10, column=0)
x15.grid(row=10, column=1)
x16.grid(row=10, column=2)
x17.grid(row=10, column=3)

t16.grid(row=11, column=0)
t17.grid(row=11, column=1)
t18.grid(row=11, column=2)


t1.grid(row=12, column=0)
t6.grid(row=12, column=1)
t11.grid(row=12, column=2)

t2.grid(row=13, column=0)
t7.grid(row=13, column=1)
t12.grid(row=13, column=2)

t3.grid(row=14, column=0)
t8.grid(row=14, column=1)
t13.grid(row=14, column=2)

t4.grid(row=15, column=0)
t9.grid(row=15, column=1)
t14.grid(row=15, column=2)

t5.grid(row=16, column=0)
t10.grid(row=16, column=1)
t15.grid(row=16, column=2)

a1.grid(row=17, column=1)
a2.grid(row=17, column=2)


a3.grid(row=18, column=1)
a4.grid(row=18, column=2)

a5.grid(row=19, column=1)
a6.grid(row=19, column=2)

b2.grid(row=20, column=3)

window.mainloop()
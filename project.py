from tkinter import *
from PIL import ImageTk,Image
import pandas as pd
xls = pd.ExcelFile('Attendance_Data.xlsx')
df1 = pd.read_excel(xls, 'Sheet1')
print(df1)

root = Tk() 
root.title("Attendance Record & Management")
root.geometry("400x400")
classroom = NONE
Present = []
Absent = []

def show():
    myLabel = Label(root,text=clicked.get()).pack()
def onClick(idx):
    print(idx) # Print the index value
 #   print(Attended[idx].cget("text"))

def markpresent(x):
    df1['Day1'].iloc[x] = Present[x].cget("text")
    df1.to_excel('Attendance_Data.xlsx',index = False)

def markabsent(x):
    df1['Day1'].iloc[x] = Absent[x].cget("text")
    df1.to_excel('Attendance_Data.xlsx',index = False)

 

def open():
    root.withdraw()
    global classroom
    global Present
    global Absent
    classroom = Toplevel()
    classroom.title(clicked.get())
    classroom.geometry("400x400")
    MyLabel1 = Label(classroom,text="Welcome to Attendance Record for " + clicked.get()).grid(row=0,columnspan=4,padx=10)
    MyLabel2 = Label(classroom,text="Check the Box to mark student attendance").grid(row=1,columnspan=4,padx=10)
    


    for i in range(5):
        name = Label(classroom,text=df1['Name'][i]).grid(row=i+3,column=0)
        c = Checkbutton(classroom,text="Present",variable=Present,onvalue=1,offvalue=0,command = lambda idx = i: markpresent(idx))
        d = Checkbutton(classroom,text="Absent",variable=Absent,onvalue=0,offvalue=1,command = lambda idx = i: markabsent(idx))
        c.grid(row=i+3,column=1)
        d.grid(row=i+3,column=2)
        
        Present.append(c)
        Absent.append(d)
    btn2 = Button(classroom,text="Close Window",command=lambda:[classroom.destroy(),root.deiconify()]).grid(row=11)
    


options = [
    "Course1",
    "Course2"
]

clicked = StringVar()
clicked.set("Select one from below")
drop = OptionMenu(root, clicked, *options).pack()

btn = Button(root,text="Open",command=open).pack()


root.mainloop()

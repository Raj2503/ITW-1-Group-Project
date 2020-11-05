from tkinter import *
from PIL import ImageTk,Image
import pandas as pd
xls = pd.ExcelFile('Attendance_Data.xlsx')
df1 = pd.read_excel(xls, 'Sheet1',index_col = None)

root = Tk() 
root.title("Attendance Record & Management")
root.geometry("400x300")
frame = LabelFrame(root, text= "Welcome to Attendance Manager",padx=100, pady=80)
frame.pack(fill=BOTH,padx=10,pady=10)

classroom = NONE
Present = []
Absent = []
DayN = 0
for i in range(1,41):
    if df1['Day'+str(i)].isnull().sum()==5:
        DayN =i-1
        break
#print("\n\n" + str(DayN) + "\n\n\n")
def show():
    global DayN
    myLabel = Label(root,text=clicked.get()).pack()
def onClick(idx):
    global DayN
    print(idx) # Print the index value
 #   print(Attended[idx].cget("text"))

def markpresent(x):
    global DayN
    df1['Day'+str(int(DayN))].iloc[x] = Present[x].cget("text")
    #print(df1)
    df1.to_excel('Attendance_Data.xlsx',index = False)

def markabsent(x):
    global DayN
    df1['Day'+str(int(DayN))].iloc[x] = Absent[x].cget("text")
    #print(df1)
    df1.to_excel('Attendance_Data.xlsx',index = False)

def percentage():
    root.withdraw()
    record = Toplevel()
    record.title("Attendance Percentage")
    record.geometry("450x400")
    frame = LabelFrame(record, text= "Mark Attendance",padx=50,pady=40)
    frame.grid(padx=10,pady=10)
    MyLabel1 = Label(frame,text="Welcome to Attendance Record for " + clicked.get()).grid(row=0,columnspan=4,padx=10)

    for i in range(5):
        name = Label(frame,text=df1['Name'][i]).grid(row=i+1,column=0)
        attended = (df1.iloc[i,4:]=="Present").sum()
        total = 40 - (df1.iloc[1,4:].isnull()).sum()
        if (attended/total)*100 < 75:
            percent = Label(frame,text = "{:.2f}".format((attended/total)*100) +" %", fg = "red",font = "Times").grid(row=i+1,column=1)
        else:
            percent = Label(frame,text = "{:.2f}".format((attended/total)*100) + " %", fg = "dark green",font = "Times").grid(row=i+1,column=1)
    btn2 = Button(frame,text="Close Window",command=lambda:[record.withdraw(),root.deiconify()]).grid(row=11,column=0,ipadx=20,pady=30)    


def open():
    root.withdraw()
    global classroom
    global Present
    global Absent
    global DayN
    DayN= int(DayN) + 1


    classroom = Toplevel()
    classroom.title(clicked.get())
    classroom.geometry("420x400")
    frame = LabelFrame(classroom, text= "Mark Attendance",padx=50,pady=40)
    frame.grid(padx=10,pady=10)
    MyLabel1 = Label(frame,text="Welcome to Attendance Record for " + clicked.get() + " Day"+str(DayN)).grid(row=0,columnspan=3,padx=10)
    MyLabel2 = Label(frame,text="Check the Box to mark student attendance").grid(row=1,columnspan=3,padx=10,pady=(10,10))
    


    for i in range(5):
        name = Label(frame,text=df1['Name'][i]).grid(row=i+3,column=0)
        c = Checkbutton(frame,text="Present",variable=Present,onvalue=1,offvalue=0,command = lambda idx = i: markpresent(idx))
        d = Checkbutton(frame,text="Absent",variable=Absent,onvalue=0,offvalue=1,command = lambda idx = i: markabsent(idx))
        c.grid(row=i+3,column=1)
        d.grid(row=i+3,column=2)
        
        Present.append(c)
        Absent.append(d)
    btn2 = Button(frame,text="Close Window",command=lambda:[classroom.withdraw(),root.deiconify()]).grid(row=11,column=0,ipadx=20,pady=30)
    


options = [
    "Course1",
    "Course2"
]

clicked = StringVar()
clicked.set("Select one from below")
drop = OptionMenu(frame, clicked, *options).grid(row=1,column=2,columnspan=2)

btn = Button(frame,text="Mark Attendance",command=open).grid(row=2,column=2,columnspan=2,pady=10)
btn2 = Button(frame,text="Show Attendance Percentage",command=percentage).grid(row=3,column=2,columnspan=2,pady=10)

root.mainloop()

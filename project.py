from tkinter import *
from PIL import ImageTk,Image
import pandas as pd
xls = pd.ExcelFile('Attendance_Data.xlsx')
df1 = pd.read_excel(xls, 'Sheet1',index_col = None)

root = Tk() 
root.title("Attendance Record & Management")
root.geometry("400x400")
classroom = NONE
Present = []
Absent = []
DayN = 0
present=0
total=0
for i in range(len(df1)) :
    for j in range(0,42):
            
        if(df1.iloc[i, j]=="Present"):
            present+=1
        if(df1.iloc[i, j]=="Present" or df1.iloc[i, j]=="Absent"):
            total+=1
    print(present)
    present=0
for i in range(1,41):
    if df1['Day'+str(i)].isnull().sum()==5:
        DayN =i-1
        break
print("\n\n" + str(DayN) + "\n\n\n")
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



def open():
    root.withdraw()
    global classroom
    global Present
    global Absent
    global DayN
    DayN= int(DayN) + 1


    classroom = Toplevel()
    classroom.title(clicked.get())
    classroom.geometry("400x400")
    MyLabel1 = Label(classroom,text="Welcome to Attendance Record for " + clicked.get() + " Day"+str(DayN)).grid(row=0,columnspan=4,padx=10)
    MyLabel2 = Label(classroom,text="Check the Box to mark student attendance").grid(row=1,columnspan=4,padx=10)
    


    for i in range(5):
        name = Label(classroom,text=df1['Name'][i]).grid(row=i+3,column=0)
        c = Checkbutton(classroom,text="Present",variable=Present,onvalue=1,offvalue=0,command = lambda idx = i: markpresent(idx))
        d = Checkbutton(classroom,text="Absent",variable=Absent,onvalue=0,offvalue=1,command = lambda idx = i: markabsent(idx))
        c.grid(row=i+3,column=1)
        d.grid(row=i+3,column=2)
        
        Present.append(c)
        Absent.append(d)
    btn2 = Button(classroom,text="Close Window",command=lambda:[classroom.withdraw(),root.deiconify()]).grid(row=11)
    


options = [
    "Course1",
    "Course2"
]

clicked = StringVar()
clicked.set("Select one from below")
drop = OptionMenu(root, clicked, *options).pack()

btn = Button(root,text="Open",command=open).pack()


root.mainloop()
from tkinter import *
from PIL import ImageTk,Image
import pandas as pd

df = pd.ExcelFile('output.xlsx')
Sheet_names_list = df.sheet_names
dflist = []
#print(Sheet_names_list)
for sheet in Sheet_names_list :
   dflist.append(df.parse(sheet_name=sheet,index_col=None))

root = Tk()
root.title("Attendance Record & Management")
root.geometry("400x300")
frame = Frame(root,padx=100, pady=80)
frame.pack(fill=BOTH,padx=10,pady=10)

classroom = NONE
Present = []
Absent = []
DayN = 1



#print("\n\n" + str(DayN) + "\n\n\n")

def show():
    global DayN
    myLabel = Label(root,text=clicked.get()).pack()
def onClick(idx):
    global DayN
    print(idx) # Print the index value
 #   print(Attended[idx].cget("text"))

def markpresent(idx,x):
    global DayN
    dflist[x]['Day'+str(int(DayN))].iloc[idx] = Present[idx].cget("text")
    #print(df1)


def markabsent(idx,x):
    global DayN
    dflist[x]['Day'+str(int(DayN))].iloc[idx] = Absent[idx].cget("text")
    #print(df1)

def OnFrameConfigure(canvas):
    canvas.configure(scrollregion=canvas.bbox("all"))

def save():
    with pd.ExcelWriter('output.xlsx') as writer:
        for i in range(6):
            dflist[i].to_excel(writer,sheet_name="Sheet_name_" + str(i+1),index=False)

def percentage():
    root.withdraw()
    record = Toplevel()
    record.title("Attendance Percentage")
    record.geometry("450x400")
    y=2
    canvas = Canvas(record, borderwidth=0)
    frame = Frame(canvas)
    vsb = Scrollbar(record, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set, width=1200, height=80)       

    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((4,4), window=frame, anchor="nw", tags="frame")

    # be sure that we call OnFrameConfigure on the right canvas
    frame.bind("<Configure>", lambda event, canvas=canvas: OnFrameConfigure(canvas))
    MyLabel1 = Label(frame,text="Welcome to Attendance Record for " + clicked.get()).grid(row=0,columnspan=4,padx=10)
    x = 0
    if clicked.get()=='Course1':
        x=0
    if clicked.get()=='Course2':
        x=1
    if clicked.get()=='Course3':
        x=2
    if clicked.get()=='Course4':
        x=3
    if clicked.get()=='Course5':
        x=4
    if clicked.get()=='Course6':
        x=5
    for i in range(len(dflist[x].index)):
        rollno = Label(frame,text=dflist[x]['#'][i]).grid(row=i+1,column=0)
        name = Label(frame,text=dflist[x]['Name'][i]).grid(row=i+1,column=1)
        attended = (dflist[x].iloc[i,2:]=="Present").sum()
        total = 40 - (dflist[x].iloc[i,2:].isnull()).sum()
        #print(attended,total,END)
        if (attended/total)*100 < 75:
            percent = Label(frame,text = "{:.2f}".format((attended/total)*100) +" %", fg = "red",font = "Times").grid(row=i+1,column=2)
        else:
            percent = Label(frame,text = "{:.2f}".format((attended/total)*100) + " %", fg = "dark green",font = "Times").grid(row=i+1,column=2)
    btn2 = Button(frame,text="Close Window",command=lambda:[record.withdraw(),root.deiconify()]).grid(row=len(dflist[x].index)+5,column=0,ipadx=20,pady=30)    


def open():
    root.withdraw()
    global classroom
    global Present
    global Absent
    global DayN
    x = 0
    if clicked.get()=='Course1':
        x=0
    if clicked.get()=='Course2':
        x=1
    if clicked.get()=='Course3':
        x=2
    if clicked.get()=='Course4':
        x=3
    if clicked.get()=='Course5':
        x=4
    if clicked.get()=='Course6':
        x=5
    
    for i in range(1,41):
        if dflist[x]['Day'+str(i)].isnull().sum()==5:
            DayN =i
            break
	

    classroom = Toplevel()
    classroom.title(clicked.get())
    classroom.geometry("420x400")
    canvas = Canvas(classroom, borderwidth=0)
    frame = Frame(canvas)
    vsb = Scrollbar(classroom, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set, width=1200, height=80)       

    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((4,4), window=frame, anchor="nw", tags="frame")

    # be sure that we call OnFrameConfigure on the right canvas
    frame.bind("<Configure>", lambda event, canvas=canvas: OnFrameConfigure(canvas))
    MyLabel1 = Label(frame,text="Welcome to Attendance Record for " + clicked.get() + " Day"+str(DayN)).grid(row=0,columnspan=3,padx=10)
    MyLabel2 = Label(frame,text="Check the Box to mark student attendance").grid(row=1,columnspan=3,padx=10,pady=(10,10))
    
	
	
    for i in range(len(dflist[x].index)):
        name = Label(frame,text=dflist[x]['Name'][i]).grid(row=i+3,column=0)
        c = Checkbutton(frame,text="Present",variable=Present,onvalue=1,offvalue=0,command = lambda idx = i: markpresent(idx,x))
        d = Checkbutton(frame,text="Absent",variable=Absent,onvalue=0,offvalue=1,command = lambda idx = i: markabsent(idx,x))
        c.grid(row=i+3,column=1)
        d.grid(row=i+3,column=2)
        
        Present.append(c)
        Absent.append(d)
	
    btn2 = Button(frame,text="Close Window",command=lambda:[classroom.withdraw(),root.deiconify(),save()]).grid(row=len(dflist[x].index)+5,column=0,ipadx=20,pady=30)

    


options = [
    "Course1",
    "Course2",
    "Course3",
    "Course4",
    "Course5",
    "Course6",
]

clicked = StringVar()
clicked.set("Select one from below")
drop = OptionMenu(frame, clicked, *options).grid(row=1,column=2,columnspan=2)

btn = Button(frame,text="Mark Attendance",command=lambda:[open()]).grid(row=2,column=2,columnspan=2,pady=10)
btn2 = Button(frame,text="Show Attendance Percentage",command=lambda:[percentage()]).grid(row=3,column=2,columnspan=2,pady=10)

root.mainloop()

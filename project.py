from tkinter import *
import pandas as pd

df = pd.ExcelFile('output.xlsx')
Sheet_names_list = df.sheet_names
dflist = []

for sheet in Sheet_names_list :
   dflist.append(df.parse(sheet_name=sheet,index_col=None))

root = Tk()
root.title("Attendance Record & Management")
root.geometry("400x400")
root.iconbitmap('71404_student_attendance.ico')

frame = LabelFrame(root,text="Welcome to Attendance Manager",padx=100, pady=40)
frame.pack(fill=BOTH,padx=10,pady=10)

classroom = NONE
Present = []
Absent = []
DayN = 1
R_no = IntVar()
f_name = StringVar()
day = IntVar()
course = StringVar()

def show():
    global DayN
    myLabel = Label(root,text=clicked.get()).pack()

def onClick(idx):
    global DayN
    print(idx) 

def markpresent(idx,x):
    global DayN
    dflist[x]['Day'+str(int(DayN))].iloc[idx] = Present[idx].cget("text")

def markabsent(idx,x):
    global DayN
    dflist[x]['Day'+str(int(DayN))].iloc[idx] = Absent[idx].cget("text")

def OnFrameConfigure(canvas):
    canvas.configure(scrollregion=canvas.bbox("all"))

def save():
    with pd.ExcelWriter('output.xlsx') as writer:
        for i in range(6):
            dflist[i].to_excel(writer,sheet_name="Sheet_name_" + str(i+1),index=False)

def updateAddData():
    for i in range(6):
        df2 = {'RollNo':dflist[0].iloc[-1]['RollNo'] + 1,'Name': f_name.get()}
        dflist[i]=dflist[i].append(df2,ignore_index = True)
    save()

def updateDelData():
    name = f_name.get()
    rno = R_no.get()
    for i in range(6):
        dflist[i] = dflist[i][(dflist[i].Name != name) & (dflist[i].RollNo != rno)]
    save()

def newstudent():
    global f_name
    root.withdraw()
    newstu = Toplevel()
    newstu.title("Enroll New Student")
    newstu.geometry("400x200")
    newstu.iconbitmap('71404_student_attendance.ico')

    frame = LabelFrame(newstu,text="New Student Details",padx=50,pady=10)
    frame.pack(padx=10,pady=10)

    RollLabel = Label(frame,text = "Roll Number")
    RollLabel.grid(row=0,column=0)
    RollNo = Label(frame,text = dflist[0].iloc[-1]['RollNo'] + 1).grid(row=0,column=1)
    
    f_name = Entry(frame, width=30,borderwidth=3)
    f_name.grid(row=1, column=1, padx=20, pady=(10, 0))
    f_name_label = Label(frame, text="Full Name")
    f_name_label.grid(row=1, column=0, pady=(10, 0))
    myButton = Button(frame,text="Eroll Student", command=lambda:[newstu.withdraw(),root.deiconify(),updateAddData()])
    myButton.grid(row=4,padx=20,pady=30,columnspan=2)

def percentage():
    root.withdraw()
    record = Toplevel()
    record.title("Attendance Percentage")
    record.geometry("390x400")
    record.iconbitmap('71404_student_attendance.ico')

    canvas = Canvas(record, borderwidth=0)
    frame = LabelFrame(canvas,text="Student Attendance Record",padx=50,pady=10)
    vsb = Scrollbar(record, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set, width=1200, height=80)       
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((4,4), window=frame, anchor="nw", tags="frame")
    frame.bind("<Configure>", lambda event, canvas=canvas: OnFrameConfigure(canvas))
    
    MyLabel1 = Label(frame,text="Welcome to Attendance Record for " + clicked.get()).grid(row=0,columnspan=4,padx=10)
    x = 0
    if clicked.get()=='Course1':
        x=0
    elif clicked.get()=='Course2':
        x=1
    elif clicked.get()=='Course3':
        x=2
    elif clicked.get()=='Course4':
        x=3
    elif clicked.get()=='Course5':
        x=4
    elif clicked.get()=='Course6':
        x=5

    for i in range(len(dflist[x].index)):
        rollno = Label(frame,text=dflist[x]['RollNo'][i]).grid(row=i+1,column=0)
        name = Label(frame,text=dflist[x]['Name'][i]).grid(row=i+1,column=1)
        attended = (dflist[x].iloc[i,2:]=="Present").sum()
        total = 40 - (dflist[x].iloc[i,2:].isnull()).sum()

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
    elif clicked.get()=='Course2':
        x=1
    elif clicked.get()=='Course3':
        x=2
    elif clicked.get()=='Course4':
        x=3
    elif clicked.get()=='Course5':
        x=4
    elif clicked.get()=='Course6':
        x=5
    
    for i in range(1,41):
        if dflist[x]['Day'+str(i)].isnull().sum()==len(dflist[x].index):
            DayN =i
            break

    classroom = Toplevel()
    classroom.title(clicked.get())
    classroom.geometry("420x400")
    classroom.iconbitmap('71404_student_attendance.ico')

    canvas = Canvas(classroom, borderwidth=0)
    frame = LabelFrame(canvas,text="Mark Student Attendance",padx=50,pady=10)
    vsb = Scrollbar(classroom, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set, width=1200, height=80)       
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((4,4), window=frame, anchor="nw", tags="frame")
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
	
    btn2 = Button(frame,text="Save Record",command=lambda:[classroom.withdraw(),root.deiconify(),save()]).grid(row=len(dflist[x].index)+5,column=0,ipadx=20,pady=30)


def delstudent():
    global f_name
    global R_no
    root.withdraw()
    dele = Toplevel()
    dele.title("Delete Student")
    dele.geometry("400x200")
    dele.iconbitmap('71404_student_attendance.ico')
    frame = LabelFrame(dele,text="Student Details",padx=50,pady=10)
    frame.pack(padx=10,pady=10)
    
    RollLabel = Label(frame,text = "Roll Number")
    RollLabel.grid(row=0,column=0)
    
    R_no = Entry(frame, width=30,borderwidth=3)
    R_no.grid(row=0, column=1, padx=20, pady=(10, 0))

    f_name = Entry(frame, width=30,borderwidth=3)
    f_name.grid(row=1, column=1, padx=20, pady=(10, 0))
    f_name_label = Label(frame, text="Full Name")
    f_name_label.grid(row=1, column=0, pady=(10, 0))

    myButton = Button(frame,text="Delete Student", command=lambda:[dele.withdraw(),root.deiconify(),updateDelData()]).grid(row=4,padx=20,pady=30,columnspan=2)

def markAttendance():
    global Present
    global Absent
    global DayN
    global day
    x=0
    if course.get()=='Course1':
        x=0
    elif course.get()=='Course2':
        x=1
    elif course.get()=='Course3':
        x=2
    elif course.get()=='Course4':
        x=3
    elif course.get()=='Course5':
        x=4
    elif course.get()=='Course6':
        x=5
    
    DayN = int(day.get())
	
    classroom = Toplevel()
    classroom.title(clicked.get())
    classroom.geometry("420x400")
    classroom.iconbitmap('71404_student_attendance.ico')

    canvas = Canvas(classroom, borderwidth=0)
    frame = LabelFrame(canvas,text="Mark Student Attendance",padx=50,pady=10)
    vsb = Scrollbar(classroom, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set, width=1200, height=80)
    vsb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    canvas.create_window((4,4), window=frame, anchor="nw", tags="frame")

    frame.bind("<Configure>", lambda event, canvas=canvas: OnFrameConfigure(canvas))
    MyLabel1 = Label(frame,text="Welcome to Attendance Record for " + clicked.get() + " Day"+str(DayN)).grid(row=0,columnspan=3,padx=10)
    MyLabel2 = Label(frame,text="Check the Box to Mark Student Attendance").grid(row=1,columnspan=3,padx=10,pady=(10,10))

    for i in range(len(dflist[x].index)):
        name = Label(frame,text=dflist[x]['Name'][i]).grid(row=i+3,column=0)
        c = Checkbutton(frame,text="Present",variable=Present,onvalue=1,offvalue=0,command = lambda idx = i: markpresent(idx,x))
        d = Checkbutton(frame,text="Absent",variable=Absent,onvalue=0,offvalue=1,command = lambda idx = i: markabsent(idx,x))
        c.grid(row=i+3,column=1)
        d.grid(row=i+3,column=2)
        
        Present.append(c)
        Absent.append(d)
	
    btn2 = Button(frame,text="Save Record",command=lambda:[classroom.withdraw(),root.deiconify(),save()]).grid(row=len(dflist[x].index)+5,column=0,ipadx=20,pady=30)
    
def editAttendance():
    global course
    global day
    root.withdraw()
    edit = Toplevel()
    edit.title("Edit Attendance Record")
    edit.geometry("400x200")
    edit.iconbitmap('71404_student_attendance.ico')
    frame = LabelFrame(edit,text="Select Course & Day",padx=50,pady=10)
    frame.pack(padx=10,pady=10)

    options = [
    "Course1",
    "Course2",
    "Course3",
    "Course4",
    "Course5",
    "Course6",
    ]

    course = StringVar()
    course.set("Select one from below")
    courseLabel = Label(frame,text = "Course")
    courseLabel.grid(row=1,column=0)
    drop = OptionMenu(frame, clicked, *options).grid(row=1,column=1,columnspan=2)

    dayLabel = Label(frame,text = "Select Day")
    dayLabel.grid(row=2,column=0,pady=(10, 0))
    day = Entry(frame, width=30,borderwidth=3)
    day.grid(row=2, column=1, padx=20, pady=(10, 0))    

    myButton = Button(frame,text="Mark Attendance Again", command=lambda:[edit.withdraw(),markAttendance()]).grid(row=4,padx=20,pady=30,columnspan=2)


courLabel = Label(frame,text = "Select a course from below to\nView or Mark Attendance")
courLabel.grid(row=0,column=0,columnspan=4,pady=(0,10))
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
btn3 = Button(frame,text="Add New Student",command=lambda:[newstudent()]).grid(row=4,column=2,columnspan=2,pady=10,ipadx=13)
btn4 = Button(frame,text="Delete Student",command=lambda:[delstudent()]).grid(row=5,column=2,columnspan=2,pady=10,ipadx=21)
btn5 = Button(frame,text="Edit Attendance Record",command=lambda:[editAttendance()]).grid(row=6,column=2,columnspan=2,pady=10)

root.mainloop()
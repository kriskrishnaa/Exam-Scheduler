#pip install tabula-py
from tabula import *
from Tkinter import *
from Tkinter import Tk
from tkFileDialog import askopenfilename
put = "yes"
r = 0
while put == "yes":
    r = r+1

    def dept1():
        global filename1
        filename1 = askopenfilename()
        global save
        save = "converted"
        f = filename1.rfind("/")
        for i in range(f,len(filename1)-4):
            save = save + filename1[i]
        save = save + ".csv"
        
        
    def merge():
        convert_into(filename1, save, pages="all",output_format="csv")
        pop()
    def pop():
        global app
        app = Tk()
        app.title("POPUP")
        app.geometry("300x125")
        label = Label(app, text="DO YOU WANT TO CONTINUE", height=0, width=50).place(x=.6,y=20)
        button1 = Button(app, text="YES", width=10, command=b1).place(x=70,y=80)
        button2 = Button(app, text="NO", width=10, command=b2).place(x=180,y=80)
        app.mainloop()
    def b1():
        global put
        put = "yes"
        app.destroy()
        gui.destroy()
    def b2():
        app.destroy()
        gui.destroy()
        global put
        if put == "yes":
            put = ""
        else:
            put = ""

    #print "done"
    global gui
    gui = Tk()
    gui.geometry("440x300")
    gui.title("pdf to excel converter")
    mlabel = Label(gui,text="SELECT THE FILES TO CONVERT",fg="black",height=5, width=50).place(x=35,y=50)
    mbutton2 = Button(gui,text="choose",fg="black",width="8",height="1",command = dept1).place(x=165,y=125)
    mbutton1 = Button(gui,text="CONVERT",fg="black",width="12",height="2",command = merge).place(x=255,y=225)
    gui.mainloop()


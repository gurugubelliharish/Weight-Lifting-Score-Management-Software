from tkinter import *
#from tkscrolledframe import ScrolledFrame
#from tkinter import messagebox
from math import *

cnc1 = Tk()
cnc1.title("CNC CODER")
cnc1.geometry("730x835+0+0")
cnc1.configure(background="black")

def close():
    cnc1.destroy()
    
entry=Entry(cnc1, width=30)
button=Button(cnc1, text="ENTER", command=close, fg="white", bg="black")
label=Label(cnc1, text= "enter feed: ", bg="grey", fg="white")
feed=StringVar()
menu=OptionMenu(cnc1, feed,"feed/min","feed/rev")
#label
label.grid(row=1, column=0,ipadx=20)
entry.grid(row=2,column=0)
button.grid(row=3,column=0)
#menu
menu.grid(row=1, column=1, columnspan=3,pady=5, padx=10, ipadx=90)



cnc1.mainloop()


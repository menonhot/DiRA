import Pmw
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import sys
sys.path.insert(0, '../../apps')
sys.path.insert(0,'../../Date')
from nota4 import takeVal,notaPMB

tab1 = Tk()
tab1.title('DiRA_PMB 1.1')


#NOTA
Label(tab1,text='No Nota:').grid(row=0,column=0)
Entry(tab1,text="").grid(row=0,column=1)
prhalTemp = StringVar()
Label(tab1, text="Perihal:").grid(row=1,column=0)
opt_perihal = Pmw.OptionMenu(tab1, menubutton_textvariable=prhalTemp,items=('L15','Metric Ton','Reservasi','ShipTo bentrok dengan Plumpang'),menubutton_width=16)
opt_perihal.grid(row=1,column=1)
coList = ['NO','NOPAL','LO','PRODUK','*MT|L15','VOLUME','KETERANGAN']
Label(tab1,text=coList[0]).grid(row=2,column=0)
Label(tab1,text=coList[1]).grid(row=2,column=1)
Label(tab1,text=coList[2]).grid(row=2,column=2)
Label(tab1,text=coList[3]).grid(row=2,column=3)
Label(tab1,text=coList[4]).grid(row=2,column=4)
Label(tab1,text=coList[5]).grid(row=2,column=5)
Label(tab1,text=coList[6]).grid(row=2,column=6)
height = 11
width = 7
#createMultipleEntryTkinter
for i in range(3,height): #Rows
    for j in range(width): #Columns
        ent = Entry(tab1, text="")
        ent.grid(row=i, column=j)
def attFunc1():
    global attachPath1
    attach1.set(askopenfilename(title='Browse',filetypes = (("jpeg files","*.jpg"),("png files","*.png"))))
    attachPath1 = attach1.get()
def attFunc2():
    global attachPath2
    attach2.set(askopenfilename(title='Browse',filetypes = (("jpg files","*.jpg"),("png files","*.png"),("jpeg files","*.jpeg"))))
    attachPath2 = attach2.get()
def attFunc3():
    global attachPath3
    attach3.set(askopenfilename(title='Browse',filetypes = (("jpg files","*.jpg"),("png files","*.png"),("jpeg files","*.jpeg"))))
    attachPath3 = attach3.get()
def attFunc4():
    global attachPath4
    attach4.set(askopenfilename(title='Browse',filetypes = (("jpg files","*.jpg"),("png files","*.png"),("jpeg files","*.jpeg"))))
    attachPath4 = attach4.get()
def attFunc5():
    global attachPath5
    attach5.set(askopenfilename(title='Browse',filetypes = (("jpg files","*.jpg"),("png files","*.png"),("jpeg files","*.jpeg"))))
    attachPath5 = attach5.get()
def attFunc6():
    global attachPath6
    attach6.set(askopenfilename(title='Browse',filetypes = (("jpg files","*.jpg"),("png files","*.png"),("jpeg files","*.jpeg"))))
    attachPath6 = attach6.get()
def attFunc7():
    global attachPath7
    attach7.set(askopenfilename(title='Browse',filetypes = (("jpg files","*.jpg"),("png files","*.png"),("jpeg files","*.jpeg"))))
    attachPath7 = attach7.get()
def attFunc8():
    global attachPath8
    attach8.set(askopenfilename(title='Browse',filetypes = (("jpg files","*.jpg"),("png files","*.png"),("jpeg files","*.jpeg"))))
    attachPath8 = attach8.get()
Label(tab1,text="ATTACHMENT").grid(row=2,column=7)
attach1 = StringVar()
attach2 = StringVar()
attach3 = StringVar()
attach4 = StringVar()
attach5 = StringVar()
attach6 = StringVar()
attach7 = StringVar()
attach8 = StringVar()
attPath1 = Entry(tab1, width=15, bg="white",textvariable=attach1)
attPath2 = Entry(tab1, width=15, bg="white",textvariable=attach2)
attPath3 = Entry(tab1, width=15, bg="white",textvariable=attach3)
attPath4 = Entry(tab1, width=15, bg="white",textvariable=attach4)
attPath5 = Entry(tab1, width=15, bg="white",textvariable=attach5)
attPath6 = Entry(tab1, width=15, bg="white",textvariable=attach6)
attPath7 = Entry(tab1, width=15, bg="white",textvariable=attach7)
attPath8 = Entry(tab1, width=15, bg="white",textvariable=attach8)
attPath1.grid(row=3,column=7)
attPath2.grid(row=4,column=7)
attPath3.grid(row=5,column=7)
attPath4.grid(row=6,column=7)
attPath5.grid(row=7,column=7)
attPath6.grid(row=8,column=7)
attPath7.grid(row=9,column=7)
attPath8.grid(row=10,column=7)
Button(tab1,width = 15,text="Browse",command=attFunc1).grid(row=3,column=8)
Button(tab1,width = 15,text="Browse",command=attFunc2).grid(row=4,column=8)
Button(tab1,width = 15,text="Browse",command=attFunc3).grid(row=5,column=8)
Button(tab1,width = 15,text="Browse",command=attFunc4).grid(row=6,column=8)
Button(tab1,width = 15,text="Browse",command=attFunc5).grid(row=7,column=8)
Button(tab1,width = 15,text="Browse",command=attFunc6).grid(row=8,column=8)
Button(tab1,width = 15,text="Browse",command=attFunc7).grid(row=9,column=8)
Button(tab1,width = 15,text="Browse",command=attFunc8).grid(row=10,column=8)
def runShit():
	imageList = [attPath1.get(),attPath2.get(),attPath3.get(),attPath4.get(),attPath5.get(),attPath6.get(),attPath7.get(),attPath8.get()]
	imageList = [i for i in imageList if i != '']
	tabValue = takeVal(tab1.winfo_children())
	site = 'PMB'
	nota(tabValue,prhalTemp.get(),imageList,site)
	t2 = Toplevel(root)
	Label(t2, text='Email sent\n Please Chill the F out').pack(padx=50, pady=50)
	t2.withdraw()
	t2.grab_set()
	t2.deiconify()
	t2.transient(root)
	def eraseEntry(winfo):
		entButt = [attPath1,attPath2,attPath3,attPath4,attPath5,attPath6,attPath7,attPath8]
		[i.delete(0, END) for i in entButt]
		for i in winfo:
			if i.winfo_class() == 'Entry':
				i.delete(0, END)
	def desButton():
		eraseEntry(tab3.winfo_children())
		t2.destroy()
	Button(t2,text='OK',command=desButton).pack()

Button(tab1,width = 15,text='generate',command=runShit).grid(row=11, column=3)
tab1.mainloop()
#///

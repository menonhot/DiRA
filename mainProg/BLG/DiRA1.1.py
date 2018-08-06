import Pmw
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import sys
sys.path.insert(0, '../../apps')
sys.path.insert(0,'../../Date')
from thrusum2 import mainProg
from allDateFormat import getDate,getDashDate,getYesterday,getNow
from jbbClosing import loOpen,mailContent
from emailThingy import testAddr,eAddress,logIn
from nota4 import takeVal,nota,testNota

root = Tk()
nb = ttk.Notebook(root)
root.title('DiRA_BLG')
tab1 = Frame(nb)
tab2 = Frame(nb)
tab3 = Frame(nb)

#Variable of UI
a = StringVar()
d = StringVar()
c = StringVar()
#d = StringVar()
varA = StringVar()
varB = StringVar()
varC = StringVar()
saveFileThru = StringVar()
saveFileSum = StringVar()

#Function of UI
def browsefunc1():
    global textPath1
    a.set(askopenfilename(title='Pilih Data LO'))
    textPath1 = a.get()
def browsefunc2():
    global textPath2
    d.set(askopenfilename(title='Pilih Meter Totals'))
    textPath2 = d.get()
def browsefunc3():
    global textPath3
    c.set(askopenfilename(title='Pilih Carrier'))
    textPath3 = c.get()

#element of UserInterface
Label(tab1, text="DataLO:").grid(row=1,column=0)
Label(tab1, text="MeterTotals:").grid(row=2,column=0)
Label(tab1, text="Carrier:").grid(row=3,column=0)

Label(tab1, text="Otentikasi").grid(row=5, column=1)
Label(tab1, text="Prepared by:").grid(row = 6, column=0)
Label(tab1, text="Ops Patra:").grid(row=7,column=0)
Label(tab1, text="Penyalur:").grid(row=8,column=0)

Button(tab1, text="Browse", command=browsefunc1).grid(row=1, column=2)
Button(tab1, text="Browse", command=browsefunc2).grid(row=2, column=2)
Button(tab1, text="Browse", command=browsefunc3).grid(row=3, column=2)

pathEntry1=Entry(tab1, width=30, bg="white",textvariable=a)
pathEntry2=Entry(tab1, width=30, bg="white",textvariable=d)
pathEntry3=Entry(tab1, width=30, bg="white",textvariable=c)

nameEntry1=Entry(tab1, width=30,textvariable=varA)
nameEntry2=Entry(tab1, width=30,textvariable=varB)
nameEntry3=Entry(tab1, width=30,textvariable=varC)

nameEntry1.grid(row=6,column=1)
nameEntry2.grid(row=7,column=1)
nameEntry3.grid(row=8,column=1)

pathEntry1.grid(row=1,column=1)
pathEntry2.grid(row=2,column=1)
pathEntry3.grid(row=3,column=1)

def getReport():
    site = 'BLG'
    f,g = mainProg(site,textPath2,textPath1,textPath3,varA.get(),varB.get(),varC.get())
    #3. SAVING THE FILE
    fileDate = getDate()
    summaryName = 'Laporan Summary {}'.format(fileDate)
    saveFileSum.set(asksaveasfilename(title='Save As Laporan Summary',initialfile=summaryName,defaultextension='.xlsx'))
    savedFileSum = saveFileSum.get()
    thruputName = 'Laporan Thruput {}'.format(fileDate)
    saveFileThru.set(asksaveasfilename(title='Save As Laporan Thruput', initialfile=thruputName,defaultextension='.xlsx'))
    savedFileThru = saveFileThru.get()
    f.save(savedFileSum)
    g.save(savedFileThru)
    entryList = [nameEntry1,nameEntry2,nameEntry3,pathEntry1,pathEntry2,pathEntry3]
    [i.delete(0,END) for i in entryList]
#tab 2 jbbClosingSend
varG = StringVar()
cats = StringVar()
def browseDataLoJBB():
    global textPath7
    varG.set(askopenfilename(title='Pilih Data LO'))
    textPath7 = varG.get()
Label(tab2,text="CLOSING:").grid(row=0,column=0)
opt_cat = Pmw.OptionMenu(tab2, menubutton_textvariable=cats,items=('TODAY','YESTERDAY'),menubutton_width=16)
opt_cat.grid(row=0,column=1)
Label(tab2, text="Data LO:").grid(row=1,column=0)
Button(tab2, text="Browse", command=browseDataLoJBB).grid(row=1, column=2)
pathEntry7=Entry(tab2, width=30, bg="white",textvariable=varG)
pathEntry7.grid(row=1,column=1)
def exeJbb():
    if cats.get() == 'TODAY':
        site = 'BLG'
        strObjd = getNow()
        fromAddr, toAddr, pswd = eAddress(site)
        subs ='JBB Closing %s' % strObjd.strftime('%d-%m-%Y')
        msg = mailContent(site,subs,pathEntry7.get(),cats.get())
        logIn(fromAddr,toAddr,pswd,msg)
        pathEntry7.delete(0,END)
    elif cats.get() == 'YESTERDAY':
        site = 'BLG'
        strObjd = getYesterday()
        fromAddr, toAddr, pswd = eAddress(site)
        subs ='JBB Closing %s' % strObjd.strftime('%d-%m-%Y')
        msg = mailContent(site, subs, pathEntry7.get(),cats.get())
        logIn(fromAddr,toAddr,pswd,msg)
        pathEntry7.delete(0,END)

#NOTA
Label(tab3,text='No Nota:').grid(row=0,column=0)
Entry(tab3,text="").grid(row=0,column=1)
prhalTemp = StringVar()
Label(tab3, text="Perihal:").grid(row=1,column=0)
opt_perihal = Pmw.OptionMenu(tab3,menubutton_textvariable=prhalTemp,items=('Trip Number tanpa LO','MT UNIK','DO Pecah','Konservasi'),menubutton_width=16)
opt_perihal.grid(row=1,column=1)
coList = ['NO','NOPOL','LO','PRODUK','VOLUME','KETERANGAN']
Label(tab3,text=coList[0]).grid(row=2,column=0)
Label(tab3,text=coList[1]).grid(row=2,column=1)
Label(tab3,text=coList[2]).grid(row=2,column=2)
Label(tab3,text=coList[3]).grid(row=2,column=3)
Label(tab3,text=coList[4]).grid(row=2,column=4)
Label(tab3,text=coList[5]).grid(row=2,column=5)
height = 11
width = 6
#createMultipleEntryTkinter
for i in range(3,height): #Rows
    for j in range(width): #Columns
        b = Entry(tab3, text="")
        b.grid(row=i, column=j)
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
Label(tab3,text="ATTACHMENT").grid(row=2,column=7)
attach1 = StringVar()
attach2 = StringVar()
attach3 = StringVar()
attach4 = StringVar()
attach5 = StringVar()
attach6 = StringVar()
attach7 = StringVar()
attach8 = StringVar()
attPath1 = Entry(tab3, width=15, bg="white",textvariable=attach1)
attPath2 = Entry(tab3, width=15, bg="white",textvariable=attach2)
attPath3 = Entry(tab3, width=15, bg="white",textvariable=attach3)
attPath4 = Entry(tab3, width=15, bg="white",textvariable=attach4)
attPath5 = Entry(tab3, width=15, bg="white",textvariable=attach5)
attPath6 = Entry(tab3, width=15, bg="white",textvariable=attach6)
attPath7 = Entry(tab3, width=15, bg="white",textvariable=attach7)
attPath8 = Entry(tab3, width=15, bg="white",textvariable=attach8)
attPath1.grid(row=3,column=7)
attPath2.grid(row=4,column=7)
attPath3.grid(row=5,column=7)
attPath4.grid(row=6,column=7)
attPath5.grid(row=7,column=7)
attPath6.grid(row=8,column=7)
attPath7.grid(row=9,column=7)
attPath8.grid(row=10,column=7)
Button(tab3,width = 5,text="Browse",command=attFunc1).grid(row=3,column=8)
Button(tab3,width = 5,text="Browse",command=attFunc2).grid(row=4,column=8)
Button(tab3,width = 5,text="Browse",command=attFunc3).grid(row=5,column=8)
Button(tab3,width = 5,text="Browse",command=attFunc4).grid(row=6,column=8)
Button(tab3,width = 5,text="Browse",command=attFunc5).grid(row=7,column=8)
Button(tab3,width = 5,text="Browse",command=attFunc6).grid(row=8,column=8)
Button(tab3,width = 5,text="Browse",command=attFunc7).grid(row=9,column=8)
Button(tab3,width = 5,text="Browse",command=attFunc8).grid(row=10,column=8)


def runShit():
	imageList = [attPath1.get(),attPath2.get(),attPath3.get(),attPath4.get(),attPath5.get(),attPath6.get(),attPath7.get(),attPath8.get()]
	imageList = [i for i in imageList if i != '']
	tabValue = takeVal(tab3.winfo_children())
	site = 'BLG'
	testNota(tabValue,prhalTemp.get(),imageList,site)
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
#END OF THE ROAD
Button(tab1,text='generate',command=getReport).grid(row=10,column=1)
Button(tab2,text='send',command=exeJbb).grid(row=3,column=1)
Button(tab3,text='generate',command=runShit).grid(row=11, column=3)
nb.add(tab1, text= 'Thruput Summary')
nb.add(tab2, text= 'JBB Closing')
nb.add(tab3, text= 'Autosend Nota')
nb.grid()
root.mainloop()


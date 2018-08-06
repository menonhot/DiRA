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

root = Tk()
nb = ttk.Notebook(root)
root.title('DiRA')
tab1 = Frame(nb)
tab2 = Frame(nb)

#Variable of UI
a = StringVar()
b = StringVar()
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
    b.set(askopenfilename(title='Pilih Meter Totals'))
    textPath2 = b.get()
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
pathEntry2=Entry(tab1, width=30, bg="white",textvariable=b)
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
        #loOpen(strObjd,pathEntry8.get())
        #fromAddr,toAddr,pswd = eAddress(site)
        fromAddr, toAddr, pswd = eAddress(site)
        subs ='JBB Closing %s' % strObjd.strftime('%d-%m-%y')
        msg = mailContent(site,subs,pathEntry7.get(),cats.get())
        logIn(fromAddr,toAddr,pswd,msg)
        pathEntry7.delete(0,END)
        #pathEntry8.delete(0,END)
    elif cats.get() == 'YESTERDAY':
        site = 'BLG'
        strObjd = getYesterday()
        #loOpen(strObjd,pathEntry8.get())
        #fromAddr,toAddr,pswd = eAddress(site)
        fromAddr, toAddr, pswd = eAddress(site)
        subs ='JBB Closing %s' % strObjd.strftime('%d-%m-%y')
        msg = mailContent(site, subs, pathEntry7.get(),cats.get())
        logIn(fromAddr,toAddr,pswd,msg)
        pathEntry7.delete(0,END)
        #pathEntry8.delete(0,END)

#END OF THE ROAD
Button(tab1,text='generate',command=getReport).grid(row=10,column=1)
Button(tab2,text='send',command=exeJbb).grid(row=3,column=1)
nb.add(tab1, text= 'Thruput Summary')
nb.add(tab2, text= 'JBB Closing')
nb.grid()
root.mainloop()


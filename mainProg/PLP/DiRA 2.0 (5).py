# update loPerJam - loPerJam2
from tkinter import *
from tkinter import ttk
import Pmw
import os
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import sys
sys.path.insert(0, '../../apps')
sys.path.insert(0,'../../Date')
from thrusum import mainProgDet
from allDateFormat import getDate,getDashDate,getYesterday,getNow,getYesterdayStr
from jbbClosing import loOpen,mailContent,fetchSite,jbbReport
from emailThingy import testAddr,eAddress,logIn
from loPerJam2 import sendLO
from nota2 import takeVal,nota

root = Tk()
nb = ttk.Notebook(root)
root.title('DiRA')
tab1 = Frame(nb)
tab2 = Frame(nb)
tab3 = Frame(nb)
tab4 = Frame(nb)
tab5 = Frame(nb)
tab6 = Frame(nb)

Label(tab1,text='No Nota:').grid(row=0,column=0)
Entry(tab1,text="").grid(row=0,column=1)
prhalTemp = StringVar()
Label(tab1, text="Perihal:").grid(row=1,column=0)
opt_perihal = Pmw.OptionMenu(tab1,menubutton_textvariable=prhalTemp,items=('Trip Number tanpa LO','MT UNIK','DO Pecah','Konservasi'),menubutton_width=16)
opt_perihal.grid(row=1,column=1)
coList = ['NO','NOPOL','LO','PRODUK','VOLUME','KETERANGAN']
Label(tab1,text=coList[0]).grid(row=2,column=0)
Label(tab1,text=coList[1]).grid(row=2,column=1)
Label(tab1,text=coList[2]).grid(row=2,column=2)
Label(tab1,text=coList[3]).grid(row=2,column=3)
Label(tab1,text=coList[4]).grid(row=2,column=4)
Label(tab1,text=coList[5]).grid(row=2,column=5)
height = 11
width = 6
#createMultipleEntryTkinter
for i in range(3,height): #Rows
    for j in range(width): #Columns
        b = Entry(tab1, text="")
        b.grid(row=i, column=j)
#attachment
'''attachPath1 = ''
attachPath2 = ''
attachPath3 = ''
attachPath4 = ''
attachPath5 = ''
attachPath6 = ''
attachPath7 = ''
attachPath8 = ''
'''
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
Button(tab1,width = 5,text="Browse",command=attFunc1).grid(row=3,column=8)
Button(tab1,width = 5,text="Browse",command=attFunc2).grid(row=4,column=8)
Button(tab1,width = 5,text="Browse",command=attFunc3).grid(row=5,column=8)
Button(tab1,width = 5,text="Browse",command=attFunc4).grid(row=6,column=8)
Button(tab1,width = 5,text="Browse",command=attFunc5).grid(row=7,column=8)
Button(tab1,width = 5,text="Browse",command=attFunc6).grid(row=8,column=8)
Button(tab1,width = 5,text="Browse",command=attFunc7).grid(row=9,column=8)
Button(tab1,width = 5,text="Browse",command=attFunc8).grid(row=10,column=8)


def runShit():
	#imList = [attachPath1,attachPath2,attachPath3,attachPath4,attachPath5,attachPath6,attachPath7,attachPath8]
	imageList = [attPath1.get(),attPath2.get(),attPath3.get(),attPath4.get(),attPath5.get(),attPath6.get(),attPath7.get(),attPath8.get()]
	imageList = [i for i in imageList if i != '']
	tabValue = takeVal(tab1.winfo_children())
	nota(tabValue,prhalTemp.get(),imageList)
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
		eraseEntry(tab1.winfo_children())
		t2.destroy()
	Button(t2,text='OK',command=desButton).pack()
#THRUPUT SUMMARY
#tab2
am=StringVar()
an=StringVar()
ao=StringVar()
ap=StringVar()
saveFileThru = StringVar()
saveFileSum = StringVar()
varA = StringVar()
varB = StringVar()
varC = StringVar()
varD = StringVar()
def browsefunc1():
	global textPath1
	am.set(askopenfilename(title='Pilih Data LO'))
	textPath1 = am.get()
def browsefunc2():
	global textPath2
	an.set(askopenfilename(title='Pilih MeterDetails'))
	textPath2 = an.get()
def browsefunc3():
	global textPath3
	ao.set(askopenfilename(title='Pilih Carrier'))
	textPath3 = ao.get()
def browsefunc4():
	global textPath4
	ap.set(askopenfilename(title='Pilih Summary'))
	textPath4 = ap.get()
Label(tab2, text="DataLO:").grid(row=1,column=0)
Label(tab2, text="MeterDetails:").grid(row=2,column=0)
Label(tab2, text="Carrier:").grid(row=3,column=0)
Label(tab2, text="Summary kemaren:").grid(row=4,column=0)
Label(tab2, text="Otentikasi").grid(row=5, column=1)
Label(tab2, text="Prepared by:").grid(row = 6, column=0)
Label(tab2, text="Ops Patra:").grid(row=7,column=0)
Label(tab2, text="Penyalur:").grid(row=8,column=0)
Label(tab2, text="Approved by:").grid(row=9,column=0)

Button(tab2, text="Browse", command=browsefunc1).grid(row=1, column=2)
Button(tab2, text="Browse", command=browsefunc2).grid(row=2, column=2)
Button(tab2, text="Browse", command=browsefunc3).grid(row=3, column=2)
Button(tab2, text="Browse", command=browsefunc4).grid(row=4, column=2)

pathEntry1=Entry(tab2, width=30, bg="white",textvariable=am)
pathEntry2=Entry(tab2, width=30, bg="white",textvariable=an)
pathEntry3=Entry(tab2, width=30, bg="white",textvariable=ao)
pathEntry4=Entry(tab2, width=30, bg="white",textvariable=ap)

nameEntry1=Entry(tab2, width=30,textvariable=varA)
nameEntry2=Entry(tab2, width=30,textvariable=varB)
nameEntry3=Entry(tab2, width=30,textvariable=varC)
nameEntry4=Entry(tab2, width=30,textvariable=varD)

nameEntry1.grid(row=6,column=1)
nameEntry2.grid(row=7,column=1)
nameEntry3.grid(row=8,column=1)
nameEntry4.grid(row=9,column=1)

pathEntry1.grid(row=1,column=1)
pathEntry2.grid(row=2,column=1)
pathEntry3.grid(row=3,column=1)
pathEntry4.grid(row=4,column=1)
def generateReport():
	site = 'PLP'
	f, g = mainProgDet(site, textPath2, textPath1, textPath3, textPath4, varA.get(), varB.get(), varC.get(),varD.get())
	# 3. SAVING THE FILE
	fileDate = getDate()
	summaryName = 'Laporan Summary {}'.format(fileDate)
	saveFileSum.set(asksaveasfilename(title='Save As Laporan Summary', initialfile=summaryName, defaultextension='.xlsx'))
	savedFileSum = saveFileSum.get()
	thruputName = 'Laporan Thruput {}'.format(fileDate)
	saveFileThru.set(asksaveasfilename(title='Save As Laporan Thruput', initialfile=thruputName, defaultextension='.xlsx'))
	savedFileThru = saveFileThru.get()
	f.save(savedFileSum)
	g.save(savedFileThru)
	entryList = [nameEntry1, nameEntry2, nameEntry3,nameEntry4, pathEntry1, pathEntry2, pathEntry3, pathEntry4]
	[i.delete(0, END) for i in entryList]
#dataLOPERJAM
varE = StringVar()
varF = StringVar()
def browseDataLo():
	global textPath5
	varE.set(askopenfilename(title='Pilih Data LO'))
	textPath5 = varE.get()
def browseLoOpen():
	global textPath6
	varF.set(askopenfilename(title='Pilih LO Open'))
	textPath6 = varF.get()
Label(tab3, text="Data LO:").grid(row=1,column=0)
Label(tab3, text="LO Open:").grid(row=2,column=0)
Button(tab3, text="Browse", command=browseDataLo).grid(row=1, column=2)
Button(tab3, text="Browse", command=browseLoOpen).grid(row=2, column=2)
pathEntry5=Entry(tab3, width=30, bg="white",textvariable=varE)
pathEntry6=Entry(tab3, width=30, bg="white",textvariable=varF)
pathEntry5.grid(row=1,column=1)
pathEntry6.grid(row=2,column=1)
def loPerJam():
	sendLO(textPath5,textPath6)
	os.remove(textPath5)
	os.remove(textPath6)
	pathEntry5.delete(0, END)
	pathEntry6.delete(0, END)
#JBB
#tab 4 jbbClosingSend
varG = StringVar()
varH = StringVar()
cats = StringVar()
def browseDataLoJBB():
	global textPath7
	varG.set(askopenfilename(title='Pilih Data LO'))
	textPath7 = varG.get()
def browseLoOpenJBB():
	global textPath8
	varH.set(askopenfilename(title='Pilih LO Open'))
	textPath8 = varH.get()
Label(tab4,text="CLOSING:").grid(row=0,column=0)
opt_cat = Pmw.OptionMenu(tab4, menubutton_textvariable=cats,items=('TODAY','YESTERDAY'),menubutton_width=16)
opt_cat.grid(row=0,column=1)
Label(tab4, text="Data LO:").grid(row=1,column=0)
Label(tab4, text="LO Open:").grid(row=2,column=0)
Button(tab4, text="Browse", command=browseDataLoJBB).grid(row=1, column=2)
Button(tab4, text="Browse", command=browseLoOpenJBB).grid(row=2, column=2)
pathEntry7=Entry(tab4, width=30, bg="white",textvariable=varG)
pathEntry8=Entry(tab4, width=30, bg="white",textvariable=varH)
pathEntry7.grid(row=1,column=1)
pathEntry8.grid(row=2,column=1)
#sendReport
cats2 = StringVar()
Label(tab5,text="CLOSING:").grid(row=0,column=0)
opt_cat2 = Pmw.OptionMenu(tab5, menubutton_textvariable=cats2,items=('TODAY','YESTERDAY'),menubutton_width=16)
opt_cat2.grid(row=0,column=1)
def exeJbb():
	if cats.get() == 'TODAY':
		site = 'PLP'
		strObjd = getNow()
		loOpen(strObjd,pathEntry8.get(),site)
		#fromAddr,toAddr,pswd = eAddress(site)
		fromAddr, toAddr, pswd = eAddress(site)
		subs ='JBB Closing %s' % strObjd.strftime('%d-%m-%Y')
		msg = mailContent(site,subs,pathEntry7.get(),cats.get())
		logIn(fromAddr,toAddr,pswd,msg)
		pathEntry7.delete(0,END)
		pathEntry8.delete(0,END)
	elif cats.get() == 'YESTERDAY':
		site = 'PLP'
		strObjd = getYesterday()
		loOpen(strObjd,pathEntry8.get(),site)
		#fromAddr,toAddr,pswd = eAddress(site)
		fromAddr, toAddr, pswd = eAddress(site)
		subs ='JBB Closing %s' % strObjd.strftime('%d-%m-%y')
		msg = mailContent(site, subs, pathEntry7.get(),cats.get())
		logIn(fromAddr,toAddr,pswd,msg)
		pathEntry7.delete(0,END)
		pathEntry8.delete(0,END)
def fetchSender():
	if cats2.get() == 'TODAY':
		rightNow = getDashDate()
		notif = fetchSite(rightNow)
		if len(notif) != 0:
			for i in notif:
				print(i)
		else:
			print('Good to go, what are u waiting for!')
	elif cats2.get() == 'YESTERDAY':
		rightNow = getYesterdayStr()
		notif = fetchSite(rightNow)
		if len(notif) != 0:
			for i in notif:
				print(i)
		else:
			print('Good to go, what are u waiting for!')
def sendJBBReport():
	if cats2.get() == 'TODAY':
		likeRightnow = getDashDate()
		jbbReport(likeRightNow)
	elif cats2.get() == 'YESTERDAY':
		likeRightNow = getYesterdayStr()
		jbbReport(likeRightNow)
Button(tab2,text='generate',command=generateReport).grid(row=10,column=1)
Button(tab1,text='generate',command=runShit).grid(row=11, column=3)
Button(tab3,text='generate',command=loPerJam).grid(row=3, column=1)
Button(tab4,text='Send Email',command=exeJbb).grid(row=3,column=1)
Button(tab5,text='Fetch Sender',command=fetchSender).grid(row=3,column=1)
Button(tab5,text='Send Report',command=sendJBBReport).grid(row=4,column=1)
nb.add(tab1, text = 'AutoSend Nota')
nb.add(tab2, text = 'ThruSum Report')
nb.add(tab3, text = 'dataLO Perjam')
nb.add(tab4, text = 'JBBClosing')
nb.add(tab5, text = 'JBB Report')
nb.grid()
root.mainloop()

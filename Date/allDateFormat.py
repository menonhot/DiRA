import datetime
from datetime import timedelta

def getNow():
    now = datetime.datetime.today()
    return now
def getDate():
    now = getNow()
    rightNow = now.strftime('%d%m%Y')#ddmmyyyy
    return rightNow
def getDashDate():
    now = getNow()
    rightNow = now.strftime('%d-%m-%Y') #dd-mm-yyyy
    return rightNow
def getFullDate():
    now = getNow()
    rightNow = now.strftime('%A, %d %B %Y') #Monday, dd Month yyyy
    return rightNow
def getDateSpace():
    now = getNow()
    rightNow = now.strftime('%d %B %Y') #dd Month yyyy
    return rightNow
def getDateDash():#stringFormat
    now = getNow()
    rightNow = now.strftime('%d-%b-%Y') #dd-sep-yyyy
    return rightNow
def getDatum():
    now = getNow()
    rightNow= now.strftime('%d-%m-%Y %H.%M') #dd-mm-yyyy hh.mm
    return rightNow
def getDateSlash():
    now = getNow()
    rightNow= now.strftime('%d/%b/%Y') #dd/sep/yyyy
    return rightNow
def getHour():
    now = getNow()
    rightNow= now.strftime('%H.%M') #07.30
    return rightNow
def getYesterday():#DateFormat
    now = getNow()
    minuscule = timedelta(days=1)
    yTday = now - minuscule
    return yTday
def getYesterdayStr():
    now = getYesterday()
    rightNow = now.strftime('%d-%m-%Y')
    return rightNow
def getMidnight(now):
    t = datetime.time(hour=00,minute=00)
    midnight = datetime.datetime.combine(now,t)
    return midnight

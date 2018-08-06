
import datetime
import os
from allDateFormat import getNow,getYesterday

def structFolder(folderName,site):
    now = getNow()
    thisYear = now.strftime('%Y')
    thisMonth= now.strftime('%B')
    thisDay = now.strftime('%d')
    # makeStructuredFolder
    ddPath = '%s/%s/%s/%s/%s' % (folderName,site,thisYear, thisMonth, thisDay)
    if not os.path.exists(ddPath):
        os.makedirs(ddPath)
        print('Folders created successfully\n{}'.format(ddPath))
    return ddPath
def structFolderMan(folderName,site,nowYesterday):
    nowYesterday = datetime.datetime.strptime(nowYesterday,'%d-%m-%Y')
    thisYear = nowYesterday.strftime('%Y')
    thisMonth= nowYesterday.strftime('%B')
    thisDay = nowYesterday.strftime('%d')
    # makeStructuredFolder
    ddPath = '%s/%s/%s/%s/%s' % (folderName,site,thisYear, thisMonth, thisDay)
    if not os.path.exists(ddPath):
        os.makedirs(ddPath)
        print('Folders created successfully')
    return ddPath
def rome(n):
    if n == '01':
        rom = 'I'
        return rom
    if n == '02':
        rom = 'II'
        return rom
    if n == '03':
        rom = 'III'
        return rom
    if n == '04':
        rom = 'IV'
        return rom
    if n == '05':
        rom = 'V'
        return rom
    if n == '06':
        rom = 'VI'
        return rom
    if n == '07':
        rom = 'VII'
        return rom
    if n == '08':
        rom = 'VIII'
        return rom
    if n == '09':
        rom = 'IX'
        return rom
    if n == '10':
        rom = 'X'
        return rom
    if n == '11':
        rom = 'XI'
        return rom
    if n == '12':
        rom = 'XII'
        return rom


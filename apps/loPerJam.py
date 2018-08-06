import sys
import pandas as pd
import numpy as np
import xlrd
import os
sys.path.insert(0,'../Date')
from allDateFormat import getNow
from emailThingy import testAddr,logIn,mailAttachment,eAddress
from email.mime.multipart import MIMEMultipart

def loOpen(thisHour,textPath6):
    # proses loOpen
    xlsx = pd.ExcelFile(textPath6)
    df = pd.read_excel(xlsx, 0)
    # convert xlsx date format into unix
    '''
    df['Delivery Date'] = pd.to_datetime('30-12-1899')+pd.to_timedelta(df['Delivery Date'],'D')
    df['Order Expired'] = pd.to_datetime('30-12-1899')+pd.to_timedelta(df['Order Expired'],'D')
    df['LastUpdateMs2'] = pd.to_datetime('30-12-1899')+pd.to_timedelta(df['LastUpdateMs2'],'D')
    df['Delivery Date'] = df['Delivery Date'].dt.strftime('%d-%m-%Y')
    df['Order Expired'] = df['Order Expired'].dt.strftime('%d-%m-%Y')
    df['LastUpdateMs2'] = df['LastUpdateMs2'].dt.strftime('%d-%m-%Y')'''
    # all status
    dfPlan = df[['Spbu No.', 'Ship To Code', 'OrderNumber', 'Delivery Number', 'Product', 'Volume', 'Delivery Date',
                 'Order Expired', 'Rit', 'MS2', 'Order Status']]
    # open status
    dfOpen = dfPlan[dfPlan['Order Status'] == 0]
    # pivot
    table = pd.pivot_table(dfOpen, index=['Product'], values=['Volume'], aggfunc=[np.sum, len], fill_value=0,
                           margins=True, margins_name='TOTAL')
    table2 = pd.pivot_table(dfOpen, index=['Product'], values=['Volume'], columns=['Rit'], aggfunc=np.sum, fill_value=0,
                            margins=True, margins_name='TOTAL')
    table = table.astype(int)
    table2 = table2.astype(int)
    # theWriter
    LoOpenWriter = pd.ExcelWriter(textPath6)
    # excel proses
    dfOpen.to_excel(LoOpenWriter, sheet_name='LO OPEN', index=False)
    table.to_excel(LoOpenWriter, sheet_name='Sum Products')
    table2.to_excel(LoOpenWriter, sheet_name='Sum Volume by Rit')
    # saveTheFile
    LoOpenWriter.save()
    print('\n\t--LO OPEN Pukul {0}.00--\n\n'.format(thisHour))
    print(table)
def dataLO(thisHour,textPath5):
    # proses dataLO
    xlFile = xlrd.open_workbook(textPath5)
    dataLO = xlFile.sheet_by_index(0)
    listPre = []
    listPx = []
    listPxT = []
    listDexl = []
    listPDex = []
    listPlt = []
    listSol = []
    listBio = []

    for i in range(3, dataLO.nrows):
        if dataLO.cell_value(i, 12) == 'PREMIUM':
            listPre.append(dataLO.cell_value(i, 14))
        elif dataLO.cell_value(i, 12) == 'BIOSOLAR':
            listBio.append(dataLO.cell_value(i, 14))
        elif dataLO.cell_value(i, 12) == 'SOLAR':
            listSol.append(dataLO.cell_value(i, 14))
        elif dataLO.cell_value(i, 12) == 'DEXLITE':
            listDexl.append(dataLO.cell_value(i, 14))
        elif dataLO.cell_value(i, 12) == 'PERTAMINA-DEX':
            listPDex.append(dataLO.cell_value(i, 14))
        elif dataLO.cell_value(i, 12) == 'PERTAMAX':
            listPx.append(dataLO.cell_value(i, 14))
        elif dataLO.cell_value(i, 12) == 'PERTAMAX-TURBO':
            listPxT.append(dataLO.cell_value(i, 14))
        elif dataLO.cell_value(i, 12) == 'PERTALITE':
            listPlt.append(dataLO.cell_value(i, 14))
    premium = sum(listPre)
    bio = sum(listBio)
    sol = sum(listSol)
    dexl = sum(listDexl)
    PDex = sum(listPDex)
    px = sum(listPx)
    pxt = sum(listPxT)
    plt = sum(listPlt)
    totalitas = premium + bio + sol + dexl + PDex + px + pxt + plt
    premium = format(int(premium), ',d')
    bio = format(int(bio), ',d')
    sol = format(int(sol), ',d')
    dexl = format(int(dexl), ',d')
    PDex = format(int(PDex), ',d')
    px = format(int(px), ',d')
    pxt = format(int(pxt), ',d')
    plt = format(int(plt), ',d')
    totalitas = format(int(totalitas), ',d')
    sumDataLo = '\n\n\t--Realisasi Pukul {9}.00--\n\nPREMIUM\t\t:\t{0}\nBIOSOLAR\t:\t{1}\nSOLAR\t\t:\t{2}\nDEXLITE\t\t:\t{3}\nPERTAMINA-DEX\t:\t{4}\nPERTAMAX\t:\t{5}\nPERTAMAX-TURBO\t:\t{6}\nPERTALITE\t:\t{7}\nTOTAL\t\t:\t{8}\n'.format(
        premium, bio, sol, dexl, PDex, px, pxt, plt, totalitas, thisHour)
    print(sumDataLo)
def content(textPath5,textPath6):
    today = getNow()
    dataLOFormat = today.strftime('%d%b%Y')
    thisHour = today.strftime('%H')
    namaFile = 'DATA_LO DAN LO_OPEN {0} Pukul {1}.00'.format(dataLOFormat, thisHour)
    msg = MIMEMultipart()
    msg['Subject'] = namaFile
    # xlsxAttachment
    # attach dataLO
    att1 = mailAttachment(textPath5,os.path.basename(textPath5))
    att2 = mailAttachment(textPath6,os.path.basename(textPath6))
    attachment = [att1, att2]
    [msg.attach(i) for i in attachment]
    return msg
def sendLO(textPath5,textPath6):
    today = getNow()
    thisHour = today.strftime('%H')
    #make data LO and LO open
    loOpen(thisHour,textPath6)
    dataLO(thisHour,textPath5)
    #mail thingy
    fromAddr,toAddr,pswd = eAddress('PLP_LO')
    msg = content(textPath5,textPath6)
    logIn(fromAddr, toAddr, pswd, msg)

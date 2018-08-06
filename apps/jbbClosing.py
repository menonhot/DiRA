import pandas as pd
import imapclient
import pyzmail
import os
import re
from email.mime.multipart import MIMEMultipart
import sys
sys.path.insert(0,'../Date')
from allDateFormat import getNow,getDashDate,getYesterday,getMidnight
from emailThingy import mailAttachment,logInRead,testAddr,logIn,eAddress
from structuredFolder import structFolder,structFolderMan

def loOpen(strObjd,textPath8,site):
    #proses loOpen
    xlsx2 = pd.ExcelFile(textPath8)
    df2 = pd.read_excel(xlsx2,0)
    strObjd = getMidnight(strObjd)
    #manipulate the dataframe
    dfPlan2 = df2[['Spbu No.','Ship To Code','OrderNumber','Delivery Number','Product','Volume','Delivery Date','Order Expired','Rit','MS2','Order Status','LastUpdateMs2Mpv']]
    dfOpen2 = dfPlan2[dfPlan2['Order Status'] == 0]
    dfMsT = dfOpen2[dfOpen2['LastUpdateMs2Mpv'] > strObjd]
    dfHMinus = dfOpen2[dfOpen2['LastUpdateMs2Mpv'] < strObjd]
    #saveIt
    loOpenHMinus = pd.ExcelWriter(structFolder('Report',site)+'/LO_OPEN_H-1.xlsx') #textPath
    loOpenPlan = pd.ExcelWriter(structFolder('Report',site)+'/PlanMs2.xlsx') #textpath
    dfHMinus.to_excel(loOpenHMinus,sheet_name='outstanding',index=False)
    dfMsT.to_excel(loOpenPlan,sheet_name='Plan Ms2')
    loOpenHMinus.save()
    loOpenPlan.save()
def getAttName(site):
    tgl = getDashDate()
    if site == 'UJB':
        hMin1F= 'OutstandingUJB %s.xlsx' % tgl
        ms2F= 'PlanMs2UJB %s.xlsx' % tgl
        dLO = 'DataLO_UJB %s.xls' % tgl
        return hMin1F,ms2F,dLO
    elif site == 'TGR':
        hMin1F= 'OutstandingTG %s.xlsx' % tgl
        ms2F= 'PlanMs2TG %s.xlsx' % tgl
        dLO = 'DataLO_TG %s.xls' % tgl
        return hMin1F,ms2F,dLO
    elif site == 'PLP':
        hMin1F= 'OutstandingPLP %s.xlsx' % tgl
        ms2F= 'PlanMs2PLP %s.xlsx' % tgl
        dLO = 'DataLO_PLP %s.xls' % tgl
        return hMin1F,ms2F,dLO
    elif site == 'BLG':
        dLO = 'DATA_LO_BLG %s.xls' % tgl
        return dLO
def getAttNameYtd(site):
    tgl = getYesterday()
    if site == 'UJB':
        hMin1F= 'OutstandingUJB %s.xlsx' % tgl.strftime('%d-%m-%Y')
        ms2F= 'PlanMs2UJB %s.xlsx' % tgl.strftime('%d-%m-%Y')
        dLO = 'DataLO_UJB %s.xls' % tgl.strftime('%d-%m-%Y')
        return hMin1F,ms2F,dLO
    elif site == 'TGR':
        hMin1F= 'OutstandingTG %s.xlsx' % tgl.strftime('%d-%m-%Y')
        ms2F= 'PlanMs2TG %s.xlsx' % tgl.strftime('%d-%m-%Y')
        dLO = 'DataLO_TG %s.xls' % tgl.strftime('%d-%m-%Y')
        return hMin1F,ms2F,dLO
    elif site == 'PLP':
        hMin1F= 'OutstandingPLP %s.xlsx' % tgl.strftime('%d-%m-%Y')
        ms2F= 'PlanMs2PLP %s.xlsx' % tgl.strftime('%d-%m-%Y')
        dLO = 'DataLO_PLP %s.xls' % tgl.strftime('%d-%m-%Y')
        return hMin1F,ms2F,dLO
    elif site == 'BLG':
        dLO = 'DATA_LO_BLG %s.xls' % tgl.strftime('%d-%m-%Y')
        return dLO
def mailContent(site,subs,dLOPath,cats):
    if site == 'BLG':
        msg = MIMEMultipart()
        msg['Subject'] = subs
        if cats == 'TODAY':
            dLOF = getAttName(site)
        elif cats == 'YESTERDAY':
            dLOF = getAttNameYtd(site)
        att = mailAttachment(dLOPath,dLOF)
        msg.attach(att)
        return msg
    else:
        msg = MIMEMultipart()
        msg['Subject'] = subs
        hMin1 = 'LO_OPEN_H-1.xlsx'
        ms2 = 'PlanMs2.xlsx'
        if cats == 'TODAY':
            hMin1F,ms2F,dLOF = getAttName(site)
        elif cats == 'YESTERDAY':
            hMin1F, ms2F, dLOF = getAttNameYtd(site)
        att1 = mailAttachment(hMin1,hMin1F)
        att2 = mailAttachment(ms2,ms2F)
        att3 = mailAttachment(dLOPath,dLOF)
        attachment = [att1,att2,att3]
        [msg.attach(i) for i in attachment]
        return msg
def mailContent2(subs,DLOPath,outPath,planPath):
    likeRightNow = getDashDate()
    msg = MIMEMultipart()
    msg['Subject'] = subs
    outName = 'Outstanding JBB %s.xlsx' % likeRightNow
    planName= 'Plan Ms2 JBB %s.xlsx' % likeRightNow
    DLOName = 'DataLO JBB %s.xls' % likeRightNow
    att1 = mailAttachment(outPath,outName)
    att2 = mailAttachment(planPath,planName)
    att3 = mailAttachment(DLOPath,DLOName)
    attachment = [att1,att2,att3]
    [msg.attach(i) for i in attachment]
    return msg
def fetchSite(rightNow):
    likeRightNow = rightNow
    imapObj = imapclient.IMAPClient('imap.gmail.com',ssl=True)
    imapObj.login('jbbsitesautomail@gmail.com','113333555555')
    imapObj.select_folder('INBOX',readonly=True)
    UIDs = imapObj.gmail_search('in:inbox subject:(JBB closing %s)' % likeRightNow)
    sender = {'ujb.automail@gmail.com': 'UJUNG BERUNG', 'automailbalongan@gmail.com': 'BALONGAN',
              'plumpang.automail.com@gmail.com': 'PLUMPANG', 'tg.automail@gmail.com': 'TANJUNG GEREM'}
    senders = ['BALONGAN', 'UJUNG BERUNG', 'PLUMPANG', 'TANJUNG GEREM']
    theSender = []
    theSenders = []
    senderStats = []
    for i in UIDs:
        rawMessages = imapObj.fetch([i], ['BODY[]', 'FLAGS'])
        messages = pyzmail.PyzMessage.factory(rawMessages[i][b'BODY[]'])
        theSender.append(messages.get_addresses('from'))
    [theSenders.append(i[0][0]) for i in theSender]
    [senderStats.append(sender[i]) for i in theSenders]
    notif = [i + ' not ready' for i in senders if i not in senderStats]
    return notif
def createFolder(likeRightNow):
    # create folder
    jbbFolder = os.environ['USERPROFILE']+'/Report/jbbReport'
    blgPath = 'blgJBB'
    blgPath = structFolderMan(jbbFolder,blgPath,likeRightNow)
    ujbPath = 'ujbJBB'
    ujbPath = structFolderMan(jbbFolder,ujbPath,likeRightNow)
    tgPath = 'tgJBB'
    tgPath = structFolderMan(jbbFolder,tgPath,likeRightNow)
    plpPath = 'plpJBB'
    plpPath = structFolderMan(jbbFolder,plpPath,likeRightNow)
    jbbPath = 'jbbUnion'
    jbbPath = structFolderMan(jbbFolder,jbbPath,likeRightNow)
    return blgPath,ujbPath,tgPath,plpPath,jbbPath
def downMail(imapObj,likeRightNow):
    blgPath, ujbPath, tgPath, plpPath, jbbPath = createFolder(likeRightNow)
    UIDs = imapObj.gmail_search('in:inbox subject:(JBB closing %s)' % likeRightNow)
    blgDlo = re.compile('DATA_LO_BLG.*')
    tgDlo = re.compile('DataLO_TG.*')
    tgHmin = re.compile('OutstandingTG.*')
    tgPlan = re.compile('PlanMs2TG.*')
    ujbDlo = re.compile('DataLO_UJB.*')
    ujbHmin = re.compile('OutstandingUJB.*')
    ujbPlan = re.compile('PlanMs2UJB.*')
    plpDlo = re.compile('DataLO_PLP.*')
    plpHmin = re.compile('OutstandingPLP.*')
    plpPlan = re.compile('PlanMs2PLP.*')
    # download mail
    for i in UIDs:
        rawMessages = imapObj.fetch([i], ['BODY[]', 'FLAGS'])
        messages = pyzmail.PyzMessage.factory(rawMessages[i][b'BODY[]'])
        if messages.get_content_maintype() == 'multipart':
            for part in messages.walk():
                if part.get_content_maintype() == 'multipart': continue
                if part.get('Content-Disposition') is None: continue
                filename = part.get_filename()
                if blgDlo.match(filename):
                    # print(filename)
                    sv_path = os.path.join(blgPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                elif tgDlo.match(filename):
                    # print(filename)
                    sv_path = os.path.join(tgPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                elif tgHmin.match(filename):
                    # print(filename)
                    sv_path = os.path.join(tgPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                elif tgPlan.match(filename):
                    # print(filename)
                    sv_path = os.path.join(tgPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                elif ujbDlo.match(filename):
                    # print(filename)
                    sv_path = os.path.join(ujbPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                elif ujbHmin.match(filename):
                    # print(filename)
                    sv_path = os.path.join(ujbPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                elif ujbPlan.match(filename):
                    # print(filename)
                    sv_path = os.path.join(ujbPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                elif plpDlo.match(filename):
                    # print(filename)
                    sv_path = os.path.join(plpPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                elif plpHmin.match(filename):
                    # print(filename)
                    sv_path = os.path.join(plpPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                elif plpPlan.match(filename):
                    # print(filename)
                    sv_path = os.path.join(plpPath, filename)
                    if not os.path.isfile(sv_path):
                        print(sv_path)
                        fp = open(sv_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
    #filePath
    blgDir = blgPath+'/DATA_LO_BLG %s.xls' % likeRightNow
    ujbDirDLO = ujbPath+'/DataLO_UJB %s.xls' % likeRightNow
    ujbDirHminus = ujbPath+'/OutstandingUJB %s.xlsx' % likeRightNow
    ujbDirPlan = ujbPath+'/PlanMs2UJB %s.xlsx' % likeRightNow
    tgDirDLO = tgPath+'/DataLO_TG %s.xls' % likeRightNow
    tgDirHminus = tgPath+'/OutstandingTG %s.xlsx' % likeRightNow
    tgDirPlan = tgPath+'/PlanMs2TG %s.xlsx' % likeRightNow
    plpDirDLO = plpPath+'/DataLO_PLP %s.xls' % likeRightNow
    plpDirHminus = plpPath+'/OutstandingPLP %s.xlsx' % likeRightNow
    plpDirPlan = plpPath+'/PlanMs2PLP %s.xlsx' % likeRightNow
    #writer
    DLOPath = jbbPath+'/DataLO JBB %s.xlsx' % likeRightNow
    outPath = jbbPath+'/Outstanding JBB %s.xlsx' % likeRightNow
    planPath= jbbPath+'/Plan Ms2 JBB %s.xlsx' % likeRightNow
    writerDLO = pd.ExcelWriter(DLOPath, engine='xlsxwriter')
    writerOut = pd.ExcelWriter(outPath, engine='xlsxwriter')
    writerPlan = pd.ExcelWriter(planPath, engine='xlsxwriter')
    #readDataLO
    dataLOBlg = pd.read_excel(blgDir, sheet_name='Sheet')
    dataLOUjb = pd.read_excel(ujbDirDLO, sheet_name='Sheet')
    dataLOTg = pd.read_excel(tgDirDLO, sheet_name='Sheet')
    dataLOPlp = pd.read_excel(plpDirDLO, sheet_name='Sheet')
    #writeDataLO
    dataLOBlg.to_excel(writerDLO, sheet_name='BLG', index=False)
    dataLOUjb.to_excel(writerDLO, sheet_name='UJB', index=False)
    dataLOTg.to_excel(writerDLO, sheet_name='TG', index=False)
    dataLOPlp.to_excel(writerDLO, sheet_name='PLP', index=False)
    #readOS
    outstandingUjb = pd.read_excel(ujbDirHminus, sheet_name='outstanding')
    outstandingTg = pd.read_excel(tgDirHminus, sheet_name='outstanding')
    outstandingPlp = pd.read_excel(plpDirHminus, sheet_name='outstanding')
    #writeOS
    outstandingUjb.to_excel(writerOut, sheet_name='UJB', index=False)
    outstandingTg.to_excel(writerOut, sheet_name='TG', index=False)
    outstandingPlp.to_excel(writerOut, sheet_name='PLP', index=False)
    #readPlan
    planUjb = pd.read_excel(ujbDirPlan, sheet_name='Plan Ms2')
    planTg = pd.read_excel(tgDirPlan, sheet_name='Plan Ms2')
    planPlp = pd.read_excel(plpDirPlan, sheet_name='Plan Ms2')
    #writePlan
    planUjb.to_excel(writerPlan, sheet_name='UJB', index=False)
    planTg.to_excel(writerPlan, sheet_name='TG', index=False)
    planPlp.to_excel(writerPlan, sheet_name='PLP', index=False)
    #saveTheShit
    writerDLO.save()
    writerOut.save()
    writerPlan.save()
    return DLOPath,outPath,planPath
def jbbReport(likeRightNow):
    fromAddr, toAddr, pswd = eAddress('JBB')
    imapObj = logInRead(fromAddr,pswd)
    DLOPath,outPath,planPath = downMail(imapObj,likeRightNow)
    subs = 'JBB Closing %s' % likeRightNow
    msg = mailContent2(subs,DLOPath,outPath,planPath)
    logIn(fromAddr,toAddr,pswd,msg)

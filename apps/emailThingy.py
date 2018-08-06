import smtplib
from email.mime.base import MIMEBase
from email import encoders
from jbbClosing import *
def logIn(fromAddr,toAddr,pswd,msg):
    mailer = smtplib.SMTP('smtp.gmail.com', 587, timeout=120)
    mailer.starttls()
    mailer.ehlo()
    mailer.login(fromAddr, pswd)
    mailer.sendmail(fromAddr,toAddr,msg.as_string())
    mailer.quit()
    print('email sent')
def eAddress(site):
    if site == 'BLG':
        fromAddr = 'automailbalongan@gmail.com'
        toAddr = 'jbbsitesautomail@gmail.com', 'balongan.daily.report@prolindo.net'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'UJB':
        fromAddr = 'ujb.automail@gmail.com'
        toAddr = 'jbbsitesautomail@gmail.com', 'ujungberung.daily.report@cpan.biz'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'TGR':
        fromAddr = 'tg.automail@gmail.com'
        toAddr = 'jbbsitesautomail@gmail.com', 'gerem.daily.report@nusatama.net'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'PLP':
        fromAddr = 'plumpang.automail.com@gmail.com'
        toAddr = 'jbbsitesautomail@gmail.com','plumpang.daily.report@cpan.biz'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'JBB':
        fromAddr = 'jbbsitesautomail@gmail.com'
        toAddr = 'evi@cpan.biz','vans@cpan.biz','plumpang.daily.report@cpan.biz'
        pswd = '113333555555'
        return fromAddr,toAddr,pswd
    elif site == 'PLP_LO':
        fromAddr = 'plumpang.automail.com@gmail.com'
        toAddr = 'evi@cpan.biz', 'syahrudin@cpan.biz', 'vans@cpan.biz', 'adm.ppp.plp2@gmail.com', 'setiadi@pertamina.com','rolin_marlin@pertamina.com','adm.ppp.plp@mitrakerja.pertamina.com','afri@pertamina.com', 'ahmad.ananta@pertamina.com', 'mk.anas.mustari@mitrakerja.pertamina.com', 'andi_indrawan@pertamina.com','bagus.abimanyu@patraniaga.com', 'bayusuryo@pertamina.com', 'budipras@pertamina.com','controlroom_plumpang@pertamina.com', 'diman.santosa@pertamina.com','faizal.pujayanto@pertamina.com','gantryplumpang@gmail.com', 'gantryplumpang@yahoo.com', 'gemaip@pertamina.com', 'mk.hari.burhan@mitrakerja.pertamina.com', 't.d.kristanto@pertamina.com','layananjualplumpang@pertamina.com','allyf_machine04@yahoo.com', 'muky.agil@pertamina.com', 'munandar_taufiq@yahoo.com','m.billah73@yahoo.com', 'penyaluran_plumpang@pertamina.com', 'prestateknik@yahoo.com','mk.rucky@mitrakerja.pertamina.com', 'j_tarwoco@pertamina.com', 'toni.pradana@pertamina.com', 'dwi.jarwanto@pertamina.com ','plumpang.daily.report@cpan.biz'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'PLP_NOTA':
        fromAddr = 'plumpang.automail.com@gmail.com'
        toAddr = 'evi@cpan.biz'
        pswd = 'prisela99'
        return fromAddr, toAddr, pswd
    elif site == 'BLG_NOTA':
        fromAddr = 'automailbalongan@gmail.com'
        toAddr = 'erwin@cpan.biz'
        pswd = 'prisela99'
        return fromAddr, toAddr, pswd
    elif site == 'UJB_NOTA':
        fromAddr = 'ujb.automail@gmail.com'
        toAddr = 'linda@cpan.biz'
        pswd = 'prisela99'
        return fromAddr, toAddr, pswd
    elif site == 'TGR_NOTA':
        fromAddr = 'tg.automail@gmail.com'
        toAddr = 'sugeng@nusatama.net'
        pswd = 'prisela99'
        return fromAddr, toAddr, pswd
    elif site == 'BYL_NOTA':
        fromAddr = 'boyolali.automail@gmail.com'
        toAddr = 'suci@nusatama.net'
        pswd = 'boyolali123'
        return fromAddr, toAddr, pswd
    elif site == 'MDN_NOTA':
        fromAddr = 'labuandeli.automail@gmail.com'
        toAddr = 'mujib@prolindo.net'
        pswd = 'labuandeli123'
        return fromAddr, toAddr, pswd
    elif site == 'SBY_NOTA':
        fromAddr = 'perak.automail@gmail.com'
        toAddr = 'anita@prolindo.net'
        pswd = 'perak123'
    elif site == 'KTP_NOTA':
        fromAddr = 'kertapati.automail@gmail.com '
        toAddr = 'farida@cpan.biz'
        pswd = 'kertapati123'
    elif site == 'PJG_NOTA':
        fromAddr = 'panjang.automail@gmail.com'
        toAddr = 'agusman@rekavisi.net'
        pswd = 'panjang123'
def testAddr(site):
    if site == 'BLG':
        fromAddr = 'automailbalongan@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'UJB':
        fromAddr = 'ujb.automail@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'TGR':
        fromAddr = 'tg.automail@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'PLP':
        fromAddr = 'plumpang.automail.com@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'JBB':
        fromAddr = 'jbbsitesautomail@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = '113333555555'
        return fromAddr,toAddr,pswd
    elif site == 'PLP_LO':
        fromAddr = 'plumpang.automail.com@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'PLP_NOTA':
        fromAddr = 'plumpang.automail.com@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'prisela99'
        return fromAddr,toAddr,pswd
    elif site == 'BLG_NOTA':
        fromAddr = 'automailbalongan@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'prisela99'
        return fromAddr, toAddr, pswd
    elif site == 'UJB_NOTA':
        fromAddr = 'ujb.automail@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'prisela99'
        return fromAddr, toAddr, pswd
    elif site == 'TGR_NOTA':
        fromAddr = 'tg.automail@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'prisela99'
        return fromAddr, toAddr, pswd
    elif site == 'BYL_NOTA':
        fromAddr = 'boyolali.automail@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'boyolali123'
        return fromAddr, toAddr, pswd
    elif site == 'MDN_NOTA':
        fromAddr = 'labuandeli.automail@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'labuandeli123'
        return fromAddr, toAddr, pswd
    elif site == 'SBY_NOTA':
        fromAddr = 'perak.automail@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'perak123'
    elif site == 'KTP_NOTA':
        fromAddr = 'kertapati.automail@gmail.com '
        toAddr = 'menonhot@gmail.com'
        pswd = 'kertapati123'
    elif site == 'PJG_NOTA':
        fromAddr = 'panjang.automail@gmail.com'
        toAddr = 'menonhot@gmail.com'
        pswd = 'panjang123'
def mailAttachment(pathFile,fileName):
    att = MIMEBase('application','octet-stream')
    with open(pathFile, 'rb') as file: att.set_payload(file.read())
    encoders.encode_base64(att)
    att.add_header('Content-Disposition','attachment',filename=fileName)
    return att
def logInRead(usr,pwd):
    # read email
    imapObj = imapclient.IMAPClient('imap.gmail.com', ssl=True)
    imapObj.login(usr, pwd)
    imapObj.select_folder('INBOX', readonly=True)
    return imapObj

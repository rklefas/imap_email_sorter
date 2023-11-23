import imaplib
from datetime import datetime
from dateutil.parser import *
import json
import winsound
from win32com.client import Dispatch
from inputimeout import inputimeout, TimeoutOccurred

# ---------------
# ---------------

def do_log(message):
    dateX = datetime.now().strftime("%Y-%m-%d")
    timeX = datetime.now().strftime(" %H:%M:%S")
    file1 = open('logs/' + dateX + "-actions.log", "a")
    file1.write(dateX + timeX + " " + message + "\n")
    file1.close()

# ---------------

def getyear(dateheader):
    lessdate = str(dateheader.strip())
    lessdate = lessdate.replace('Date: ', '')
    lessdate = parse(lessdate)
    return lessdate.strftime('%Y')

# ---------------

def getfromname(fromhead):
    frommer = str(fromhead.strip())
    frommer = frommer[0:frommer.index('<')]
    frommer = frommer.replace('From: ', '')
    frommer = frommer.replace('"', '')
    frommer = frommer.strip()
    return frommer

# ---------------

def getfromemail(fromhead):
    frommer = str(fromhead.strip())
    frommer = frommer[frommer.index('<'):-1]
    frommer = frommer.strip('<')
    return frommer

# ---------------

def rearrangefrom(frommer):
    temp = frommer.replace('@', '.')
    parts = temp.split('.')
    parts.reverse()
    temp = '.'.join(parts)
    return temp

# ---------------

def determinefolder(num):

    typ, fromX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (FROM)])')
    typ, dateX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (DATE)])')
    
    FOLDERSTACK = ["PYTHON-SORT"]
    FOLDERSTACK.append(rearrangefrom(getfromemail(fromX[0][1].decode())))
    FOLDERSTACK.append(getfromname(fromX[0][1].decode()))
    FOLDERSTACK.append(getyear(dateX[0][1].decode()))
    
    return FOLDERSTACK

# ---------------

def spokeninput(q):
    Dispatch("SAPI.SpVoice").Speak(q)
    return input(q)
    
# ---------------
    
def spokeninputtimeout(q, default):
    Dispatch("SAPI.SpVoice").Speak(q)
    
    try:
        return inputimeout(q + ' (default : ' + default + ')', 15)
    except TimeoutOccurred:
        Dispatch("SAPI.SpVoice").Speak('Using default value '+default)

        return default

# ---------------

def speakline(key, val):
    println(key, val)
    do_log(key + ' ' + val)
    Dispatch("SAPI.SpVoice").Speak(key + ' ' + val)

# ---------------

def createfolder(num):

    FOLDERSTACK = determinefolder(num)
    ROOTFOLDER = '/'.join(FOLDERSTACK)
    
    pack = ''

    print('Check for folder:', ROOTFOLDER)
        
    if spokeninputtimeout('  Do you want to create this folder? ', 'y').lower().strip() == 'y':
        for FOLDER in FOLDERSTACK:
            pack = pack + '/' + FOLDER
            pack = pack.strip('/')
            
            imap_ssl.create('"'+pack+'"')
    else:
        return ['AUTOREVIEW']
        
    return FOLDERSTACK
    
# ---------------

def moveemail(FOLDERSTACK, num):

    ROOTFOLDER = '/'.join(FOLDERSTACK)
    
    print('Moving email to:', ROOTFOLDER)
    
#    return True

    imap_ssl.uid('COPY', num, '"'+ROOTFOLDER+'"')
    print('  Copied', num)
    
    imap_ssl.uid('STORE', num, '+FLAGS', '\\Deleted')
    print('  Deleted', num)
    
    return True

# ---------------

def println(key, value):
    timeX = datetime.now().strftime("%H:%M:%S ")
    print(timeX, key, '           ', value)

# ---------------

def showemail(num):

    print('\nMessage # %s' % (num))

    typ, fromX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (FROM)])')
    typ, dateX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (DATE)])')
    typ, subjectX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (SUBJECT)])')
    
    print(dateX[0][1].decode().strip(), '   ', fromX[0][1].decode().strip())
    print(subjectX[0][1].decode().strip())

# ---------------
# ---------------

configs = json.load(open('./config.json', 'r'))

################ IMAP SSL ##############################

with imaplib.IMAP4_SSL(host=configs['host'], port=imaplib.IMAP4_SSL_PORT) as imap_ssl:

    ############### Login to Mailbox ######################
    
    println("Logging into mailbox:   ", configs['host'])
    resp_code, response = imap_ssl.login(configs['user'], configs['pass'])

    println("Login Result:           ", str(resp_code))

    #################### List Emails #####################
    
    resp_code, mail_count = imap_ssl.select(mailbox='"INBOX"')
    speakline("Inbox Count:", str(mail_count[0].decode()))

 
    
    while True:
    
        for num in range(1, 7):
            showemail(str(num))
        
        print("")
        print("")
        
        peekEmail = spokeninputtimeout('WHICH emails to sort? ', '1')
        
        if (peekEmail == ''):
            break
        
        typ, fromX = imap_ssl.fetch(peekEmail, '(RFC822.SIZE BODY[HEADER.FIELDS (FROM)])')
        typ, dateX = imap_ssl.fetch(peekEmail, '(RFC822.SIZE BODY[HEADER.FIELDS (DATE)])')
        yearX = getyear(dateX[0][1].decode())

        searchString = '(SINCE "01-Jan-'+yearX+'" BEFORE "31-Dec-'+yearX+'" FROM "'+getfromname(fromX[0][1].decode().strip())+' '+getfromemail(fromX[0][1].decode().strip())+'")'

        typ, data = imap_ssl.uid('search', None, searchString)
        
        speakline("Query", searchString)
        speakline("Emails searched and found:", str(len(data[0].decode().split())))
        
        FOLDERSTACK = determinefolder(peekEmail)
        createfolder(peekEmail)
        
        print(data)
        
        for this_uid in data[0].decode().split():
        
            try:
            
                moveemail(FOLDERSTACK, this_uid)

            except Exception as e:
            
                print('  Failed to move email!')
                print(e)


        speakline("Emails sorted:", str(len(data[0].decode().split())))

    ############# Close Selected Mailbox #######################
    imap_ssl.close()
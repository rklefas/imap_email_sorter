import imaplib
from datetime import datetime
from dateutil.parser import *
import json


def getyear(dateheader):
    lessdate = str(dateheader.strip())
    lessdate = lessdate.replace('Date: ', '')
    lessdate = parse(lessdate)
    return lessdate.strftime('%Y')



def getfrom(fromhead):
    frommer = str(fromhead.strip())
    frommer = frommer[frommer.index('<'):-1]
    frommer = frommer.strip('<')
    return frommer


def rearrangefrom(frommer):
    temp = frommer.replace('@', '.')
    parts = temp.split('.')
    parts.reverse()
    temp = '.'.join(parts)
    return temp



def determinefolder(num):

    typ, fromX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (FROM)])')
    typ, dateX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (DATE)])')
    
    FOLDERSTACK = ["AUTOSORT"]
    FOLDERSTACK.append(rearrangefrom(getfrom(fromX[0][1].decode())))
    FOLDERSTACK.append(getyear(dateX[0][1].decode()))
    
    return FOLDERSTACK


def createfolder(num):

    FOLDERSTACK = determinefolder(num)
    
    pack = ''

    for FOLDER in FOLDERSTACK:
        pack = pack + '/' + FOLDER
        pack = pack.strip('/')
        
        print('Check for folder:', pack)
        
        if input('  Do you want to create this folder? ').lower().strip() == 'y':
            imap_ssl.create('"'+pack+'"')
        else:
            return ['AUTOREVIEW']
        
    return FOLDERSTACK
    



def moveemail(FOLDERSTACK, num):

    return True

    ROOTFOLDER = '/'.join(FOLDERSTACK)
    
    print('Moving email to:', ROOTFOLDER)
    
    imap_ssl.copy(num, ROOTFOLDER)
    imap_ssl.store(num, '+FLAGS', '\\Deleted')

    return True


def showemail(num):
    typ, fromX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (FROM)])')
    print(fromX[0][1].decode().strip())
    
    typ, dateX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (DATE)])')
    print(dateX[0][1].decode().strip())
    
    typ, subjectX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (SUBJECT)])')
    print(subjectX[0][1].decode().strip())



configs = json.load(open('./config.json', 'r'))

################ IMAP SSL ##############################

with imaplib.IMAP4_SSL(host=configs['host'], port=imaplib.IMAP4_SSL_PORT) as imap_ssl:

    ############### Login to Mailbox ######################
    
    print("Logging into mailbox:   ", configs['host'])
    resp_code, response = imap_ssl.login(configs['user'], configs['pass'])

    print("Login Result:            {}".format(resp_code))

    #################### List Emails #####################
    
    resp_code, mail_count = imap_ssl.select(mailbox='"INBOX"')
    print("Fetch Inbox Count:       {}".format(mail_count[0].decode()))
    
    
    peekEmail = input('How many emails to sort? ')
    
    
    
    showemail(peekEmail)
    
    typ, fromX = imap_ssl.fetch(peekEmail, '(RFC822.SIZE BODY[HEADER.FIELDS (FROM)])')
    print(fromX[0][1].decode().strip())
    
    typ, dateX = imap_ssl.fetch(peekEmail, '(RFC822.SIZE BODY[HEADER.FIELDS (DATE)])')
    yearX = getyear(dateX[0][1].decode())


    typ, data = imap_ssl.search(None, '(SINCE "01-Jan-'+yearX+'" BEFORE "31-Dec-'+yearX+'" FROM "'+getfrom(fromX[0][1].decode().strip())+'")')
    
    print("Emails searched and found:       {}".format(len(data[0].decode().split())))

    
    for num in data[0].decode().split():
    
        
        print('\nMessage # %s' % (num))

        showemail(num)
        
        FOLDERSTACK = determinefolder(num)

       
        try:
        
            moveemail(FOLDERSTACK, num)
    
        except Exception as e:
        
            print('  Failed to move email!')
            createfolder(num)
            moveemail(FOLDERSTACK, num)

            
    ############# Close Selected Mailbox #######################
    imap_ssl.close()
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


def createfolder(FOLDERS):

    pack = ''

    for FOLDER in FOLDERS:
        pack = pack + '/' + FOLDER
        pack = pack.strip('/')
        
        print('Check for folder:', pack)
        
        try:
            
            resp_code, mail_count = imap_ssl.select(mailbox='"'+pack+'"')
            
            if int(mail_count[0].decode()) >= 0:
                print("Fetch Mail Count:        {}".format(mail_count[0].decode()))

        except Exception as e:
        
            if input('Do you want to create this folder? ') == 'y':
                imap_ssl.create('"'+pack+'"')
            else:
                return False
        
    resp_code, mail_count = imap_ssl.select(mailbox='"INBOX"')

    return True


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
    
    typ, data = imap_ssl.search(None, 'ALL')
    for num in data[0].decode().split():
        
        print('Message # %s\n' % (num))

        typ, fromX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (FROM)])')
        print(fromX[0][1].decode().strip())
        
        typ, dateX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (DATE)])')
        print(dateX[0][1].decode().strip())
        
        typ, subjectX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (SUBJECT)])')
        print(subjectX[0][1].decode().strip())
        
        FOLDERSTACK = ["TESTINGBOX"]
        FOLDERSTACK.append(rearrangefrom(getfrom(fromX[0][1].decode())))
        FOLDERSTACK.append(getyear(dateX[0][1].decode()))
                
        if createfolder(FOLDERSTACK):
            imap_ssl.copy(num, '/'.join(FOLDERSTACK))
            imap_ssl.store(num, '+FLAGS', '\\Deleted')
            imap_ssl.expunge()
        else:
            print('Skipping email')

        break
    ############# Close Selected Mailbox #######################
    imap_ssl.close()
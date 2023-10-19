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





configs = json.load(open('./config.json', 'r'))

v_host = configs['host']
v_user = configs['user']
v_pass = configs['pass']

################ IMAP SSL ##############################

with imaplib.IMAP4_SSL(host=v_host, port=imaplib.IMAP4_SSL_PORT) as imap_ssl:

    ############### Login to Mailbox ######################
    print("Logging into mailbox:   ", v_host)
    resp_code, response = imap_ssl.login(v_user, v_pass)

    print("Login Result:            {}".format(resp_code))
#    print("Response:                {}".format(response[0].decode()))

    #################### List Directores #####################
    resp_code, directories = imap_ssl.list()

    print("Fetch Folder List:       {}".format(resp_code))

    resp_code, mail_count = imap_ssl.select(mailbox='"INBOX"', readonly=True)
    typ, data = imap_ssl.search(None, 'ALL')
    for num in data[0].split():
        
        print('Message \n\n%s\n' % (num))

        typ, fromX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (FROM)])')
        print(fromX[0][1].strip())
        
        typ, dateX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (DATE)])')
        print(dateX[0][1].strip())
        
        typ, subjectX = imap_ssl.fetch(num, '(RFC822.SIZE BODY[HEADER.FIELDS (SUBJECT)])')
        print(subjectX[0][1].strip())
        
        FOLDER = 'TESTINGBOX/' + getfrom(fromX[0][1].decode()) + '/' + getyear(dateX[0][1].decode())
        
        print('Will create folder:', FOLDER)
        
        imap_ssl.create('"'+FOLDER+'"')


        break
    ############# Close Selected Mailbox #######################
    imap_ssl.close()
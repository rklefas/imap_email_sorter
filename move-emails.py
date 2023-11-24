import imap_tools
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

def show_message(msg):
    print(msg.uid, msg.date, msg.from_values.name, msg.from_values.email, len(msg.text or msg.html))

# ---------------

def rearrangefrom(frommer):
    temp = frommer.replace('@', '.')
    parts = temp.split('.')
    parts.reverse()
    temp = '.'.join(parts)
    return temp

# ---------------

def determinefolder(msg):

    FOLDERSTACK = ["TEST-SORT"]
    FOLDERSTACK.append(rearrangefrom(msg.from_values.email))
    FOLDERSTACK.append(msg.from_values.name)
    FOLDERSTACK.append(msg.date.strftime('%Y'))
    
    return FOLDERSTACK

# ---------------

def spokeninput(q):
    Dispatch("SAPI.SpVoice").Speak(q)
    return input(q)
    
# ---------------
    
def spokeninputtimeout(q, default):
    Dispatch("SAPI.SpVoice").Speak(q)
    
    try:
        return inputimeout(q + ' (default : ' + default + ') ', 10)
    except TimeoutOccurred:
        Dispatch("SAPI.SpVoice").Speak('Using default value '+default)

        return default

# ---------------

def speakline(key, val):
    println(key, val)
    do_log(key + ' ' + val)
    Dispatch("SAPI.SpVoice").Speak(key + ' ' + val)

# ---------------

def createfolder(FOLDERSTACK, mailbox):

    ROOTFOLDER = '/'.join(FOLDERSTACK)
    
    pack = ''

    println('Check for folder:', ROOTFOLDER)
        
    for FOLDER in FOLDERSTACK:
        pack = pack + '/' + FOLDER
        pack = pack.strip('/')
            
        if mailbox.folder.exists(pack) == False:
        
            println('  Does not exist:', pack)

            if spokeninputtimeout('  Do you want to create this folder? ', 'y').lower().strip() == 'y':
                mailbox.folder.create(pack)
            else:
                return ['AUTOREVIEW']
        
    return FOLDERSTACK
    
# ---------------

def println(key, value):
    timeX = datetime.now().strftime("%H:%M:%S ")
    print(timeX, key, '           ', value)

# ---------------
# ---------------

configs = json.load(open('./config.json', 'r'))

################ IMAP SSL ##############################

with imap_tools.MailBox(configs['host']).login(configs['user'], configs['pass']) as server:

    ############### Login to Mailbox ######################
    
    println("Logging into mailbox:   ", configs['host'])

    #################### List Emails #####################
    
    stat = server.folder.status('INBOX')
    print(stat)
 
    
    while True:
        preview = list(server.fetch(limit=7))
        
        for msg in preview:
            show_message(msg)
    
        
        print("")
        print("")
        
        peekEmail = spokeninputtimeout('WHICH emails to sort? ', '1')
        
        if (peekEmail == ''):
            break
        
        selectedEmail = preview[int(peekEmail)]
        
        fromX = selectedEmail.from_values
        yearX = selectedEmail.date.strftime('%Y')
        FOLDERSTACK = determinefolder(selectedEmail)
        createfolder(FOLDERSTACK, server)

        searchString = 'FROM "'+fromX.name+' '+fromX.email+'" SINCE "01-Jan-'+yearX+'" BEFORE "31-Dec-'+yearX+'"'
        speakline("Query", searchString)

        results = list(server.fetch(searchString))
        counting = 0
        ROOTFOLDER = '/'.join(FOLDERSTACK)
        
        speakline("   Result Count", str(len(results)))
        println('  Moving emails to:', ROOTFOLDER)
        

        for msg in results:
            show_message(msg)
        
            try:
                
                server.move(msg.uid, ROOTFOLDER)
                
                counting = counting + 1

            except Exception as e:
            
                print('  Failed to move email!')
                print(e)


        speakline("Emails sorted:", str(counting))


import imap_tools
from datetime import datetime
import json
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

def show_message(index, msg):
    print('Email', index, ' | ', msg.from_values.name, '  ', msg.from_values.email, '  ', msg.date)
    print('         |           ', msg.subject, msg.uid, len(msg.text or msg.html))

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
        return inputimeout(q + ' (default : ' + default + ') ', 30)
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

    FULLPATH = '/'.join(FOLDERSTACK)
    pack = ''

    println('Check for folder', FULLPATH)

    if mailbox.folder.exists(FULLPATH) == False:
        if spokeninputtimeout('  Do you want to create this folder? ', 'y').lower().strip() == 'y':
        
            for FOLDER in FOLDERSTACK:
                pack = pack + '/' + FOLDER
                pack = pack.strip('/')
                    
                if mailbox.folder.exists(pack) == False:
                
                    println('  Creating folder', pack)
                    mailbox.folder.create(pack)
        else:
            return createfolder(['AUTOREVIEW'])
        
    return FOLDERSTACK
    
# ---------------

def println(key, value):
    timeX = datetime.now().strftime("%H:%M:%S ")
    print(timeX, key, ':           ', value)

# ---------------
# ---------------

configs = json.load(open('./config.json', 'r'))

############### Login to Mailbox ######################

with imap_tools.MailBox(configs['host']).login(configs['user'], configs['pass']) as server:

    println("Logged into mailbox", configs['host'])

    #################### List Emails #####################
    
    stat = server.folder.status('INBOX')
    print(stat)

    
    while True:
        preview = list(server.fetch(limit=7))

        print("")
        print("")
        
        for index, msg in enumerate(preview):
            show_message(index, msg)
    
        print("")
        print("")
        
        peekEmail = spokeninputtimeout('Pick an email to sort. ', '0')
        
        if (peekEmail == ''):
            break
        
        selectedEmail = preview[int(peekEmail)]
        
        fromX = selectedEmail.from_values
        yearX = selectedEmail.date.strftime('%Y')

        searchString = 'FROM "'+fromX.name+' '+fromX.email+'" SINCE "01-Jan-'+yearX+'" BEFORE "31-Dec-'+yearX+'"'
        speakline("Query", searchString)

        results = list(server.fetch(searchString))
        FOLDERSTACK = determinefolder(selectedEmail)
        
        createfolder(FOLDERSTACK, server)
        
        counting = 0
        FULLPATH = '/'.join(FOLDERSTACK)
        
        speakline("  Result Count", str(len(results)))
        println('  Moving emails to', FULLPATH)
        

        for index, msg in enumerate(results):
            show_message(index, msg)
        
            try:
                
                server.move(msg.uid, FULLPATH)
                
                counting = counting + 1

            except Exception as e:
            
                print('  Failed to move email!')
                print(e)


        speakline("Emails sorted:", str(counting))


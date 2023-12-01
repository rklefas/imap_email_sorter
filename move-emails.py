import imap_tools
from datetime import datetime
import json
from win32com.client import Dispatch
from inputimeout import inputimeout, TimeoutOccurred
import unidecode

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
    print('         |           ', msg.subject[0:50], msg.uid, len(msg.text or msg.html))

# ---------------

def rearrangefrom(frommer):
    temp = frommer.replace('@', '.').lower()
    parts = temp.split('.')
    parts.reverse()
    temp = '.'.join(parts)
    return temp

# ---------------

def determinefolder(msg):

    FOLDERSTACK = ["PYTHON-SORT"]
    
    FOLDERSTACK.append(rearrangefrom(msg.from_values.email))
    FOLDERSTACK.append(unidecode.unidecode(msg.from_values.name).strip())
    FOLDERSTACK.append(msg.date.strftime('%Y'))
    
    return FOLDERSTACK

# ---------------

def spokeninput(q):
    Dispatch("SAPI.SpVoice").Speak(q)
    return input(q)
    
# ---------------
    
def spokeninputtimeout(q, default):
    Dispatch("SAPI.SpVoice").Speak(q)
    
    global dynamic_timeout

    try:
        val = inputimeout(q + ' (' + str(dynamic_timeout) + ' sec timeout, default : ' + default + ') ', dynamic_timeout)
        dynamic_timeout = max_timeout
        return val
    except TimeoutOccurred:
        Dispatch("SAPI.SpVoice").Speak('Defaulted to '+default)
        dynamic_timeout = max(min_timeout, int(dynamic_timeout / 2))
        return default

# ---------------

def speakline(key, val):
    println(key, val)
    Dispatch("SAPI.SpVoice").Speak(key + ' ' + val)

# ---------------

def createfolder(FOLDERSTACK, mailbox, count = None):

    FULLPATH = '/'.join(FOLDERSTACK)
    pack = ''

    println('Check for folder', FULLPATH)

    if mailbox.folder.exists(FULLPATH) == False:
    
        if count == 0:
            return createfolder(['ERROR-FETCHING'], mailbox)
        elif count == 1:
            return createfolder(['PYTHON-SORT', 'SINGLE-EMAIL'], mailbox)
        elif spokeninputtimeout('  Not found.  Create this folder? ', 'y').lower().strip() == 'y':
        
            for FOLDER in FOLDERSTACK:
                pack = pack + '/' + FOLDER
                pack = pack.strip('/')
                    
                if mailbox.folder.exists(pack) == False:
                    println('  Creating folder', pack)
                    mailbox.folder.create(pack)
        else:
            return createfolder(['PYTHON-SORT', 'AUTOREVIEW'], mailbox)
        
    return FULLPATH
    
# ---------------

def println(key, value):
    do_log(key + ' ' + value)
    timeX = datetime.now().strftime("%H:%M:%S ")
    print(timeX, key, ':           ', value)
    
# ---------------

def deletefolder(server, folder, status):

    if '/' in folder.name:
        if '\\HasNoChildren' in folder.flags and status.get('MESSAGES') == 0:
            if spokeninputtimeout(folder.name + ' has no children folders and no emails.  Delete? ', 'y') == 'y':
                println('DELETE', folder.name)
                server.folder.delete(folder.name)
    
# ---------------

def moveemails(server, FULLPATH, listings):

    pack = ''
    
    for xid in listings:
    
        pack = pack + xid + ','
        
        if pack.count(',') == 10:
            server.move(pack.strip(','), FULLPATH)
            println('  Moving emails', pack)
            pack = ''
    
    if pack != '':
        server.move(pack.strip(','), FULLPATH)
        println('  Moving emails', pack)

    stat = server.folder.status(FULLPATH)
    print(stat)


# ---------------
# ---------------

configs = json.load(open('./config.json', 'r'))
min_timeout = 2
max_timeout = 60
dynamic_timeout = max_timeout


mode = spokeninput('Which mode? Delete / Move / Sort (D M S) ').upper()
server = imap_tools.MailBox(configs['host']).login(configs['user'], configs['pass'])


    

if mode == 'D':

    for cycle in range(1, 3):

        folders = list(server.folder.list())
        speakline('Cycle ' + str(cycle) + ' folders to scan', str(len(folders)))

        for f in folders:
                    
            try:
                
                print(f.name)  # FolderInfo(name='INBOX|cats', delim='|', flags=('\\Unmarked', '\\HasChildren'))
                stat = server.folder.status(f.name)
                
                deletefolder(server, f, stat)

            except Exception as e:
                speakline('Failed to stat folder for deletion', str(e))

        
            
elif mode == 'M':

    while True:

        go = spokeninput('Folder filter: ')
        go = '*' + go + '*'
        
        folders = list(server.folder.list(search_args=go))        
        speakline('Folders to scan', str(len(folders)))
        
        for f in folders:
            stat = server.folder.status(f.name)
            
            if '\\HasNoChildren' in f.flags and stat.get('MESSAGES') > 0:
                println(f.name, 'has no children folders and ' + str(stat.get('MESSAGES')) + ' emails')
            elif '\\HasNoChildren' in f.flags and stat.get('MESSAGES') == 0:
                deletefolder(server, f, stat)
            else:
                println(f.name, '')
                    
        if spokeninput('Empty all of these folders? ') == 'y':

            destinationfolder = spokeninput('Which folder to put in? ')
            
            stat = server.folder.status(destinationfolder)
            print(stat)


            for f in folders:
                println(f.name, '')
                server.folder.set(f.name)
                
                preview = list(server.fetch(bulk=True))
                EMAILLIST = []

                for index, msg in enumerate(preview):
                    EMAILLIST.append(msg.uid)


                moveemails(server, destinationfolder, EMAILLIST)
            

elif mode == 'S':

############### Login to Mailbox ######################

    server = imap_tools.MailBox(configs['host']).login(configs['user'], configs['pass'])

    println("Logged into mailbox", configs['host'])

    #################### List Emails #####################
    
    stat = server.folder.status('INBOX')
    print(stat)
    
    runtimecount = 0
    
    while True:
    
        if dynamic_timeout == min_timeout:
        
            preview = list(server.fetch(limit=1, bulk=True, reverse=True))
            selectedEmail = preview[0]

        else:
            preview = list(server.fetch(limit=7, bulk=True, reverse=True))

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
        
        
        
        EMAILLIST = []
        

        try:
        
            fromX = selectedEmail.from_values.email
            yearX = selectedEmail.date.strftime('%Y')
            searchString = 'FROM "'+fromX+'"'
            
            results = list(server.fetch(searchString, limit=500, bulk=True, reverse=True))
            
            println("Query", searchString)
            println("  Emails from " + fromX, str(len(results)))
            
            for index, msg in enumerate(results):
            
                thisYear = msg.date.strftime('%Y')
                thisName = msg.from_values.name
            
                if thisYear != yearX:
                    print('  Email year ' + thisYear)
                elif selectedEmail.from_values.name != thisName:
                    print('  Email from ' + thisName)
                else:
                    show_message(index, msg)        
                    EMAILLIST.append(msg.uid)
            
        except Exception as e:
            speakline('Failed to fetch emails', str(e))
            EMAILLIST.append(selectedEmail.uid)


        FOLDERSTACK = determinefolder(selectedEmail)
        FULLPATH = createfolder(FOLDERSTACK, server, len(EMAILLIST))                
        
        try:
            
            moveemails(server, FULLPATH, EMAILLIST)
            
            counting = len(EMAILLIST)
            runtimecount = runtimecount + counting
            
            speakline("  Emails sent in " + yearX  + " from " + fromX , str(counting))
            speakline("Total emails sorted", str(runtimecount))

        except Exception as e:
            speakline('Failed to move emails', str(e))


        

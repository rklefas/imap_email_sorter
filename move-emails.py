import imap_tools
from datetime import datetime
import json
import re
from win32com.client import Dispatch
from inputimeout import inputimeout, TimeoutOccurred
import unidecode

# ---------------
# ---------------

def do_log(message):
    dateX = datetime.now().strftime("%Y-%m-%d")
    timeX = datetime.now().strftime(" %H:%M:%S")
    file1 = open('logs/' + dateX + "-actions.log", "a")
    file1.write(dateX + timeX + " " + unidecode.unidecode(message) + "\n")
    file1.close()

# ---------------

def show_message(index, msg):
    print('Email', index, ' | ', msg.from_values.name, '  ', msg.from_values.email, '  ', msg.date)
    print('         |           ', msg.subject[0:50], '(', msg.uid, ')', str(int(len(msg.text or msg.html)/1024)) + 'kb')

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
    
    nameX = unidecode.unidecode(msg.from_values.name).strip().replace('  ', ' ')
    nameX = nameX.replace('/', '-')
    
    if nameX == '':
        nameX = msg.from_values.email

    nameX = nameX + msg.date.strftime(' [%Y]')

    FOLDERSTACK.append(nameX)
    
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

def speaknumber(key, val):
    if val > 20:
        speakline(key, str(val))
    else:
        println(key, str(val))

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
    
        altered = folder.name
        altered = altered.replace(' 2018', ' [2018]')
        altered = altered.replace(' 2019', ' [2019]')
        altered = altered.replace(' 2020', ' [2020]')
        altered = altered.replace(' 2021', ' [2021]')
        altered = altered.replace(' 2022', ' [2022]')
        altered = altered.replace(' 2023', ' [2023]')
        
        if folder.name != altered:
            server.folder.rename(folder.name, altered)
            println('Folder renamed', altered)
            return
    
        if '\\HasNoChildren' in folder.flags and status.get('MESSAGES') == 0:
            println(folder.name, 'has no folders or emails')
            println('  DELETE', folder.name)
            server.folder.delete(folder.name)
            return 1
            
    return 0
    
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

def refresh_connection():
    configs = json.load(open('./config.json', 'r'))
    server = imap_tools.MailBox(configs['host']).login(configs['user'], configs['pass'])
    speakline("Logged into mailbox", configs['host'])

    return server

# ---------------

def mode_delete():

    for cycle in range(1, 3):

        count = 0
        server = refresh_connection()
        
        folders = list(server.folder.list())
        speakline('Cycle ' + str(cycle) + ' folders to scan', str(len(folders)))

        for f in folders:
                    
            try:
                print(f.name) 
                stat = server.folder.status(f.name)
                
                count += deletefolder(server, f, stat)

            except Exception as e:
                speakline('Failed to stat folder for deletion', str(e))

        speakline('Deleted folders', str(count))

# ---------------

def mode_read(server, folderx):

    server.folder.set(folderx)
    speakline('Current Folder', folderx)
    
    while True:
        
        preview = list(server.fetch(criteria=imap_tools.AND(seen=False), limit=50, bulk=True, reverse=True, mark_seen=False))
        
        if (len(preview) == 0):
            speakline('', 'No unseen emails left.  Loading already seen emails.')
            preview = list(server.fetch(criteria=imap_tools.AND(seen=True), limit=50, bulk=True, reverse=True, mark_seen=False))
            
        alllength = 0
        
        for index, msg in enumerate(preview):
                    
            shrunken = cleanbody(msg.text)
            alllength += len(shrunken)
            
        speakline('Fetched Emails', str(len(preview)))
        speakline('Hours To Read', str(int(alllength / 36000)))
        speaknumber('Emails Length', alllength)

        if (len(preview) == 0):
            return

        for index, msg in enumerate(preview):
        
            speakline('From', msg.from_values.name)
            speakline('Date', str(msg.date))
            speakline('Subject', msg.subject)
            
            shrunken = cleanbody(msg.text)
            
            speakline('Minutes To Read', str(int(len(shrunken) / 600)))
            speaknumber('Email Length', len(shrunken))
            
            if spokeninputtimeout('Do you want to read this email? ', 'y') == 'y':
                
                speakitem(shrunken)
                
                try:
                    server.folder.status(folderx)
                except Exception as e:
                    server = refresh_connection()
            
                server.flag([msg.uid], imap_tools.MailMessageFlags.SEEN, True)

            after_command = spokeninput('Email end.  Press D to delete or S to save.  Q to quit. ')

            if after_command == 'd' or after_command == 'dq':
                moveemails(server, 'Trash', [msg.uid])
                speakline('', 'Email deleted')

            if after_command == 's' or after_command == 'sq':
                speakline('', 'Email saved')
            
            if after_command == 'q' or after_command == 'dq' or after_command == 'sq':
                return

# ---------------

def cleanreplacer(vv, find, puts):

    xx = vv.replace(find, puts)
    befores = len(vv)
    afters = len(xx)

    if (befores != afters):
        print('Find', find)
        print('Puts', puts)
        print('Before Length', befores)
        print('After Length', afters)

    return xx

# ---------------

def cleanbody(vv):

    vv = cleanreplacer(vv, '* * ', '****')
    vv = cleanreplacer(vv, '- - ', '----')
    vv = cleanreplacer(vv, '&nbsp;', ' ')
    vv = cleanreplacer(vv, '&amp;', ' and ')
    vv = cleanreplacer(vv, '==', '**')
    vv = cleanreplacer(vv, '*=', '**')
    vv = cleanreplacer(vv, '__', '**')
    vv = cleanreplacer(vv, '*_', '**')
    vv = cleanreplacer(vv, '  ', ' ')
    vv = cleanreplacer(vv, '<', '-')
    vv = cleanreplacer(vv, '>', '-')
    vv = cleanreplacer(vv, '\r\n\r\n\r\n', '\r\n')
    vv = cleanreplacer(vv, 'https:', 'http:')
    
    vv = vv.strip()
    vv = re.sub("http://(\S+)", "", vv)
    
    vv = breakfooter(vv, 'Copyright © 202')
    vv = breakfooter(vv, 'You are receiving this email')
    vv = vv.strip()

    return vv

# ---------------

def breakfooter(xx, breakoff):
    return xx

# ---------------

def speakitem(vv):

    parts = vv.split('\r\n')
    
    for part in parts:
        print(part)
        Dispatch("SAPI.SpVoice").Speak(part)

# ---------------

def folderselection(server):
    go = spokeninput('Folder filter: ')
    go = '*' + go + '*'
    
    folders = list(server.folder.list(search_args=go))
    speakline('Folders found', str(len(folders)))
    
    if len(folders) == 0:
        return folderselection(server)
        
    for f in folders:
        print(f.name)

    if len(folders) == 1:
        return folders

    if spokeninput('Do you want to select these folders? ') != 'y':
        return folderselection(server)
        
    return folders

# ---------------
# ---------------

min_timeout = 2
max_timeout = 60
dynamic_timeout = max_timeout


println('Press S', 'To sort your inbox to subfolders.')
println('Press M', 'To empty out select subfolders.')
println('Press D', 'To delete empty subfolders.')
println('Press R', 'To read emails in your inbox.')
    
mode = spokeninput('Select a mode: ').upper()

if mode == 'D':

    mode_delete()        

elif mode == 'R':

    while True:
    
        server = refresh_connection()
        folders = folderselection(server)

        for f in folders:

            stat = server.folder.status(f.name)
            
            if stat.get('MESSAGES') > 0:
                mode_read(server, f.name)


elif mode == 'M':

    while True:
    
        server = refresh_connection()

        folders = folderselection(server)
        
        for f in folders:
            try:

                stat = server.folder.status(f.name)
                
                if '\\HasNoChildren' in f.flags and stat.get('MESSAGES') > 0:
                    println(f.name, 'has no children folders and ' + str(stat.get('MESSAGES')) + ' emails')
                elif '\\HasNoChildren' in f.flags and stat.get('MESSAGES') == 0:
                    deletefolder(server, f, stat)
                else:
                    println(f.name, '')
                    
            except Exception as e:
                println(f.name, '')
                speakline('Failed to prepare folder for moving', str(e))

                    
        if spokeninput('Empty all of these folders? ') == 'y':
        
            folders = list(server.folder.list(search_args=go))
            destinationfolder = spokeninput('Which folder to put in? ')
            
            stat = server.folder.status(destinationfolder)
            print(stat)


            for f in folders:
                try:
                    for cycle in range(1, 100):
                    
                        println(f.name, '')
                        server.folder.set(f.name)
                    
                        preview = list(server.fetch(bulk=True, limit=100))
                        
                        if len(preview) == 0:
                            break
                            
                        FILTERED_UIDS = []

                        for index, msg in enumerate(preview):
                            FILTERED_UIDS.append(msg.uid)

                        moveemails(server, destinationfolder, FILTERED_UIDS)
                        
                
                except Exception as e:
                    println(f.name, '')
                    speakline('Failed to stat folder for moving', str(e))

            

elif mode == 'S':

    ############### Login to Mailbox ######################

    server = refresh_connection()

    #################### List Emails #####################
    
    
    runtimecount = 0
    
    for cycle in range(1, 200):
    
        if cycle % 25 == 0:
            server = refresh_connection()    
    
        if dynamic_timeout == min_timeout:
        
            preview = list(server.fetch(limit=1, bulk=True, reverse=True))
            peekEmail == '0'

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

        
        
        if len(preview) == 0:
            speakline('Congratulations!', 'You have achieved inbox zero.')
            
            if spokeninputtimeout('Do you want to run delete mode? ', 'y') == 'y':
                mode_delete()
            
            break
        
        selectedEmail = preview[int(peekEmail)]
        FILTERED_UIDS = []
        

        try:
        
            fromX = selectedEmail.from_values.email
            yearX = selectedEmail.date.strftime('%Y')
            searchString = 'FROM "'+fromX+'"'
            
            FETCHED_EMAILS = list(server.fetch(searchString, limit=500, bulk=True, reverse=True))
            
            println("Query", searchString)
            println("  Emails from " + fromX, str(len(FETCHED_EMAILS)))
            
            for index, msg in enumerate(FETCHED_EMAILS):
            
                thisYear = msg.date.strftime('%Y')
                thisName = msg.from_values.name
            
                if thisYear != yearX:
                    print('  Email year ' + thisYear)
                elif selectedEmail.from_values.name != thisName:
                    print('  Email from ' + thisName)
                else:
                    show_message(index, msg)        
                    FILTERED_UIDS.append(msg.uid)
            
        except Exception as e:
            speakline('Failed to fetch emails', str(e))
            FILTERED_UIDS.append(selectedEmail.uid)


        FOLDERSTACK = determinefolder(selectedEmail)
        FULLPATH = createfolder(FOLDERSTACK, server, len(FILTERED_UIDS))                
        
        try:
            
            moveemails(server, FULLPATH, FILTERED_UIDS)
            
            counting = len(FILTERED_UIDS)
            runtimecount = runtimecount + counting
            
            speaknumber("  Emails sent in " + yearX  + " from " + fromX , counting)
            speaknumber("Total emails sorted", runtimecount)


        except Exception as e:
            speakline('Failed to move emails', str(e))


        

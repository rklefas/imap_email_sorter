import imap_tools
from datetime import datetime
import gc # Garbage Collector
import json
import re
from win32com.client import Dispatch
from inputimeout import inputimeout, TimeoutOccurred
import random
import textwrap
import time
import unidecode
from bs4 import BeautifulSoup
import yake

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

def summarizer(msg):

    print('-------------------------')
    println('From', msg.from_values.name)
    println('Date', str(msg.date))
    println('Subject', msg.subject)
    println('Flags', msg.flags)
    
    if len(msg.html):
        bodytype = 'html'
        raw = BeautifulSoup(msg.html).body.get_text()
    else:
        bodytype = 'text'
        raw = msg.text
        
    shrunken = cleanbody(msg)

    print('--------- BODY ----------')
    println('Body Type', bodytype)
    println('Time To Read', timetoread(len(shrunken)))
    println('Link Count', bodylinks(raw))
    println('Readability', readability(raw, shrunken))
    println('Shrunken Length', len(shrunken))
    println('Key phrases', getkeywords(shrunken))

# ---------------

def rearrangefrom(frommer):
    temp = frommer.split('@')
    parts = temp[1].lower().split('.')
    parts.reverse()
    temp = '.'.join(parts)
    return temp

# ---------------

def determinefolder(msg):

    mailbox = refresh_connection()
    
    foundPriority = 'PYTHON-SORT'
    
    for baseFolder in ['PRIORITY-A', 'PRIORITY-B', 'PRIORITY-C', 'PRIORITY-F', 'PYTHON-SORT']:

        TMPSTACK = [baseFolder]
        TMPSTACK.append(rearrangefrom(msg.from_values.email))
        
        if mailbox.folder.exists('/'.join(TMPSTACK)) == True:
            foundPriority = baseFolder
            break
            
            
    FOLDERSTACK = [foundPriority]
    FOLDERSTACK.append(rearrangefrom(msg.from_values.email))

    acc = msg.from_values.email.split('@')
    
    nameX = msg.date.strftime('%Y')
    nameX = nameX + ' '+ acc[0] + ' '
    nameX = nameX + '(' + unidecode.unidecode(msg.from_values.name).strip().replace('  ', ' ') + ')'
    nameX = nameX.replace('/', '-')
    
    if nameX == '':
        nameX = msg.from_values.email

    FOLDERSTACK.append(nameX)
    
    return FOLDERSTACK

# ---------------

def input_for_mode_selection(q, default_or_prompt):

    if default_or_prompt == '[ask]':
        return spokeninput(q)
    else:
        return spokeninputtimeout(q, default_or_prompt)

# ---------------

def prettyinput(q):
    print('-----------------------')
    xx = input('<< ' + q + ' >> ')
    print('-----------------------')
    return xx

# ---------------

def spokeninput(q):
    Dispatch("SAPI.SpVoice").Speak(q)
    return prettyinput(q)
    
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
#    Dispatch("SAPI.SpVoice").Speak(key + ' ' + val)

# ---------------

def createfolder(FOLDERSTACK, count = None):

    mailbox = refresh_connection()


    FULLPATH = '/'.join(FOLDERSTACK)
    pack = ''

    if mailbox.folder.exists(FULLPATH) == False:
    
        if mailbox.folder.exists(folderparent(FULLPATH)) == True:
            println('  Domain parent folder found, creating child folder', FULLPATH)
            mailbox.folder.create(FULLPATH)
        elif count == 0:
            return createfolder(['ERROR-FETCHING'])
        elif count == 1:
            return createfolder(['PYTHON-SORT', 'SINGLE-EMAIL'])
        elif spokeninputtimeout('  Not found.  Create this folder? ', 'y').lower().strip() == 'y':
        
            for FOLDER in FOLDERSTACK:
                pack = pack + '/' + FOLDER
                pack = pack.strip('/')
                    
                if mailbox.folder.exists(pack) == False:
                    println('  Creating folder', pack)
                    mailbox.folder.create(pack)
        else:
            return createfolder(['PYTHON-SORT', 'AUTOREVIEW'])
        
    println('Check for creation', FULLPATH)
    
    return FULLPATH
    
# ---------------

def println(key, value):
    do_log(key + ' ' + str(value))
    timeX = datetime.now().strftime("%H:%M:%S ")
    print(timeX, str(key + ':').ljust(25, " "), value)
    
# ---------------

def dropfolder(fname):

    println('Trying to drop folder', fname)

    server = refresh_connection()
    folders = list(server.folder.list(search_args=fname))
    
    for f in folders:
        deletefolder(server, f)
    
# ---------------

def deletefolder(server, folder):

    # DO NOT DELETE DOMAIN FOLDERS THAT ARE PRIORITIZED
    if 'PRIORITY-' in folder.name:
        if folderdepth(folder.name) == 2:
            return 0
            

    if 'PYTHON' in folder.name:
                
        if folderdepth(folder.name) == 3:
        
            fparent = folderparent(folder.name)
        
            if server.folder.exists(fparent) == False:
                println('Parent missing', fparent)
                createfolder(fparent.split('/'))
                return 0

    if folderdepth(folder.name) > 1:
    
        status = server.folder.status(folder.name)

        if '\\HasNoChildren' in folder.flags and status.get('MESSAGES') == 0:
            println('  DELETING EMPTY FOLDER', folder.name)
            server.folder.delete(folder.name)
            return 1
            
    return 0
    
# ---------------

def moveemails(server, FULLPATH, uid_list):

    pack = ''
    
    for uid_one in uid_list:
    
        pack = pack + uid_one + ','
        
        if pack.count(',') == 10:
            print(server.move(pack.strip(','), FULLPATH))
            println('  Moving emails', pack)
            pack = ''
    
    if pack != '':
        print(server.move(pack.strip(','), FULLPATH))
        println('  Moving emails', pack)

    stat = server.folder.status(FULLPATH)
    print(stat)

# ---------------

def refresh_connection(set_folder = None):

    global mailbox_server
    
    try:
        mailbox_server.folder.status()
    except Exception as e:

#        if random.randint(0, 1) == 1:
#            return mailbox_server

        configs = json.load(open('./config.json', 'r'))
        mailbox_server = imap_tools.MailBox(configs['host']).login(configs['user'], configs['pass'])
        println("Logged into mailbox", configs['host'])

    
    if set_folder != None:
        mailbox_server.folder.set(set_folder)
        println('  Browsing Folder', mailbox_server.folder.get())
    

    return mailbox_server

# ---------------

def mode_prioritize(folderx):

   
    if folderdepth(folderx) == 2:
    
        print('-----------------------')
        println('Prioritize folder', folderx)
        
        pri = prettyinput('What priority? (A B C F ?) ').upper().strip()
        
        if pri == 'A' or pri == 'B' or pri == 'C' or pri == 'F':
            topfolder = 'PRIORITY-' + pri
        else:
            return
         
        createfolder([topfolder])

        folderparts = folderx.split('/')
        folderparts.reverse()
        folderparts.pop()
        folderparts.append(topfolder)
        folderparts.reverse()
        
        newfolder = '/'.join(folderparts)
        
        folder_rename(folderx, newfolder)
        


def mode_queue(folderx):

    server = refresh_connection(folderx)
    
    while True:
        
        preview = list(server.fetch(criteria=imap_tools.AND(seen=False), limit=100, bulk=True, reverse=True, mark_seen=False))
        
        if (len(preview) == 0):
            if spokeninput('Do you want to look for already seen emails? ') == 'y':
                preview = list(server.fetch(criteria=imap_tools.AND(seen=True), limit=100, bulk=True, reverse=True, mark_seen=False))
            
        speakline('Fetched Emails', str(len(preview)))
        
        if (len(preview) == 0):
            dropfolder(folderx)
            return

        actionstack = []

        for index, msg in enumerate(preview):
            
            summarizer(msg)
            
            after_command = prettyinput('Press R to read.  Press T to trash or S to star.  Q to run queue now. ')
            
            if after_command == 'q':
                break

            actionstack.append(after_command)
            speakline('Emails in Queue', str(len(actionstack)))

        for index, msg in enumerate(preview):

            after_command = actionstack.pop(0)
            mode_read_process(msg, after_command)

# ---------------

def mode_read_process(msg, after_command):

    server = refresh_connection()

    if after_command == 'r':
    
        summarizer(msg)
        
        shrunken = cleanbody(msg)
        
        speakitem(shrunken)

        after_command = spokeninput('Email end.  Press T to trash or S to star. ')
        
        try:
            server.folder.status()
        except Exception as e:
            server = refresh_connection()
    
        server.flag([msg.uid], imap_tools.MailMessageFlags.SEEN, True)

    if after_command == 't':
        moveemails(server, 'Trash', [msg.uid])
        speakline('', 'Email deleted')

    if after_command == 's':
        moveemails(server, 'Review Later', [msg.uid])
        speakline('', 'Email starred')

# ---------------

def mode_delete():

    for cycle in range(1, 3):

        count = 0
        server = refresh_connection()
        
        folders = folderselection()
        speakline('Cycle ' + str(cycle) + ' folders to scan', str(len(folders)))

        FOLDERCOUNTS = []
        uniqueDomains = 0
        
        for f in folders:
            if folderdepth(f.name) == 2:
                uniqueDomains += 1
        
        speakline('  Unique Domains', uniqueDomains)


        for f in folders:
        
            if folderdepth(f.name) == 2:
                uniqueDomains -= 1
                println('  Domains left', uniqueDomains)
                    
            try:
                stat = server.folder.status(f.name)
                
                tmp = str(stat.get('MESSAGES')).rjust(5, '0')
                tmp = tmp + ' - ' + f.name
                
                print(tmp)

                FOLDERCOUNTS.append(tmp)
                count += deletefolder(server, f)

            except Exception as e:
                println('Failed to stat folder for deletion', str(e))


        speakline('Deleted folders', str(count))
        
        print(sorted(FOLDERCOUNTS))

# ---------------

def mode_read(folderx, mode_selection):

    server = refresh_connection(folderx)
    speakline('Current Folder', folderx)
    
    while True:
        
        preview = list(server.fetch(criteria=imap_tools.AND(seen=False), limit=50, bulk=True, reverse=True, mark_seen=False))
        
        if (len(preview) == 0):
            if spokeninput('Do you want to look for already seen emails? ') == 'y':
                preview = list(server.fetch(criteria=imap_tools.AND(seen=True), limit=100, bulk=True, reverse=True, mark_seen=False))
            
        speakline('Fetched Emails', str(len(preview)))
        
        if (len(preview) == 0):
            dropfolder(folderx)
            return

        alllength = 0
        
        for index, msg in enumerate(preview):
                    
            shrunken = cleanbody(msg)
            alllength += len(shrunken)
            
        speakline('Time To Read', timetoread(alllength))
        speaknumber('Emails Length', alllength)
        print('-----------------------')

        for index, msg in enumerate(preview):
            
            summarizer(msg)
            shrunken = cleanbody(msg)
            
            if len(shrunken) == 0:
                continue
            
            default_preview = 'r' if mode_selection == 'SL' else '[ask]'
            default_finish = 't' if mode_selection == 'SL' else '[ask]'
            
            after_command = input_for_mode_selection('Press R to read.  Press T to trash or S to star.  Q to quit. ', default_preview)

            if after_command == 'r':
                
                speakitem(shrunken)

                after_command = input_for_mode_selection('Email end.  Press T to trash or S to star.  Q to quit. ', default_finish)
                
                try:
                    server.folder.status(folderx)
                except Exception as e:
                    server = refresh_connection(folderx)
            
                server.flag([msg.uid], imap_tools.MailMessageFlags.SEEN, True)

                

            if after_command == 't' or after_command == 'tq':
                moveemails(server, 'Trash', [msg.uid])
                speakline('', 'Email deleted')

            if after_command == 's' or after_command == 'sq':
                moveemails(server, 'Review Later', [msg.uid])
                speakline('', 'Email starred')
            
            if after_command == 'q' or after_command == 'dq' or after_command == 'sq':
                return 'q'

# ---------------

def timetoread(xx):
    
    tmp = int(xx / 600)
    unit = ' Minutes'
    
    if (tmp > 59):
        tmp = int(tmp / 60)
        unit = ' Hours'

    return str(tmp) + unit

# ---------------

def cleanreplacer(vv, find, puts):

    xx = vv.replace(find, puts)
    
    return xx

    befores = len(vv)
    afters = len(xx)

    if (befores != afters):
        print('Find', find)
        print('Puts', puts)
        print('Before Length', befores)
        print('After Length', afters)

    return xx

# ---------------

def readability(raw, cleaned):
    return str(int((len(cleaned) / len(raw)) * 100)) + '%'

# ---------------

def cleanbody(msg):

    if len(msg.html):
        bodytype = 'html'
        vv = BeautifulSoup(msg.html).body.get_text()
    else:
        bodytype = 'text'
        vv = msg.text

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
    vv = cleanreplacer(vv, '\r\n', '\n')
    vv = cleanreplacer(vv, '\n\n\n', '\n')
    vv = cleanreplacer(vv, 'https:', 'http:')
    
    vv = vv.strip()
    vv = re.sub("http://(\S+)", "", vv)
    
    vv = breakfooter(vv, 'Copyright © 20')
    vv = breakfooter(vv, 'You are receiving this email')
    vv = breakfooter(vv, 'If you no longer wish to receive our emails')
    vv = vv.strip()

    return vv
    
# ---------------

def bodylinks(raw):

    xx = raw.count("http://") + raw.count("https://") 
    
    return xx

# ---------------

def getkeywords(texty):

    limit = int(len(texty) / 500)
    limit = min(limit, 30)

    kw_extractor = yake.KeywordExtractor(top=limit)
    keywords = kw_extractor.extract_keywords(texty)
    phrases = []
    
    for kw, v in keywords:
        phrases.append(kw)
    

    return "    ".join(phrases)

# ---------------

def breakfooter(xx, breakoff):

    position = xx.find(breakoff)
    
    if position > 50:
        xx = xx[0:position]

    return xx.strip()

# ---------------

def speakitem(vv):

#    tmp = textwrap.wrap(vv, replace_whitespace=False, drop_whitespace=False)
    parts = vv.split('\n')
    count = len(parts)
    start_time = time.time()
    
    for index, part in enumerate(parts):
    
        seconds = time.time() - start_time
        convert = time.strftime("%M:%S", time.gmtime(seconds))
    
        print('(', (index+1), 'of', count, ')  [', convert, ']  ', part)
        Dispatch("SAPI.SpVoice").Speak(part)

# ---------------

def folderselection():
    go = spokeninput('Folder filter: ')
    go = '*' + go + '*'
    
    server = refresh_connection()
    folders = list(server.folder.list(search_args=go))

    println('Folders found', len(folders))
    
    if len(folders) == 0:
        return folderselection()
        
    for f in folders:
        print(f.name)

    if len(folders) == 1:
        return folders
        
    println('Folders found', len(folders))

    if spokeninput('Do you want to select these folders? ') != 'y':
        return folderselection()
        
    server = refresh_connection()

    return folders
    
# ---------------

def folderchildren(folderx):
    server = refresh_connection()
    folders = list(server.folder.list(search_args=folderx + '/*'))

    return folders

# ---------------

def folderparent(folderx):
    folderstack = folderx.split('/')
    folderstack.pop()
    
    return "/".join(folderstack)

# ---------------

def folderdepth(folderx):
    folderstack = folderx.split('/')
    
    return len(folderstack)

# ---------------

def folder_rename(oldname, newname):
    server = refresh_connection()

    println('Folder Name', oldname)
    println('  Rename To', newname)
    
    if server.folder.exists(folderparent(newname)) == False:
        println('  Error', 'Parent does not exist')
        return
        
    server.folder.rename(oldname, newname)

# ---------------

def mode_sort():

    ############### Login to Mailbox ######################

    server = refresh_connection()

    #################### List Emails #####################
    
    
    runtimecount = 0
    
    for cycle in range(1, 200):
    
        if cycle % 25 == 0:
            server = refresh_connection()
    
        if dynamic_timeout == min_timeout:
        
            preview = list(server.fetch(limit=1, bulk=True, reverse=True))
            peekEmail = '0'

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
        FULLPATH = createfolder(FOLDERSTACK, len(FILTERED_UIDS))                
        
        try:
            
            moveemails(server, FULLPATH, FILTERED_UIDS)
            
            counting = len(FILTERED_UIDS)
            runtimecount = runtimecount + counting
            
            speaknumber("  Emails sent in " + yearX  + " from " + fromX , counting)
            speaknumber("Total emails sorted", runtimecount)


        except Exception as e:
            speakline('Failed to move emails', str(e))




# ---------------
# ---------------

min_timeout = 2
max_timeout = 60
dynamic_timeout = max_timeout
mailbox_server = None


println('Press A', 'Run (A)ll day and keep inbox sorted')
println('Press S', 'Automatically sort emails in your inbox to subfolders.')
println('Press M', 'Empty out select subfolders.')
println('Press C', 'Cleanup mailbox.  Delete empty subfolders.')
println('Press P', 'Prioritize senders')
println('Press R', 'Read emails in your inbox or other folder.')
println('Press SL', 'Sit and Listen.  Automatically read and delete emails in your inbox or other folder.')
println('Press Q', 'Fill up a player queue')


mode_selection = spokeninput('Select a mode: ').upper()

if mode_selection == 'A':

    while True:
    
        mode_sort()

        if random.randint(0, 4) == 1:
            mode_delete()
        
        gc.collect()
        print('Waiting for a while...')
        time.sleep(60*60*4)

    
    
elif mode_selection == 'C':

    mode_delete()

elif mode_selection == 'Q':

    while True:
    
        folders = folderselection()

        for f in folders:
            mode_queue(f.name)



elif mode_selection == 'P':

    while True:
    
        folders = folderselection()

        for f in folders:
            mode_prioritize(f.name)


elif mode_selection == 'R' or mode_selection == 'SL':

    while True:
    
        folders = folderselection()

        for f in folders:

            if mode_read(f.name, mode_selection) == 'q':
                break


elif mode_selection == 'M':

    while True:
    
        server = refresh_connection()
        folders = folderselection()
        
        for f in folders:
            try:

                stat = server.folder.status(f.name)
                
                if '\\HasNoChildren' in f.flags and stat.get('MESSAGES') > 0:
                    println(f.name, 'has no children folders and ' + str(stat.get('MESSAGES')) + ' emails')
                elif '\\HasNoChildren' in f.flags and stat.get('MESSAGES') == 0:
                    deletefolder(server, f)
                else:
                    println(f.name, '')
                    
            except Exception as e:
                println(f.name, '')
                speakline('Failed to prepare folder for moving', str(e))

                    
        if spokeninput('Empty all of these folders? ') == 'y':
        
            println('Option', 'INBOX')
            println('Option', 'Trash')
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
                            deletefolder(server, f)
                            break
                            
                        FILTERED_UIDS = []

                        for index, msg in enumerate(preview):
                            FILTERED_UIDS.append(msg.uid)

                        moveemails(server, destinationfolder, FILTERED_UIDS)
                        
                
                except Exception as e:
                    println(f.name, '')
                    speakline('Failed to stat folder for moving', str(e))

            

elif mode_selection == 'S':

    mode_sort()
    


        

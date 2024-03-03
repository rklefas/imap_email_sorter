import imap_tools
from datetime import datetime
import gc # Garbage Collector
import json
import re
from win32com.client import Dispatch
from inputimeout import inputimeout, TimeoutOccurred
import os
import pathlib
import random
import textwrap
import time
import unidecode
from bs4 import BeautifulSoup
import yake

# ---------------
# ---------------

def screen_clear():
    print('')
    print('Switching to new screen...')
    os.system('cls')

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

def summarizer(msg, folder = None, clearing = False):

    if clearing == True:
        screen_clear()
    
    if folder != None:
        println('Folder', folder)
        print('-------------------------')
        
    println('From', msg.from_values.name)
    println('Date', str(msg.date))
    speakline('Subject', msg.subject)
    println('Flags', msg.flags)

    
    raw = cleantext(msg.html, 'html')
    raw = BeautifulSoup(raw, "html.parser").body.get_text()
    shrunken = cleantext(raw)
    bodysummary('html', msg.html, shrunken)


    shrunken = cleantext(msg.text, 'text')
    bodysummary('text', msg.text, shrunken)

# ---------------

def bodysummary(bodytype, raw, shrunken):

    print('---------', bodytype.upper(), 'BODY ----------')
    println('Raw Length', len(raw))
    println('  Link Count', bodylinks(raw))
    println('  Readability %', readability(raw, shrunken))
    println('Shrunken Length', len(shrunken))
    println('  Time To Read', timetoread(len(shrunken)))
    println('  Key phrases', getkeywords(shrunken))

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
    
    for baseFolder in ['MANUAL-REVIEW', 'PRIORITY-A', 'PRIORITY-B', 'PRIORITY-C', 'PRIORITY-F', 'PYTHON-SORT']:

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

    global dynamic_timeout


    if dynamic_timeout == None:
        return spokeninput(q)
    else:
        return spokeninputtimeout(q, default_or_prompt)

# ---------------

def prettyinput(q, defaultValue = None):
    print('')
    tmp = '<< ' + q + ' >> '
    
    if defaultValue != None:
        print(' *** Default value is', defaultValue)
    
    print('-' * len(tmp))
    xx = input(tmp)
    
    if xx == '' and defaultValue != None:
        return defaultValue
    
    print('-' * len(tmp))
    return xx

# ---------------

def spokeninput(q, defaultValue = None):
    Dispatch("SAPI.SpVoice").Speak(q)
    return prettyinput(q, defaultValue)
    
# ---------------
   
def inputcontrols(num = None):

    if num == None:
        num = spokeninput('What timeout for the inputs?  0 will turn off timeouts, empty will change nothing: ')
    
    global dynamic_timeout

    if num == '0':
        dynamic_timeout = None
    elif num == '':
        print('No change to timeout setting')
    else:
        dynamic_timeout = int(num)
        
    return dynamic_timeout
    
# ---------------
   
def spokeninputtimeout(q, default, specific_timeout = None):
    
    global dynamic_timeout
    
    if specific_timeout == None:
        newtx = dynamic_timeout
    else:
        newtx = specific_timeout

    if newtx != None and newtx > 30:
        Dispatch("SAPI.SpVoice").Speak(q)
    
    try:
        print('')
        print_divider()
        print('  Enter == to alter timeout')

        val = inputimeout(q + ' (' + str(newtx) + ' sec timeout, default : ' + default + ') ', newtx)

        print_divider()
        
        if val == '==':
            newtx = inputcontrols()
        
            if newtx == None:
                return spokeninput(q, default)
            else:
                return spokeninputtimeout(q, default, newtx)
            
        return val
    except TimeoutOccurred:
    
        if newtx != None and newtx > 30:
            Dispatch("SAPI.SpVoice").Speak('Defaulted to '+default)

        return default

# ---------------

def speakline(key, val):
    println(key, val)
    
    global dynamic_timeout

    if dynamic_timeout != None and dynamic_timeout > 30:
        Dispatch("SAPI.SpVoice").Speak(key + ' ' + str(val))

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
        elif input_for_mode_selection('  Not found.  Create this folder? ', 'y').lower().strip() == 'y':
        
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

def print_divider():
    print('----------------------------------')
    
# ---------------

def println(key, value):
    do_log(key + ' ' + str(value))
    timeX = datetime.now().strftime("%H:%M:%S ")
    print(timeX, str(key + ':').ljust(25, " "), value)
    
# ---------------

def dropfolder(fname):

    println('Checking to drop folder', fname)

    server = refresh_connection()
    folders = list(server.folder.list(search_args=fname))
    
    for f in folders:
        deletefolder(server, f)
    
# ---------------

def deletefolder(server, folder):

    # DO NOT DELETE DOMAIN FOLDERS THAT ARE PRIORITIZED
    if 'PRIORITY-' in folder.name or 'MANUAL-REVIEW' in folder.name:
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

def reliable_fetch(p_limitx):

    server = refresh_connection()
    try:
        
        preview = list(server.fetch(criteria=imap_tools.AND(seen=False), limit=p_limitx, bulk=True, reverse=True, mark_seen=False))
        
        # todo: add fetch unread and read messages to this
    except Exception as e:
        server = refresh_connection()
        preview = list(server.fetch(criteria=imap_tools.AND(seen=False), limit=p_limitx, bulk=True, reverse=True, mark_seen=False))
        
    if (len(preview) == 0):
        preview = list(server.fetch(limit=p_limitx, bulk=True, reverse=True))

    return preview

# ---------------

def reliable_move(FULLPATH, x_uid):

    server = refresh_connection()
    
    if server.folder.exists(FULLPATH) == False:
        FOLDERSTACK = FULLPATH.split('/')
        createfolder(FOLDERSTACK)

    print_divider()
    speakline('Move to Folder', FULLPATH)
    
    # todo : this must verify that the moves happened
    moveemails(server, FULLPATH, [x_uid])

# ---------------

def moveemails(server, FULLPATH, uid_list):

    pack = ''
    remainingEmails = len(uid_list)
    
    for uid_one in uid_list:
    
        pack = pack + uid_one + ','
        remainingEmails -= 1
        
        if pack.count(',') == 10 or remainingEmails == 0:
        
            pack = pack.strip(',')
        
            result = server.move(pack, FULLPATH)            
            print('Create', result[0])
            print('Delete', result[1])
            
            println('  Moved emails', pack)
            pack = ''
    

    stat = server.folder.status(FULLPATH)
    print(stat)

# ---------------

def refresh_connection(set_folder = None):

    global mailbox_server
    
    try:
        mailbox_server.folder.status()
    except Exception as e:

        configs = json.load(open('./config.json', 'r'))
        mailbox_server = imap_tools.MailBox(configs['host']).login(configs['user'], configs['pass'])
        println("Logged into mailbox", configs['host'])

    
    if set_folder != None:
        mailbox_server.folder.set(set_folder)
        println('Reconnected to Folder', mailbox_server.folder.get())

    return mailbox_server

# ---------------

def mode_prioritize(folderx):

    if folderdepth(folderx) == 2:
    
        print_divider()
        println('Prioritize folder', folderx)
        
        pri = prettyinput('Manual review?  Or what priority? (A B C F - M) ').upper().strip()
        
        if pri == 'A' or pri == 'B' or pri == 'C' or pri == 'F':
            topfolder = 'PRIORITY-' + pri
        elif pri == 'M':
            topfolder = 'MANUAL-REVIEW'
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

# ---------------

def mode_queue(folderx):

    server = refresh_connection(folderx)
    
    while True:
        
        preview = reliable_fetch(100)
            
        speakline('Fetched Emails', str(len(preview)))
        
        if (len(preview) == 0):
            dropfolder(folderx)
            return

        actionstack = []

        for index, msg in enumerate(preview):
            
            summarizer(msg)
            
            after_command = prettyinput('Press R to read.  Press T to trash or S to star.  Q to quit and run queue now. ')
            
            if exit_command(after_command):
                break

            actionstack.append(after_command)
            speakline('Emails in Queue', str(len(actionstack)))

        for index, msg in enumerate(preview):

            after_command = actionstack.pop(0)
            summarizer(msg, server.folder.get(), True)
            mode_read_process(msg, after_command, server.folder.get())

# ---------------

def mode_read_process(msg, after_command, folderx):

    server = refresh_connection(folderx)
    
    if after_command == 'r':
    
        shrunken = cleanbody(msg)

        if len(shrunken) == 0:
            reliable_move(folderx + '/Unreadable', msg.uid)
            return
            
        speakitem(shrunken)
        
        summarizer(msg, server.folder.get(), False)

        after_command = input_for_mode_selection('Email end.  Press T to trash or RV to review later. ', 't')
        
        try:
            server.folder.status()
        except Exception as e:
            server = refresh_connection(folderx)
    
        server.flag([msg.uid], imap_tools.MailMessageFlags.SEEN, True)

    if after_command == 'b' or after_command == 'bq':
        reliable_move('Review for Bugs', msg.uid)
        after_command = after_command.replace('b', '')

    if after_command == 't' or after_command == 'tq':
        reliable_move('Trash', msg.uid)
        after_command = after_command.replace('t', '')

    if after_command == 'rv' or after_command == 'rvq':
        reliable_move('Review Later', msg.uid)
        after_command = after_command.replace('rv', '')
        
    if after_command == 'q':
        return 'q'

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

def mode_move(folders):

    server = refresh_connection()
    
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
                
                    preview = reliable_fetch(100)
                    
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






def mode_read(folderx):

    if folderdepth(folderx) != 3:
        return

    speakline('READ MODE - Folder', folderx)
    server = refresh_connection(folderx)
    
    while True:
        
        preview = reliable_fetch(50)

        speakline('Fetched Emails', str(len(preview)))
        
        if (len(preview) == 0):
            dropfolder(folderx)
            return

        alllength = 0
        
        for index, msg in enumerate(preview):
                    
            shrunken = cleanbody(msg)
            alllength += len(shrunken)
            
        speakline('Time To Read', timetoread(alllength))
        speakline('Emails Length', alllength)
        print_divider()

        for index, msg in enumerate(preview):
            
            summarizer(msg, server.folder.get())
            after_command = input_for_mode_selection('Press R to read.  Press T to trash or S to star.  Q to quit. ', 'r')
            done_reading = mode_read_process(msg, after_command, server.folder.get())
            
            if exit_command(done_reading):
                return 'q'

# ---------------

def timetoread(xx):
    
    tmp = max(int(xx / 600), 0)
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
    
    if len(raw) == 0:
        return 0

    return int((len(cleaned) / len(raw)) * 100)

# ---------------

def cleanbody(msg):

    html_cleaned = cleantext(msg.html, 'html')
    html_cleaned = BeautifulSoup(html_cleaned, "html.parser").body.get_text()
    html_cleaned = cleantext(html_cleaned)
    
    text_cleaned = cleantext(msg.text, 'text')
    
    if len(html_cleaned) > len(text_cleaned):
        return 'html body\n\n' + html_cleaned
    else:
        return 'text body\n\n' + text_cleaned

# ---------------

def cleantext(vv, bodytype = None):

    if bodytype == 'html':

        # remove styles in the body
        vv = re.sub(r'<style(.+)</style>', '', vv, flags=re.DOTALL)
    
        vv = cleanreplacer(vv, '</p><', '</p>\n\n<')
        vv = cleanreplacer(vv, '</h1><', '</h1>\n\n<')
        vv = cleanreplacer(vv, '</div><', '</div>\n\n<')


    if bodytype == 'text':

        vv = cleanreplacer(vv, '<', '-')
        vv = cleanreplacer(vv, '>', '-')
    
    
    vv = cleanreplacer(vv, '* * ', '****')
    vv = cleanreplacer(vv, '- - ', '----')
    vv = cleanreplacer(vv, '&nbsp;', ' ')
    vv = cleanreplacer(vv, '&amp;', ' and ')
    vv = cleanreplacer(vv, '==', '**')
    vv = cleanreplacer(vv, '*=', '**')
    vv = cleanreplacer(vv, '__', '**')
    vv = cleanreplacer(vv, '*_', '**')
    
    vv = cleanreplacer(vv, 'https:', 'http:')
    vv = re.sub("http://(\S+)", "", vv)
    
    vv = cleanreplacer(vv, '   ', ' ')
    vv = cleanreplacer(vv, '   ', ' ')
    vv = cleanreplacer(vv, '\r\n', '\n')
    vv = cleanreplacer(vv, '\n \n', '\n\n')
    vv = cleanreplacer(vv, '\n\n\n', '\n')
    vv = breakfooter(vv, 'Privacy Policy')
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
    limit = max(limit, 3)

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
# put in textwrap

# put in options to exit speech early, async speech

    parts = vv.split('\n')
    count = len(parts)
    start_time = time.time()
    pause_time = None
    pause_count = 0
    
    for index, partraw in enumerate(parts):
    
        if pause_time == None:
            pause_count += 1
            pause_time = time.time() + (60 * pause_count)
        
        
        if time.time() > pause_time:
            if exit_command(spokeninputtimeout('X to stop, or do nothing to continue', '', 20)):
                break
            else:
                pause_time = None
            
        spokenpart = partraw.strip()
        runtime = time.time() - start_time
        convert = time.strftime("%M:%S", time.gmtime(runtime))
    
        print('(', (index+1), 'of', count, ')  [', convert, ']  ', spokenpart)
        
        try:
            Dispatch("SAPI.SpVoice").Speak(spokenpart)
        except Exception as e:
            Dispatch("SAPI.SpVoice").Speak('Recovering from exception. ')
            print(e)

# ---------------

def folderselection():
    go = prettyinput('Folder filter: ')
    go = '*' + go + '*'
    
    server = refresh_connection()
    folders = list(server.folder.list(search_args=go))

    println('Folders found', len(folders))
    
    if len(folders) == 0:
        return folderselection()
    
    showing = 0
    
    for f in folders:
        showing += 1
        print(showing, f.name)
        
        if (showing % 30 == 0):
            if exit_command(prettyinput('Press enter to show more folders, or q to quit: ')):
                break
            else:
                screen_clear()

    if len(folders) == 1:
        return folders
        
    println('Folders found', len(folders))

    if prettyinput('Do you want to select these folders? ') != 'y':
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

def exit_command(vv):

    if vv == None:
        return False
    if vv.upper() == 'Q':
        return True
    if vv.upper() == 'X':
        return True

    return False

# ---------------

def mode_sort():

    ############### Login to Mailbox ######################

    server = refresh_connection('INBOX')

    #################### List Emails #####################
    
    
    runtimecount = 0
    
    for cycle in range(1, 1000):
    
        if dynamic_timeout == min_timeout:
                    
            preview = reliable_fetch(1)
           
            peekEmail = '0'

        else:
            preview = reliable_fetch(7)

            print("")
            print("")
            
            for index, msg in enumerate(preview):
                show_message(index, msg)
        
            print("")
            print("")
            
            peekEmail = input_for_mode_selection('Pick an email to sort. ', '0')
            
            if (exit_command(peekEmail) or peekEmail == ''):
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
            
            server = refresh_connection('INBOX')
            continue


        FOLDERSTACK = determinefolder(selectedEmail)
        FULLPATH = createfolder(FOLDERSTACK, len(FILTERED_UIDS))                
        
        try:
            
            moveemails(server, FULLPATH, FILTERED_UIDS)
            
            counting = len(FILTERED_UIDS)
            runtimecount = runtimecount + counting
            
            speakline("  Emails sent in " + yearX  + " from " + fromX , counting)
            speakline("Total emails sorted", runtimecount)


        except Exception as e:
            speakline('Failed to move emails', str(e))




# ---------------
# ---------------

min_timeout = 5
max_timeout = 300
dynamic_timeout = max_timeout
mailbox_server = None
uptime_tracker = datetime.now()
os.chdir(pathlib.Path(__file__).parent)

while True:

    screen_clear()
    
    println('Started At', uptime_tracker.strftime("%I:%M%p"))
    println('Now', datetime.now().strftime("%I:%M%p"))
    print_divider()
    
    println('Press A', 'Run (A)ll day and keep inbox sorted')
    println('Press S', 'Automatically sort emails in your inbox to subfolders.')
    println('Press M', 'Empty out select subfolders.')
    println('Press C', 'Cleanup mailbox.  Delete empty subfolders.')
    println('Press P', 'Prioritize senders')
    println('Press R', 'Read emails in your inbox or other folder.')
    println('Press SL', 'Sit and Listen.  Automatically read and delete emails in your inbox or other folder.')
    println('Press L', 'Fill up a player list queue')
    print_divider()
    println('Press X', 'To quit')

    mode_selection = spokeninputtimeout('Select a mode: ', 's', (4*60*60)).upper()

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

    elif mode_selection == 'L':

        folders = folderselection()

        for f in folders:
            mode_queue(f.name)

    elif mode_selection == 'P':

        folders = folderselection()

        for f in folders:
            mode_prioritize(f.name)

    elif mode_selection == 'R' or mode_selection == 'SL':

        folders = folderselection()
        
        if mode_selection == 'R':
            inputcontrols('')
        else:
            inputcontrols()

        for f in folders:
            response = mode_read(f.name)
            
            if exit_command(response):
                break

    elif mode_selection == 'M':
        
        folders = folderselection()
        mode_move(folders)

    elif mode_selection == 'S':

        inputcontrols('20')
        
        mode_sort()
        
    elif exit_command(mode_selection):
    
        break

        

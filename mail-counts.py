import imaplib
import json


filename = input('Which config to load? ')

if filename == '':
    filename = 'config'

configs = json.load(open('./'+filename+'.json', 'r'))
in_total = 0

################ IMAP SSL ##############################

with imaplib.IMAP4_SSL(host=configs['host'], port=imaplib.IMAP4_SSL_PORT) as imap_ssl:
    print("Connection Object:       {}".format(imap_ssl))

    ############### Login to Mailbox ######################
    print("Logging into mailbox:   ", configs['host'])
    resp_code, response = imap_ssl.login(configs['user'], configs['pass'])

    print("Login Result:            {}".format(resp_code))
    print("Response:                {}".format(response[0].decode()))

    #################### List Directores #####################
    resp_code, directories = imap_ssl.list(pattern='"' + input('Search pattern: ') + '"')

    print("Fetch List:              {}".format(resp_code))
    print("List Count:              {}".format(len(directories)))

    ############### Number of Messages per Directory ############
    print("\n=========== Mail Count Per Directory ===============\n")
    for directory in directories:
        directory_name = directory.decode().split('"')[-2]
        directory_name = '"' + directory_name + '"'
        if directory_name == '"[Gmail]"':
            continue
        try:
            resp_code, mail_count = imap_ssl.select(mailbox=directory_name, readonly=True)
            print("{} - {}".format(directory_name, mail_count[0].decode()))
            in_total = in_total + int(mail_count[0].decode())
        except Exception as e:
            print("{} - ErrorType : {}, Error : {}".format(directory_name, type(e).__name__, e))
            resp_code, mail_count = None, None

    ############# Close Selected Mailbox #######################
    imap_ssl.close()
    

print("Total Emails:             {}".format(in_total))
print("\n\n")
input('Press enter key to exit')
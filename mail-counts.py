import imaplib
import json

configs = json.load(open('./config.json', 'r'))

v_host = configs['host']
v_user = configs['user']
v_pass = configs['pass']

################ IMAP SSL ##############################

with imaplib.IMAP4_SSL(host=v_host, port=imaplib.IMAP4_SSL_PORT) as imap_ssl:
    print("Connection Object:       {}".format(imap_ssl))

    ############### Login to Mailbox ######################
    print("Logging into mailbox:   ", v_host)
    resp_code, response = imap_ssl.login(v_user, v_pass)

    print("Login Result:            {}".format(resp_code))
    print("Response:                {}".format(response[0].decode()))

    #################### List Directores #####################
    resp_code, directories = imap_ssl.list()

    print("Fetch List:              {}".format(resp_code))

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
        except Exception as e:
            print("{} - ErrorType : {}, Error : {}".format(directory_name, type(e).__name__, e))
            resp_code, mail_count = None, None

    ############# Close Selected Mailbox #######################
    imap_ssl.close()
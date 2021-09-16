import imaplib
import email
import traceback
import numpy as np
import pandas as pd
from tqdm import trange
import dateutil
from datetime import datetime
import email.header
import argparse
import configparser

def decode_mime_words(s):
    return u''.join(
        word.decode(encoding or 'utf8') if isinstance(word, bytes) else word
        for word, encoding in email.header.decode_header(s))
    
def list_folders():
    try:
        print(f"Retreiving folders for {FROM_EMAIL}")
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        for f in mail.list()[1]:
            print(f)
    except:
        raise ConnectionError("Unable to retreive folder list")

def read_email_from_gmail(search_criteria=None, folder=None):
    subject = []
    received = []
    from_email = []
    folder_cat = []
    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        if folder == None:
            folder = 'inbox'
        mail.select(folder, readonly=True)

        if search_criteria:
            search = mail.search(None, f'({search_criteria})')
        else:
            search = mail.search(None, "ALL")
        mail_ids = search[1]
        id_list = mail_ids[0].split()   
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        for i in trange(latest_email_id,first_email_id, -1):
            data = mail.fetch(str(i), '(RFC822)' )
            for response_part in data:
                arr = response_part[0]
                if isinstance(arr, tuple):
                    try:
                        msg = email.message_from_string(str(arr[1],'latin-1'))
                        email_subject = decode_mime_words(msg['subject'])
                        email_from = decode_mime_words(msg['from'])
                        email_received = decode_mime_words(msg['Received'].split(";")[1].strip())
                        # email_content = decode_mime_words(msg['Content-Type'])
                        email_received = dateutil.parser.parse(email_received).astimezone()
                        # email_content = msg['Content-Type']
                        subject.append(email_subject)
                        received.append(email_received)
                        from_email.append(email_from)
                        folder_cat.append(folder)
                        # content.append(email_content)
                    except Exception:
                        pass
                    # print('From : ' + email_from + '\n')
                    # print('Subject : ' + email_subject + '\n')
                    # print('Received : ' + email_received + '\n')
                    
        # return pd.DataFrame(np.column_stack([received, from_email, subject,content]),columns=['Received Date','From','Email Subject','Message Content'])
        return pd.DataFrame(np.column_stack([received, from_email, subject,folder_cat]),columns=['Received Date','From','Email Subject','Folder'])
    except KeyboardInterrupt:
        print('Interrupted')
        try:
            # return pd.DataFrame(np.column_stack([received, from_email, subject,content]),columns=['Received Date','From','Email Subject','Message Content'])
            return pd.DataFrame(np.column_stack([received, from_email, subject,folder_cat]),columns=['Received Date','From','Email Subject','Folder'])
            # sys.exit(0)
        except SystemExit:
            # os._exit(0)
            pass
    except Exception as e:
        
        traceback.print_exc() 
        print(str(e))
        # return pd.DataFrame(np.column_stack([received, from_email, subject,content]),columns=['Received Date','From','Email Subject','Message Content'])
        return pd.DataFrame(np.column_stack([received, from_email, subject,folder_cat]),columns=['Received Date','From','Email Subject','Folder'])
        # pass

if __name__ == "__main__":
    # read_email_from_gmail()
    parser = argparse.ArgumentParser()
    parser.add_argument("-d", type=str, default="gmail",help="Email to search")
    parser.add_argument("-s",type=str,help="Specify the search criteria. eg. since '14-Sep-2021' before '30-Sep-2021'")
    parser.add_argument("-f",type=str,default=None, help="Email folder to retrive the emails from.")
    args = parser.parse_args()
    search_criteria = args.s
    email_id = args.d.upper()
    folder = args.f
    config = configparser.ConfigParser()
    config.read('settings.ini')
    
    FROM_EMAIL = config.get(email_id,"FROM_EMAIL")
    FROM_PWD = config.get(email_id,"FROM_PWD")
    SMTP_SERVER = config.get(email_id,"SMTP_SERVER")
    SMTP_PORT = config.get(email_id,"SMTP_PORT")

    if folder:
        final = read_email_from_gmail(search_criteria, folder)
        time_now = datetime.now().strftime("%m-%d-%Y %H-%M-%S")
        final.to_csv(f"Output - {time_now}.csv", index=False)
    else:
        list_folders()
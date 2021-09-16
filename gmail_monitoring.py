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

def read_email_from_gmail(search_criteria=None):
    subject = []
    received = []
    from_email = []
    content = []
    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox', readonly=True)

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
                        # content.append(email_content)
                    except Exception:
                        pass
                    # print('From : ' + email_from + '\n')
                    # print('Subject : ' + email_subject + '\n')
                    # print('Received : ' + email_received + '\n')
                    
        # return pd.DataFrame(np.column_stack([received, from_email, subject,content]),columns=['Received Date','From','Email Subject','Message Content'])
        return pd.DataFrame(np.column_stack([received, from_email, subject]),columns=['Received Date','From','Email Subject'])
    except KeyboardInterrupt:
        print('Interrupted')
        try:
            # return pd.DataFrame(np.column_stack([received, from_email, subject,content]),columns=['Received Date','From','Email Subject','Message Content'])
            return pd.DataFrame(np.column_stack([received, from_email, subject]),columns=['Received Date','From','Email Subject'])
            # sys.exit(0)
        except SystemExit:
            # os._exit(0)
            pass
    except Exception as e:
        
        traceback.print_exc() 
        print(str(e))
        # return pd.DataFrame(np.column_stack([received, from_email, subject,content]),columns=['Received Date','From','Email Subject','Message Content'])
        return pd.DataFrame(np.column_stack([received, from_email, subject]),columns=['Received Date','From','Email Subject'])
        # pass

if __name__ == "__main__":
    # read_email_from_gmail()
    parser = argparse.ArgumentParser()
    parser.add_argument("-d", type=str, default="gmail",help="Email to search")
    parser.add_argument("-s",type=str,help="Specify the search criteria. eg. since '14-Sep-2021' before '30-Sep-2021'")
    args = parser.parse_args()
    search_criteria = args.s
    email_id = args.d.upper()
    config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation())
    config.read('settings.ini')
    
    FROM_EMAIL = config.get(email_id,"FROM_EMAIL")
    FROM_PWD = config.get(email_id,"FROM_PWD")
    SMTP_SERVER = config.get(email_id,"SMTP_SERVER")
    SMTP_PORT = config.get(email_id,"SMTP_PORT")

    final = read_email_from_gmail(search_criteria)
    time_now = datetime.now().strftime("%m-%d-%Y %H-%M-%S")
    final.to_csv(f"Output - {time_now}.csv", index=False)
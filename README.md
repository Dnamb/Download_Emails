# Download_Emails
Download emails from Gmail, Outlook, IMAP enabled email services.

Setup Instructions:
1. Download python from python.org (3.9 or newer)
2. Download the script files and extract into a folder.
3. Open terminal/powershell/cmd in the extracted folder.
4. Run "pip install -r Requirements.txt" command.
5. Open the settings.ini file and update the GMAIL and/or OUTLOOK email and password.

Usage Intructions:
1. Once the basic setup is performed, run "python gmail_monitoring.py -d gmail" command to download the list of all emails on your gmail inbox.
2. Run "python gmail_monitoring.py -d gmail -f spam" command to get list of emails in spam folder. (Note: Some email providers have "Junk" folder instead of spam)
3. Run "python gmail_monitoring.py -d outlook" command to download the list of all emails on your outlook inbox.
4. Run "python gmail_monitoring.py -d gmail -s 'since "1-Sept-2021" before "17-Sept-2021"' command to get emails between a specific time period.

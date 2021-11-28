import imaplib
import os
from dotenv import load_dotenv
load_dotenv()

imap_host = "outlook.office365.com"
imap_user = os.getenv("EMAIL")
imap_pass = os.getenv("PASSWORD")

# connect to host using SSL
imap = imaplib.IMAP4_SSL(imap_host)

## login to server
imap.login(imap_user, imap_pass)

imap.select("Inbox")


# Get the latest email details:
tmp, data = imap.search(None, "ALL")
num=data[0].split()[len(data[0].split())-1].decode('UTF-8')
print(num, end=" ")
tmp, data = imap.fetch(num, "(RFC822)")
print("Message: {0}, status: {1}".format(num, tmp))
print(data[0][1].decode('UTF-8'))


# Get all emails till now:
tmp, data = imap.search(None, "ALL")
for num in data[0].split():
    num=num.decode('UTF-8')
    print(num, end=" ")
    tmp, data = imap.fetch(num, "(RFC822)")
    print("Message: {0}, status: {1}".format(num, tmp))
    print(data[0][1].decode('UTF-8'))
      

imap.close()
imap.logout()

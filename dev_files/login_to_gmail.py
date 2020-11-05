# https://www.thepythoncode.com/article/reading-emails-in-python

import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import html2text

# account credentials
username = "toj320@gmail.com"
password = r"2rhD@o9IHNJ$"

imap = imaplib.IMAP4_SSL("imap.gmail.com")
# authenticate
imap.login(username, password)

status, messages = imap.select("INBOX")
# number of top emails to fetch
N = 3
# total number of emails
messages = int(messages[0])



for i in range(messages, messages-N, -1):
    # fetch the email message by ID
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            # decode the email subject
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                # if it's a bytes, decode to str
                subject = subject.decode()
            # email sender
            from_ = msg.get("From")
            print("Subject:", subject)
            print("From:", from_)
            # if the email message is multipart
            if msg.is_multipart():
                # iterate over email parts
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # print text/plain emails and skip attachments
                        print(body)
                    elif "attachment" in content_disposition:
                        # download attachment
                        filename = part.get_filename()
                        if filename:
                            if not os.path.isdir(subject):
                                # make a folder for this email (named after the subject)
                                os.mkdir(subject)
                            filepath = os.path.join(subject, filename)
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                # get the email body
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    # print only text email parts
                    print(body)
            if content_type == "text/html":
                # if it's HTML, create a new HTML file and open it in browser
                if not os.path.isdir(subject):
                    # make a folder for this email (named after the subject)
                    os.mkdir(subject)
                filename = f"{subject[:50]}.html"
                filepath = os.path.join(subject, filename)
                # write the file
                open(filepath, "w").write(body)
                # open in the default browser
                webbrowser.open(filepath)
            print("="*100)
imap.close()
imap.logout()

"""
example output
Subject: Thanks for Subscribing to our Newsletter !
From: rockikz@thepythoncode.com
====================================================================================================
Subject: An email with a photo as an attachment
From: Python Code <rockikz@thepythoncode.com>
Get the photo now!

====================================================================================================
Subject: A Test message with attachment
From: Python Code <rockikz@thepythoncode.com>
There you have it!

====================================================================================================
"""

[(b'13705 (RFC822 {2122}',
  b'Return-Path: <toj320@gmail.com>\r\nReceived: from smtp.gmail.com (096-032-002-083.res.spectrum.com. [96.32.2.83])\r\n        by smtp.gmail.com with ESMTPSA id 6sm2462811qtz.31.2020.10.28.04.03.21\r\n        for <toj320@gmail.com>\r\n        (version=TLS1_2 cipher=ECDHE-ECDSA-AES128-GCM-SHA256 bits=128/128);\r\n        Wed, 28 Oct 2020 04:03:21 -0700 (PDT)\r\nMIME-Version: 1.0\r\nDate: Wed, 28 Oct 2020 07:03:19 -0400\r\nFrom: Thomas Francis <toj320@gmail.com>\r\nSubject: sdfawe\r\nThread-Topic: sdfawe\r\nMessage-ID: <7CEAF142-DBE5-479D-A041-528C124A5031@hxcore.ol>\r\nTo: Thomas Francis <toj320@gmail.com>\r\nContent-Transfer-Encoding: quoted-printable\r\nContent-Type: text/html; charset="utf-8"\r\n\r\n<html xmlns:o=3D"urn:schemas-microsoft-com:office:office" xmlns:w=3D"urn:sc=\r\nhemas-microsoft-com:office:word" xmlns:m=3D"http://schemas.microsoft.com/of=\r\nfice/2004/12/omml" xmlns=3D"http://www.w3.org/TR/REC-html40"><head><meta ht=\r\ntp-equiv=3DContent-Type content=3D"text/html; charset=3Dutf-8"><meta name=\r\n=3DGenerator content=3D"Microsoft Word 15 (filtered medium)"><style><!--\r\n/* Font Definitions */\r\n@font-face\r\n\t{font-family:"Cambria Math";\r\n\tpanose-1:2 4 5 3 5 4 6 3 2 4;}\r\n@font-face\r\n\t{font-family:Calibri;\r\n\tpanose-1:2 15 5 2 2 2 4 3 2 4;}\r\n@font-face\r\n\t{font-family:"Malgun Gothic";\r\n\tpanose-1:2 11 5 3 2 0 0 2 0 4;}\r\n@font-face\r\n\t{font-family:"\\@Malgun Gothic";}\r\n/* Style Definitions */\r\np.MsoNormal, li.MsoNormal, div.MsoNormal\r\n\t{margin:0in;\r\n\tfont-size:11.0pt;\r\n\tfont-family:"Calibri",sans-serif;}\r\na:link, span.MsoHyperlink\r\n\t{mso-style-priority:99;\r\n\tcolor:blue;\r\n\ttext-decoration:underline;}\r\n.MsoChpDefault\r\n\t{mso-style-type:export-only;}\r\n@page WordSection1\r\n\t{size:8.5in 11.0in;\r\n\tmargin:1.0in 1.0in 1.0in 1.0in;}\r\ndiv.WordSection1\r\n\t{page:WordSection1;}\r\n--></style></head><body lang=3DEN-US link=3Dblue vlink=3D"#954F72"><div cla=\r\nss=3DWordSection1><p class=3DMsoNormal>aseaef</p><p class=3DMsoNormal><o:p>=\r\n&nbsp;</o:p></p><p class=3DMsoNormal>Sent from <a href=3D"https://go.micros=\r\noft.com/fwlink/?LinkId=3D550986">Mail</a> for Windows 10</p><p class=3DMsoN=\r\normal><o:p>&nbsp;</o:p></p></div></body></html>=\r\n\r\n'),
 b' FLAGS (\\Seen))']

res, msg = imap.fetch(str(messages), "(RFC822)")
email_obj = email.message_from_bytes(msg[0][1])

email.message.EmailMessage().as_string(msg)

#email_obj.get_payload(decode=False).decode()
email_obj.get_payload(decode=False) # returns normal string format
email_obj.get_payload(decode=True) # returns byte string format
email_obj.get_payload(decode=True).decode() # returns normal string format

body = email_obj.get_payload(decode=True).decode()

'<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=utf-8"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><style><!--\r\n/* Font Definitions */\r\n@font-face\r\n\t{font-family:"Cambria Math";\r\n\tpanose-1:2 4 5 3 5 4 6 3 2 4;}\r\n@font-face\r\n\t{font-family:Calibri;\r\n\tpanose-1:2 15 5 2 2 2 4 3 2 4;}\r\n@font-face\r\n\t{font-family:"Malgun Gothic";\r\n\tpanose-1:2 11 5 3 2 0 0 2 0 4;}\r\n@font-face\r\n\t{font-family:"\\@Malgun Gothic";}\r\n/* Style Definitions */\r\np.MsoNormal, li.MsoNormal, div.MsoNormal\r\n\t{margin:0in;\r\n\tfont-size:11.0pt;\r\n\tfont-family:"Calibri",sans-serif;}\r\na:link, span.MsoHyperlink\r\n\t{mso-style-priority:99;\r\n\tcolor:blue;\r\n\ttext-decoration:underline;}\r\n.MsoChpDefault\r\n\t{mso-style-type:export-only;}\r\n@page WordSection1\r\n\t{size:8.5in 11.0in;\r\n\tmargin:1.0in 1.0in 1.0in 1.0in;}\r\ndiv.WordSection1\r\n\t{page:WordSection1;}\r\n--></style></head><body lang=EN-US link=blue vlink="#954F72"><div class=WordSection1><p class=MsoNormal>aseaef</p><p class=MsoNormal><o:p>&nbsp;</o:p></p><p class=MsoNormal>Sent from <a href="https://go.microsoft.com/fwlink/?LinkId=550986">Mail</a> for Windows 10</p><p class=MsoNormal><o:p>&nbsp;</o:p></p></div></body></html>\r\n'

h = html2text.html2text()
h.ignore_links = True

print(html2text.html2text(body))

"""
CSV to Python
"""

import csv



with open('persons.csv', 'w') as csvfile:
    filewriter = csv.DictWriter(csvfile)
    filewriter.writerows(
        )

https://www.geeksforgeeks.org/writing-csv-files-in-python/

import csv  
    
# field names  
fields = ['Name', 'Branch', 'Year', 'CGPA']  
    
# data rows of csv file  
rows = [ ['Nikhil', 'COE', '2', '9.0'],  
         ['Sanchit', 'COE', '2', '9.1'],  
         ['Aditya', 'IT', '2', '9.3'],  
         ['Sagar', 'SE', '1', '9.5'],  
         ['Prateek', 'MCE', '3', '7.8'],  
         ['Sahil', 'EP', '2', '9.1']]  
    
# name of csv file  
filename = "university_records.csv"
    
# writing to csv file  
with open('names.csv', 'w', newline='') as csvfile:  
    # creating a csv writer object  
    csvwriter = csv.writer(csvfile)  
        
    # writing the fields  
    csvwriter.writerow(fields)  
        
    # writing the data rows  
    csvwriter.writerows(rows) 
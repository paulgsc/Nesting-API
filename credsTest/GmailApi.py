#!/usr/bin/env python
# coding: utf-8

# In[1]:


from credsTest.Google import Create_Service
import os
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import mimetypes


# In[4]:


CLIENT_SECRET_FILE = 'client_secret_file.json'
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://mail.google.com/']


# In[5]:


service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)


# In[22]:


def sendEmail(msg_string, recepient, subject, template='html'):
    emailMsg = msg_string
    mimeMessage = MIMEMultipart()
    mimeMessage['to'] = recepient
    mimeMessage['subject'] = subject
    mimeMessage.attach(MIMEText(emailMsg, template))
    raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
    message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
    if message['labelIds']== 'SENT':
        print('Email Successfully Sent')


# In[9]:





# In[24]:


# HTML message, would use mako templating in real scenario
msg_html = """
<!DOCTYPE html>
<html>
<head><style type="text/css">
.attribution {{ color: #aaaaaa; font-size: 8pt }}
.greeting {{ font-size: 14pt; font-styweight: bold}}
</style></head>
<body>
<span class="greeting">Hello, {}!</span>
<p>As our valued customer, we would like to invite you to our annual sale!</p>
<p class="attribution">
<a href="https://www.youtube.com/watch?v=JRCJ6RtE3xU&ab_channel=CoreySchafer">
Image by FreeVector.com
</a></p>
</body></html>
"""


# In[25]:


sendPlainTextEmail(msg_html, 'paulgathondudev@gmail.com', 'testing html')


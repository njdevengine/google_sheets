import win32com.client
import pythoncom
import re

messages = []

ol = win32com.client.Dispatch( "Outlook.Application")
inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
for message in inbox.Items:
    if message.UnRead == True:
        if "Partners Page Form - Website Submission" in message.Subject:
            #print(message.Subject)
            messages.append(message)

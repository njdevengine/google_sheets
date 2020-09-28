import win32com.client
import pythoncom
import re
import pandas as pd

messages = []

ol = win32com.client.Dispatch( "Outlook.Application")
inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
for message in inbox.Items:
    if message.UnRead == True:
        if "Partners Page Form - Website Submission" in message.Subject:
            #print(message.Subject)
            messages.append(message)
            
words = ["Name:","Company Name:","Email:","Phone Number:","How did you hear about us?:"]

data = []
for m in range(len(messages)):
    for i in range(len(words)):
        try:
            begin = messages[m].Body.find(words[i])
            end = messages[m].Body.find(words[i+1])
            length = len(words[i])
            data.append(messages[m].Body[(begin+length):end])
        except:
            begin = messages[m].Body.find(words[i])
            end = messages[m].Body.find("Reply to customer")
            length = len(words[i])
            data.append(messages[m].Body[(begin+length):end])
            
names = []
companies = []
emails = []
phones = []
sources = []
dates = []

for m in range(len(messages)):
    begin = messages[m].Body.find("Name:")
    end = messages[m].Body.find("Company Name:")
    length = len("Name:")
    name = messages[m].Body[(begin+length):end].strip().title()
    if len(name) <=100:
        names.append(name)
    
    begin = messages[m].Body.find("Company Name:")
    end = messages[m].Body.find("Email:")
    length = len("Company Name:")
    company = messages[m].Body[(begin+length):end].strip()
    if len(company) <=100:
        companies.append(company)
    
    begin = messages[m].Body.find("Email:")
    end = messages[m].Body.find("Phone Number:")
    length = len("Email:")
    email = messages[m].Body[(begin+length):end].strip()
    if len(email) <=100:
        emails.append(email)
    
    begin = messages[m].Body.find("Phone Number:")
    end = messages[m].Body.find("How did you hear about us?:")
    length = len("Phone Number:")
    phone = messages[m].Body[(begin+length):end].strip().split(" <")[0]
    if len(phone) <=100:
        phones.append(phone)
    
    begin = messages[m].Body.find("How did you hear about us?:")
    end = messages[m].Body.find("Reply to customer")
    length = len("How did you hear about us?:")
    source = messages[m].Body[(begin+length):end].strip()
    if len(source) <=100:
        sources.append(source)
for i in messages:
    try:
        date = i.SentOn.strftime("%m/%d/%Y")
        dates.append(date)
    except: pass
    
df2 = pd.DataFrame({"Name":names,"Company":companies,"Email":emails,"Phone":phones,"Source":sources})

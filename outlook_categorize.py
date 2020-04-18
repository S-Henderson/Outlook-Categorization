import os

##############################
desktopPath = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')

os.chdir(desktopPath)

get_path = os.getcwd()

import win32com.client

import time


from win32com.client import Dispatch


outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

root_folder = outlook.Folders.Item("RAC Reporting") # Change number in brackets for folders change

print (root_folder.Name)

your_folder = root_folder.Folders['Inbox']

for message in your_folder.Items:
    print(message.Subject)
    
messages = your_folder.Items

print("There Are " + str(messages.Count) + " Emails In This Folder")

named_category = 'Scotty'

from datetime import datetime

now = datetime.now() # Current Datetime

today = now.strftime("%m-%d-%Y")
today2 = '{dt.month}-{dt.day}-{dt.year}'.format(dt = datetime.now()) # Gets rid of leading zero in day

strings = ["[Data Report] Please find attached the RAC CVI Consumer Check v2 Report for 4-16-2020",
                 "[Data Report] Please find attached the RAC Different Dates Report for 4-16-2020"]

new_strings = []

for string in strings:
    new_string = string.replace("4-16-2020", today2)
#Modify old string

    new_strings.append(new_string)
#Add new string to list


print(new_strings)

# Want to change category to "Scotty" & download attachment
for i in range(messages.Count):
    if messages[i].Subject in new_strings:
        messages[i].GetInspector()
        messages[i].Categories = named_category # Change category marker name here
        messages[i].Save()
        
        for x in messages[i].Attachments:
            x.SaveAsFile(os.path.join(get_path,x.FileName))
            print("successfully downloaded attachments")
            print(messages[i].Attachments.Item(1).DisplayName)

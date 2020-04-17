import win32com.client

import time


from win32com.client import Dispatch


outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

root_folder = outlook.Folders.Item("shenderson@360insights.com") # Change number in brackets for folders change

print (root_folder.Name)

########################

your_folder = root_folder.Folders['Test']

for message in your_folder.Items:
    print(message.Subject)
    
messages = your_folder.Items

print(messages.Count)

########################

my_list = ["[Data Report] Please find attached the RAC CVI Consumer Check v2 Report for 4-16-2020",
          "[Data Report] Please find attached the RAC Different Dates Report for 4-16-2020"]
		  
#######################################

import os

##############################
desktopPath = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')

os.chdir(desktopPath)

get_path = os.getcwd()
##############################

named_category = 'Scotty'

#############################

# Want to change category to "Scotty" & download attachment
for i in range(messages.Count):
    if messages[i].Subject in my_list:
        messages[i].GetInspector()
        messages[i].Categories = named_category # Change category marker name here
        messages[i].Save()
        
        for x in messages[i].Attachments:
            x.SaveASFile(os.path.join(get_path,x.FileName))
            print("successfully downloaded attachments")
            print(messages[i].Attachments.Item(1).DisplayName)

input("Press enter to exit :)")        
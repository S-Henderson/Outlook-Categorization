"""
By: Scott Henderson
Last Updated: Apr 22, 2020
Purpose: In Outlook find reports to categorize and download attachments based on Subject line search
"""

import os

from win32com.client import Dispatch

#--------------- PURPOSE ---------------

print("Purpose: In Outlook find reports to categorize and download attachments based on Subject line search")

print("*************************")

#--------------- ASCII ART ---------------
print(r"""
________          __  .__                 __     _________         __                             .__               
\_____  \  __ ___/  |_|  |   ____   ____ |  | __ \_   ___ \_____ _/  |_  ____   ____   ___________|__|_______ ____  
 /   |   \|  |  \   __\  |  /  _ \ /  _ \|  |/ / /    \  \/\__  \\   __\/ __ \ / ___\ /  _ \_  __ \  \___   // __ \ 
/    |    \  |  /|  | |  |_(  <_> |  <_> )    <  \     \____/ __ \|  | \  ___// /_/  >  <_> )  | \/  |/    /\  ___/ 
\_______  /____/ |__| |____/\____/ \____/|__|_ \  \______  (____  /__|  \___  >___  / \____/|__|  |__/_____ \\___  >
        \/                                    \/         \/     \/          \/_____/                       \/    \/ 
""")

#--------------- SETUP ---------------

# Set download path for report files
file_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads') # Change 2nd string to change final location
os.chdir(file_path)

print("The Download File Path is -> " + file_path)

print("*************************")

#--------------- OUTLOOK CONNECTION ---------------

# Connect to Outlook
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

# Set main folder
root_folder = outlook.Folders.Item("RAC Reporting") # Change string/number in brackets for root folder change (eg shenderson@360insights.com)

print ("The Email Root Folder is -> " + root_folder.Name)

print("*************************")

#--------------- OUTLOOK FOLDER SELECT ---------------

# Set sub folder -> usually just Inbox
sub_folder = root_folder.Folders['Inbox']

# set email object
messages = sub_folder.Items

print("There Are " + str(messages.Count) + " Emails In This Folder")

""" 
Optional -> print all email subjects in sub_folder

for message in your_folder.Items:
    print(message.Subject)
"""

#--------------- SET CATEGORY VALUE ---------------

# Change Outlook category marker name here
named_category = 'Scotty'

#--------------- CATEGORIZE & DOWNLOAD ---------------

print("*************************")

# Loop to find report emails to categorize & download
for i in range(messages.Count):
    
    # Manually insert partial name match of report to find -> based on Subject line search
    if "RAC Different Dates Report" in messages[i].Subject or \
       "All Clients Duplicate Serial Report" in messages[i].Subject or \
       "RAC CVI Consumer Check v2 Report" in messages[i].Subject or \
       "RAC HVAC Cross Module Compliance" in messages[i].Subject or \
       "Lennox Dup Serial Exception Report" in messages[i].Subject or \
       "Lennox Duplicate Serial report" in messages[i].Subject or \
       "RAC Lennox Potential Over Payments report" in messages[i].Subject:

        # Categorize
        messages[i].GetInspector()
        messages[i].Categories = named_category 
        messages[i].Save()
        
        print("Categorizing Email for -> " + messages[i].Subject)
        
        # Download attachments
        for attachment in messages[i].Attachments:
            attachment.SaveAsFile(os.path.join(file_path, attachment.FileName))
            
            print("Downloading Attachments for -> " + messages[i].Attachments.Item(1).DisplayName)

#--------------- ENDING ---------------

print("*************************")

print("Script Completed")

print("*************************")

print("Have A Great Day! Here Are Some Cats!")

print(r"""
   /\     /\
  {  `---'  }
  {  O   O  }
  ~~>  V  <~~
   \  \|/  /
    `-----'__
    /     \  `^\_
   {       }\ |\_\_   W
   |  \_/  |/ /  \_\_( )
    \__/  /(_E     \__/
      (  /
       MM
""")

print(r"""
        _,'|             _.-''``-...___..--';)
       /_ \'.      __..-' ,      ,--...--'''
      <\    .`--'''       `     /'
       `-';'               ;   ; ;
 __...--''     ___...--_..'  .;.'
(,__....----'''       (,..--''
""")

input("Press enter to exit :)")

"""
By: Scott Henderson
Last Updated: Apr 22, 2020
Purpose: In Outlook find reports to categorize and downloads attachments based on Subject line search
"""

import os
import re
from win32com.client import Dispatch

#--------------- PURPOSE ---------------

print("Purpose: In Outlook find reports to categorize and downloads attachments based on Subject line search")

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

# Change working directory to download path
os.chdir(file_path)

print("The Download File Path is -> " + file_path)

print("*************************")

#--------------- OUTLOOK CONNECTION ---------------

# Connect to Outlook
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

# Set main folder
root_folder = outlook.Folders.Item("RAC Reporting") # Change string/number in brackets for root folder change (eg xyz@360insights.com)

print ("The Email Root Folder is -> " + root_folder.Name)

print("*************************")

#--------------- OUTLOOK FOLDER SELECT ---------------

# Set sub folder -> usually just Inbox
sub_folder = root_folder.Folders['Inbox']

print("The Email Sub Folder is -> " + sub_folder.Name)

# set Email object
messages = sub_folder.Items

print("There Are " + str(messages.Count) + " Emails In This Folder")

# Set emails (plural) to all emails in sub_folder
emails = range(messages.Count)

""" 
Optional -> print all email subjects in sub_folder

for message in your_folder.Items:
    print(message.Subject)
"""

#--------------- SET CATEGORY VALUE ---------------

# Change Outlook category marker name here
named_category = 'Scotty'

#--------------- REPORT LIST ---------------

# Report List
report_list =["RAC Different Dates Report",                         # Main Report
              "All Clients Duplicate Serial Report",                # Main Report
              "RAC CVI Consumer Check v2 Report",                   # CVI Report
              "RAC HVAC Cross Module Compliance",                   # Lennox Report
              "Lennox Dup Serial Exception Report",                 # Lennox Report
              "Lennox Duplicate Serial report",                     # Lennox Report
              "RAC Lennox Potential Over Payments report"]          # Lennox Report

# Compile a regular expression pattern into a regular expression object, which can be used for matching
# Source -> https://stackoverflow.com/questions/6750240/how-to-do-re-compile-with-a-list-in-python/6750274#6750274          
report_str = re.compile(r'\b(?:%s)\b' % '|'.join(report_list))     

#--------------- CATEGORIZE & DOWNLOAD ---------------

print("*************************")

"""
needed_email = [messages[email].Subject for email in emails if re.findall(report_str, messages[email].Subject)]

print(needed_email)

for i in needed_email:
    
    i.GetInspector()
    i.Categories = named_category 
    i.Save()
    
    print("Successfully Actioned   -> " + i)
    

"""

# Loop to find report emails to categorize & download
for email in emails:
    
    # Searches list of reports and finds all hits in email Subject Line
    if re.findall(report_str, messages[email].Subject):
    
        # Categorize
        messages[email].GetInspector()
        messages[email].Categories = named_category 
        messages[email].Save()
        
        print("Successfully Categorized -> " + messages[email].Subject)

        # Download attachments
        for attachment in messages[email].Attachments:
            attachment.SaveAsFile(os.path.join(file_path, attachment.FileName))

            print("Successfully Downloaded  -> " + messages[email].Attachments.Item(1).DisplayName)
            
#--------------- ENDING ---------------

print("*************************")

print("Script Completed")
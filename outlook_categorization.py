""" 
############################################################
Name         : Outlook Categorization
Author       : Scott Henderson
Created      : 10/10/2020
Last Updated : 4/28/2021
Purpose      : In Outlook, find reports to categorize and download attachment based on email Subject line search
Input        : Outlook emails in RAC Reporting inbox
Output       : Categorized Outlook emails in RAC reporting inbox & downloaded attachments into local downloads folder
Workflow     : Connect to Outlook inbox, combine reports into string to search, loop through email inbox based on Subject line,
               find matched emails, categorize them and download the attachment
############################################################
"""

import os
import re
from win32com.client import Dispatch

#--------------- ASCII ART ---------------#

print(r"""
________          __  .__                 __     _________         __                             .__               
\_____  \  __ ___/  |_|  |   ____   ____ |  | __ \_   ___ \_____ _/  |_  ____   ____   ___________|__|_______ ____  
 /   |   \|  |  \   __\  |  /  _ \ /  _ \|  |/ / /    \  \/\__  \\   __\/ __ \ / ___\ /  _ \_  __ \  \___   // __ \ 
/    |    \  |  /|  | |  |_(  <_> |  <_> )    <  \     \____/ __ \|  | \  ___// /_/  >  <_> )  | \/  |/    /\  ___/ 
\_______  /____/ |__| |____/\____/ \____/|__|_ \  \______  (____  /__|  \___  >___  / \____/|__|  |__/_____ \\___  >
        \/                                    \/         \/     \/          \/_____/                       \/    \/ 
""")

print("############################################################")

#--------------- PURPOSE ---------------#

print("Purpose: In Outlook, find reports to categorize and download their file attachment based on a Subject line search match")
print("############################################################")

#--------------- REPORT LIST ---------------#

report_list = [
    "RAC CVI Consumer Check v2 Report",  
    "RAC HVAC Cross Module Compliance",
    "Lennox Dup Serial Exception Report",
    "Lennox Duplicate Serial report",
    "RAC Lennox Potential Over Payments report",
    "Alcon Duplicate Serial Report V2",
    "Alcon Invalid Codes",
    "Alcon Invalid Codes 2.0",
    "CVICA Purchase Date to Submission Date",
    "RAC Different Dates Report"
]         
              
# Compile a regular expression pattern into a regular expression object, which can be used for matching
# https://stackoverflow.com/questions/6750240/how-to-do-re-compile-with-a-list-in-python/6750274#6750274    
report_string = re.compile(r'\b(?:%s)\b' % '|'.join(report_list))

#--------------- OUTLOOK CATEGORY VALUE ---------------#

named_category = "Scotty"  

#--------------- SAVE PATH ---------------#

# To save attachments to user Downloads folder
save_path = os.path.join(os.path.expanduser("~"), "Downloads")
os.chdir(save_path)

#--------------- OUTLOOK CONNECTION ---------------#

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI") # Outlook connection
inbox = outlook.Folders.Item("RAC Reporting").Folders['Inbox'] # Folder selection
emails = inbox.Items 

#--------------- CATEGORIZE & DOWNLOAD ---------------#

# https://stackoverflow.com/questions/45442442/os-not-letting-me-save-email-attachment-using-pywin32 

def categorize_and_download_outlook_reports():
    """
    Loops through emails in Outlook inbox to categorize and download the file attachment 
    """
    print("The Download File Path is -> " + save_path)
    print ("The Email Folder is -> " + inbox.Name)
    print("There Are " + str(emails.Count) + " Emails In This Folder")
    print("############################################################")

    for email in list(emails):
        if re.findall(report_string, email.Subject):
            
            # Categorize emails
            email.GetInspector()
            email.Categories = named_category 
            email.Save()
    
            print("Successfully Categorized -> " + email.Subject)
    
            # Check for attachments
            email_attachments = email.Attachments
            print(str(email_attachments.Count) + ' attachments found.')
            
            if email_attachments.Count > 0:
                for i in range(email_attachments.Count):
                    
                    # MS Outlook list indices are 1-based hence the + 1
                    email_attachment = email_attachments.Item(i + 1) 
                    email_attachment.SaveAsFile(os.path.join(save_path, email_attachment.FileName))
                    
                    print("Successfully Downloaded  -> " + str(email_attachment.FileName))
                    print("############################################################")
            
            else:
                print('No Attachment Found To Download.')
                print("############################################################")

# Call function
if __name__ == '__main__':
    categorize_and_download_outlook_reports()

#--------------- SCRIPT COMPLETED ---------------#

#print("------------------------------")
print("Script Completed")

#input("Press Enter to Continue...")

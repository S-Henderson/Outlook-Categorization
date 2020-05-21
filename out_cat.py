"""
By: Scott Henderson
Last Updated: May 20, 2020
Purpose: In Outlook, find reports to categorize and download attachment based on Subject line search
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

#--------------- CONNECTION ---------------

save_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads') # Set download path for report files

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")                 # Outlook connection

#--------------- SELECT FOLDERS ---------------

root_folder = outlook.Folders.Item("RAC Reporting")                            # Set main folder
sub_folder = root_folder.Folders['Inbox']                                      # Set sub folder -> usually just Inbox

#--------------- EMAIL OBJECTS ---------------

messages = sub_folder.Items                                                    # Set email object
emails = range(messages.Count)                                                 # Set emails (plural) to all emails in sub_folder

#--------------- CATEGORY VALUE ---------------

named_category = "Scotty"                                                      # Set Outlook category marker name

#--------------- REPORT LIST ---------------

# Report List
report_list =["RAC Different Dates Report",                # Main Report
              "All Clients Duplicate Serial Report",       # Main Report
              "RAC CVI Consumer Check v2 Report",          # CVI Report
              "RAC HVAC Cross Module Compliance",          # Lennox Report
              "Lennox Dup Serial Exception Report",        # Lennox Report
              "Lennox Duplicate Serial report",            # Lennox Report
              "RAC Lennox Potential Over Payments report"] # Lennox Report

# Compile a regular expression pattern into a regular expression object, which can be used for matching
# Source -> https://stackoverflow.com/questions/6750240/how-to-do-re-compile-with-a-list-in-python/6750274#6750274    
report_string = re.compile(r'\b(?:%s)\b' % '|'.join(report_list))     

#--------------- CATEGORIZE & DOWNLOAD ---------------

def categorize_and_download_outlook_reports():
    """
    Loops through emails in Outlook sub_folder to categorize and download them 
    """
    os.chdir(save_path) # Change working directory to download path
    
    print("The Download File Path is -> " + save_path)
    print("*************************")
    
    print ("The Email Root Folder is -> " + root_folder.Name)
    print("*************************")

    print("The Email Sub Folder is -> " + sub_folder.Name)
    print("*************************")
    
    print("There Are " + str(messages.Count) + " Emails In This Folder")
    print("*************************")
    
    # Loop
    for email in emails:
        
        # Searches list of reports and finds all hits in email Subject Line
        if re.findall(report_string, messages[email].Subject):
        
            # Categorize
            messages[email].GetInspector()
            messages[email].Categories = named_category 
            messages[email].Save()
            
            print("Successfully Categorized -> " + messages[email].Subject)

            # Download attachment
            for attachment in messages[email].Attachments:
                attachment.SaveAsFile(os.path.join(save_path, attachment.FileName))

                print("Successfully Downloaded  -> " + messages[email].Attachments.Item(1).DisplayName)
                
                print("*************************")
        
    else:
        print("No Reports Found")

# Call loop function
categorize_and_download_outlook_reports()

#--------------- SCRIPT COMPLETED ---------------

print("*************************")

print("Script Completed")
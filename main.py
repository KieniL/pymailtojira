# import pymsgbox for displaying a messagebox
import pymsgbox
#import pandas for reading the xlsx File
import pandas as pd
#import the request to check if the URL from the mapping is available
import urllib.request,urllib.parse,urllib.error
#Import the Handler for Outlook
import win32com.client

from collections import Counter
from jira import JIRA

#Create a Messagebox with Yes/No
result = pymsgbox.confirm('Shall we create JIRA Issues from Mail?', 'Outlook to JIRA', ["Yes", 'No'])

#Declare the filepath to the mappingtable
filepath = "Mappingtable.xlsx"


#End the Script if the Selection was NO or None
if result == 'No' or result is None:
    print("End")
    quit()

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts;


#Load file into variable
data = pd.read_excel(filepath)

global inbox
for index, row in data.iterrows():
    try:
        url_requested = urllib.request.urlopen(row['JIRAURL'])
        if 200 == url_requested.code:
            for account in accounts:
                if account.DisplayName == row['MailAccount']:
                    inbox = outlook.Folders(account.DeliveryStore.DisplayName)
                    folders = inbox.Folders
                    for folder in folders:
                        if folder.Name == "Posteingang":
                            messages = folder.Items
                            for message in reversed(messages):
                                rightCategory = False

                                if message.Categories == row['Label']:                               
                                    message.Categories = ""


                                    options = {
                                        'server': row['JIRAURL'],
                                        }
                                    basic_auth=(row['JIRAUser'], row['JiraPW'])
                                    jira = JIRA(server=row['JIRAURL'], basic_auth=basic_auth)
                                    issue_dict = {
                                        'project': {'id': row['ProjectID']},
                                        'summary': message.Subject,
                                        'description': message.Body,
                                        'issuetype': {'name': row['IssueType']},
                                    }
                                    new_issue = jira.create_issue(fields=issue_dict)
                                    message.save()                
                            break                  
    #Except if the URL could not be read
    except urllib.error.URLError as e: print('URL ' + row['JIRAURL'] + ' could not be read')
    #Except a ValueError and prints it
    except ValueError as e: print(e)


#import pandas for reading the xlsx File
import pandas as pd
#import pymsgbox for displaying a messagebox, the request to check if the URL from the mapping is available, the Handler for outlook, the time for the sleep, the custom py for jira
import pymsgbox, urllib.request,urllib.parse,urllib.error, win32com.client, time, sys, jira_py

#impot the os for creating a tempo attachment
import os

def changeMailName(mail, issue, addJIRAKeytoMailName):
    if addJIRAKeytoMailName:
        mail.Subject = str(issue) + "_" + mail.Subject
    return mail


def fileHandler(jiraURL, jiraUSER, jirapw, issue, attachment):
    path = os.getcwd()+ "\\" + attachment.FileName
    attachment.SaveAsFile(path)
    if jira_py.add_attachment(jiraURL, jiraUSER, jirapw, issue, path):
        os.remove(path)
        print("removed")
    
    

#Get Arguments from Batfile
if sys.argv:
    iterateTimeInSeconds = int(sys.argv[1])
    addJIRAKeytoMailName = sys.argv[2]
    mailCounter = int(sys.argv[3])
    desiredFolder = sys.argv[4]


#Create a Messagebox with Yes/No
result = pymsgbox.confirm('Shall we create JIRA Issues from Mail?', 'Outlook to JIRA', ["Yes", 'No'])

#Declare the filepath to the mappingtable
filepath = "Mappingtable.xlsx"

#End the Script if the Selection was NO or None
if result == 'No' or result is None:
    print("End")
    quit()

#Get Outlook from the Computer
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#Get the Outlook Accounts
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts


#Load file into variable
data = pd.read_excel(filepath)


global inbox


while True:
    
    #Iterate over Mappingtable.xlsx
    for index, row in data.iterrows():
        counter = 0
        #Iterate over all Accounts in the account variable
        for account in accounts:
            #Only do Further Stuff if account is like the defined Account in the Mappingtable
            if account.DisplayName == row['MailAccount']:
                #get All Folders from the Account
                inbox = outlook.Folders(account.DeliveryStore.DisplayName)
                folders = inbox.Folders
                for folder in folders:
                    #Check if the Folder is like the searchingFolder
                    if folder.Name == desiredFolder:
                        messages = folder.Items
                        #Iterate over the First 50 Messages from newest to latest
                        for message in reversed(messages):
                            if counter == mailCounter:
                                break
                            #Check if the Category of the Message is like the defined Category in the Mappingtable
                            if message.Categories == row['Label']:
                                try:
                                    #Try to open the URL to check if it is reachable
                                    url_requested = urllib.request.urlopen(row['JIRAURL'])
                                    if 200 == url_requested.code:
                                        #Create JIRA Issue and clear Category if jira Issue was created
                                        new_issue = jira_py.createjiraIssue(row['JIRAURL'], row['JIRAUser'], row['JiraPW'], row['ProjectID'], message.Subject, message.Body, row['IssueType'])
                                        if new_issue:
                                            #Add All Attacments to JIRA Issue if there are any
                                            if message.Attachments:
                                                for attachment in message.Attachments:
                                                    fileHandler(row['JIRAURL'], row['JIRAUser'], row['JiraPW'], new_issue, attachment)
                                            message = changeMailName(message, new_issue, addJIRAKeytoMailName)
                                            message.Categories = ""
                                            message.save()  
                                #Except if the URL could not be read
                                except urllib.error.URLError as e: print('URL ' + row['JIRAURL'] + ' could not be read')
                                #Except a ValueError and prints it
                                except ValueError as e: print(e)
                            counter += 1

    print("Iterate")
    time.sleep(iterateTimeInSeconds)
    



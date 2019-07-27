# import pymsgbox for displaying a messagebox
import pymsgbox
#import pandas for reading the xlsx File
import pandas as pd
#import the request to check if the URL from the mapping is available, the Handler for outlook and the time for the sleep
import urllib.request,urllib.parse,urllib.error, win32com.client, time 
#Import the Handler for Outlook

#import the custom py for jira
import jira_py 

#Create a Messagebox with Yes/No
result = pymsgbox.confirm('Shall we create JIRA Issues from Mail?', 'Outlook to JIRA', ["Yes", 'No'])

#Declare the filepath to the mappingtable
filepath = "Mappingtable.xlsx"
#Set the Folder where the Application should search for the Labels
desiredFolder = "Posteingang"

#Define the Time How Often The Application Should be run, (60 = 1 Minute)
IterateTimeinSeconds = 60


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
                        #Iterate over the First 500 Messages from newest to latest
                        for message in reversed(messages):
                            if counter == 500:
                                continue
                            #Check if the Category of the Message is like the defined Category in the Mappingtable
                            if message.Categories == row['Label']:
                                try:
                                    #Try to open the URL to check if it is reachable
                                    url_requested = urllib.request.urlopen(row['JIRAURL'])
                                    if 200 == url_requested.code:
                                        #Create JIRA Issue and clear Category if jira Issue was created
                                        if jira_py.createjiraIssue(row['JIRAURL'], row['JIRAUser'], row['JiraPW'], row['ProjectID'], message.Subject, message.Body, row['IssueType']):
                                            message.Categories = ""
                                            message.save()  
                                #Except if the URL could not be read
                                except urllib.error.URLError as e: print('URL ' + row['JIRAURL'] + ' could not be read')
                                #Except a ValueError and prints it
                                except ValueError as e: print(e)
                            counter += 1

    print("Iterate")
    time.sleep(IterateTimeinSeconds)
    



# import pymsgbox for displaying a messagebox
import pymsgbox
#import pandas for reading the xlsx File
import pandas as pd
#import the request to check if the URL from the mapping is available
import urllib.request,urllib.parse,urllib.error

#Create a Messagebox with Yes/No
result = pymsgbox.confirm('Shall we create JIRA Issues from Mail?', 'Outlook to JIRA', ["Yes", 'No'])

#Declare the filepath to the mappingtable
filepath = "Mappingtable.xlsx"


#End the Script if the Selection was NO or None
if result == 'No' or result is None:
    print("End")
    quit()




#Load file into variable
data = pd.read_excel(filepath)
for index, row in data.iterrows():
    try:
        url_requested = urllib.request.urlopen(row['JIRAURL'])
        if 200 == url_requested.code:
            print("True")
    except urllib.error.URLError as e: print('URL ' + row['JIRAURL'] + ' could not be read')

    
    






# pymailtojira

A Python Application which reads an Excel file and the outlook accounts and creates jira issue in a jira location.

The mapping is done in Excel on which the user adds in the jiraURL, the JIRAUsername and password, the projectID, the IssueType, the mailaccount in outlook and the label in outlook.

The Application search for the label in the account and creates one jira issue for each labelled mail. Afterwards it will be unlabelled.


The Application provide a bat file (Run_Script.bat). You need to enter the Path to your python interpreter and send a shortcut from the directory where you want to click on the bat.

Read HowTo for Additional Information on how to use the application.
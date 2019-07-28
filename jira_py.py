import jira
from jira import JIRA

#Authentication Method
def authenticate(username, password):
    basic_auth=(username, password)
    return basic_auth

#Create a JIRA Object which can be used later on
def createJIRAObject(jiraURL, username, password):
    jira = JIRA(server=jiraURL, basic_auth=authenticate(username, password))
    return jira

#Creates a JIRA Issues and returns it
def createjiraIssue(jiraURL, username, password, projectID, summary, description, issueTypeName):
    jira = createJIRAObject(jiraURL, username, password)
    issue_dict = {
        'project': {'id': projectID},
        'summary': summary,
        'description': description,
        'issuetype': {'name': issueTypeName},
        'assignee': {'name': username}
    }
    new_issue = jira.create_issue(fields=issue_dict)

    return new_issue


def add_attachment(jiraURL, username, password, issue, URL):
    jira = createJIRAObject(jiraURL, username, password)
    return jira.add_attachment(issue=issue, attachment=URL)
	

from jira import JIRA

#Authentication Method
def authenticate(username, password):
    basic_auth=(username, password)
    return basic_auth


#Creates a JIRA Issues and returns it
def createjiraIssue(jiraURL, username, password, projectID, summary, description, issueTypeName):
    jira = JIRA(server=jiraURL, basic_auth=authenticate(username, password))
    issue_dict = {
        'project': {'id': projectID},
        'summary': summary,
        'description': description,
        'issuetype': {'name': issueTypeName},
        'assignee': {'name': username}
    }
    new_issue = jira.create_issue(fields=issue_dict)

    return new_issue

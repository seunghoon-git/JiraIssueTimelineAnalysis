import requests
import json
from datetime import datetime
import numpy as np
import logging
import pandas as pd
from IPython.display import display

def getIssueListFromJql(jql):
    jira_api = 'https://<your jira domain>/rest/issueNav/1/issueTable'
    headers = {
        'Authorization': '<Jira Auth Key>',
        'X-Atlassian-Token': 'nocheck'
    }
    data = {
        'startIndex' : 0,
        'jql' : jql,
        'layoutKey' : 'list-view'
    }

    try:
        res = requests.post(jira_api, headers=headers, data = data)
        assert res.status_code == 200, res.text
        res = json.loads(res.text.encode('utf-8'))
        return res.get('issueTable').get('issueKeys')
    except Exception as e:
        return e

def getIssueStatusHistory(key):
    jira_api = 'https://<your jira domain>/rest/gira/1/'
    headers = {'Authorization': '<Jira Auth Key>'}
    data = {
        "query":"\n    query IssueHistoryQuery($issueKey: String!, $startAt: Long, $maxResults: Int, $orderBy: String) {\n        viewIssue(issueKey: $issueKey) {\n            history(orderBy: $orderBy, startAt: $startAt, maxResults: $maxResults) {\n                isLast\n                totalCount\n                startIndex\n                nodes {\n                    fieldId\n                    fieldType\n                    fieldSchema {\n                        type\n                        customFieldType\n                    }\n                    timestamp\n                    actor {\n                        avatarUrls {\n                            large\n                        }\n                        displayName\n                        accountId\n                    }\n                    fieldDisplayName\n                    from {\n                        ... on GenericFieldValue {\n                            displayValue\n                            value\n                        }\n                        ... on AssigneeFieldValue {\n                            displayValue\n                            value\n                            avatarUrl\n                        }\n                        ... on PriorityFieldValue {\n                            displayValue\n                            value\n                            iconUrl\n                        }\n                        ... on StatusFieldValue {\n                            displayValue\n                            value\n                            categoryId\n                        }\n                        ... on WorkLogFieldValue {\n                            displayValue\n                            value\n                            worklog {\n                                id\n                                timeSpent\n                            }\n                        }\n                        ... on RespondersFieldValue {\n                            displayValue\n                            value\n                            responders {\n                                ari\n                                name\n                                type\n                                avatarUrl\n                            }\n                        }\n                    }\n                    to {\n                        ... on GenericFieldValue {\n                            displayValue\n                            value\n                        }\n                        ... on AssigneeFieldValue {\n                            displayValue\n                            value\n                            avatarUrl\n                        }\n                        ... on PriorityFieldValue {\n                            displayValue\n                            value\n                            iconUrl\n                        }\n                        ... on StatusFieldValue {\n                            displayValue\n                            value\n                            categoryId\n                        }\n                        ... on WorkLogFieldValue {\n                            displayValue\n                            value\n                            worklog {\n                                id\n                                timeSpent\n                            }\n                        }\n                        ... on RespondersFieldValue {\n                            displayValue\n                            value\n                            responders {\n                                ari\n                                name\n                                type\n                                avatarUrl\n                            }\n                        }\n                    }\n                }\n            }\n        }\n    }\n"
        , "variables":{"issueKey":"","startAt":0,"maxResults":100,"orderBy":"created"}
    }
    data['variables']['issueKey'] = key

    data = json.dumps(data)
    response = requests.post(jira_api, headers=headers, data=data)
    result = {}

    try:
        assert response.status_code == 200, response.text
        historyData=json.loads(response.text.encode('utf-8'))
        statusHistory=historyData.get('data').get('viewIssue').get('history').get('nodes')
        isInitiated=0
        
        assert statusHistory != None, "There is no history"
        for item in statusHistory:
            if item.get('fieldId') == 'status' and item.get('to').get('displayValue') != None and item.get('timestamp') != None:
                if item['to']['displayValue'] == 'In Progress':
                    if isInitiated == 0:
                        result['date_in_progress_min'] = str(datetime.fromtimestamp(round(item['timestamp']/1000)))
                        isInitiated = 1
                    result['date_in_progress_max'] = str(datetime.fromtimestamp(round(item['timestamp']/1000)))
                elif item['to']['displayValue'] == 'In Review':
                    result['date_in_review'] = str(datetime.fromtimestamp(round(item['timestamp']/1000)))
                elif item['to']['displayValue'] == 'Resolved':
                    result['date_resolved'] = str(datetime.fromtimestamp(round(item['timestamp']/1000)))
                elif item['to']['displayValue'] == 'Closed':
                    result['date_closed'] = str(datetime.fromtimestamp(round(item['timestamp']/1000)))
        return result
    except Exception as e:
        print(e)
        return result        

def getIssueInfo(key):
    jira_api = 'https://<your jira domain>/rest/graphql/1/'
    headers = {'Authorization': '<Jira Auth Key>'}
    query = 'query {\n        issue(issueIdOrKey: \"'+key+'\", latestVersion: true, screen: \"view\") {\n            id\n            viewScreenId \n            fields {\n                key\n                title\n                editable\n                required\n                autoCompleteUrl\n                allowedValues\n                content\n                renderedContent\n                schema {\n                    custom\n                    system\n                    configuration {\n        key\n        value\n    }\n    \n                    items\n                    type\n                    renderer\n                }\n                configuration\n            }\n            expandAssigneeInSubtasks\n            expandAssigneeInIssuelinks\n            expandTimeTrackingInSubtasks\n            systemFields {\n                descriptionAdf {\n                    value\n                }\n                environmentAdf {\n        value\n    }\n            }\n            customFields {\n                textareaAdf {\n                    key\n                    value\n                }\n            }            \n            tabs {\n        id\n        name\n        items {\n            id\n            type\n        }\n    }\n            \n    isHybridAgilityProject\n    \n            \n    agile {\n        epic {\n          key\n        },\n    }\n        }\n        \n        project(projectIdOrKey: \"PI\") {\n            id\n            name\n            key\n            projectTypeKey\n            simplified\n            avatarUrls {\n                key\n                value\n            }\n            archived\n            deleted\n        }\n    }'

    try:
        res = requests.post(jira_api, headers=headers, json={"query" : query})
        assert res.status_code == 200, res.text

        res = json.loads(res.text.encode('utf-8'))
        result = {}
        
        assert res.get('data').get('issue') != None, "Could not find the issue with the key '{}'".format(key)
        
        for  i in res.get('data').get('issue').get('fields'):
            #key
            result['key'] = key
            # issuetype
            if i.get('key') == 'issuetype' and i.get('content').get('name'):
                result['issuetype'] = i['content']['name']
            # summary
            if i.get('key') == 'summary' and i.get('content'):
                result['summary'] = i['content']
            # components
            if i.get('key') == 'components' and i.get('content'):
                for c in i['content']:
                    componentName = c.get('name') if c.get('name') else ''
                    if 'SQ - ' in componentName:
                        result['components'] = componentName
                        break
            # status
            if i.get('key') == 'status' and i.get('content'):
                result['status'] = i.get('content').get('name') if i.get('content').get('name') else ''
            # labels
            if i.get('key') == 'labels' and i.get('content'):
                result['labels'] = i['content']
            # Project Name
            if i.get('key') == 'labels' and i.get('content'):
                for l in i['content']:
                    if l.isupper():
                        result['Project Name'] = l
                        break
            # subtasks
            if i.get('key') == 'subtasks' and i.get('content'):
                result['subtasks']=[st.get('key') for st in i['content']]
            # Date created
            if i.get('key') == 'created' and i.get('content'):
                result['Date created']=i['content'][:10]
            # Date resolved
            if i.get('key') == 'resolutiondate' and i.get('content'):
                result['Date resolved']=i['content'][:10]
            # Date released (from fixVersions)
            if i.get('key') == 'fixVersions' and i.get('content'):
                result['Date released'] = min(f.get('releaseDate') for f in i['content'])
        # Date in reiview & Date in progress & Date closed
        historyLog = getIssueStatusHistory(key) 
        if historyLog.get('date_in_progress_min'): result['Date in progress'] = historyLog.get('date_in_progress_min')[:10] 
        if historyLog.get('date_in_review'): result['Date in review'] = historyLog.get('date_in_review')[:10]
        if historyLog.get('date_closed'): result['Date closed'] = historyLog.get('date_closed')[:10]
            
        return result
    except Exception as e:
        print(e)
        return {} 
    

jql = input("JQL Query : ").strip()
fileName = 'JiraProductivityAnalysis_{}'.format(datetime.today().strftime("%Y%m%d%H%M%S"))
logger = logging.getLogger('jiralog')
logger.setLevel(logging.DEBUG)
logger.addHandler(logging.FileHandler(fileName+'.log'))

logger.info('Get the list of Jira keys using the following jql')
logger.info('-------------------------------------------------')
logger.info(jql)
logger.info('-------------------------------------------------')

print('1. Getting Jira Keys from the jql')
issueList = getIssueListFromJql(jql)
defaultColumns = ['key','issuetype','summary','components','labels','status','subtasks','Project Name','Date created','Date in progress','Date in progress adj','Date in review','Date resolved','Date closed','Date released','Dev Time','Cycle Time']   
result = []

if len(issueList)>0:    
    logger.info('{} keys found'.format(len(issueList)))
    
    for n, i in enumerate(issueList, start=1):
        print('2. Processing...{}/{}'.format(n, len(issueList)), end='\r')
        
        # Get basic information
        msg = '{} -> Extract Basic Info'.format(i)
        defaultData = dict.fromkeys(defaultColumns)
        data = getIssueInfo(i)
        data = {**defaultData, **data}

        # Check subtasks' date in progress
        if data.get('subtasks'):
            msg += " -> Check subtasks' date in progress"
            subtask_HistoryLog = []
            for st in data['subtasks']:
                summary = getIssueInfo(st).get('summary')
                if not summary: continue
                # Exclude subtask for writing test cases 
                if summary and not ('TC' in summary and '작성' in summary):
                    if getIssueStatusHistory(st).get('date_in_progress_max'): subtask_HistoryLog.append(getIssueStatusHistory(st).get('date_in_progress_max'))
            if len(subtask_HistoryLog) > 0: data['Date in progress adj'] = min(subtask_HistoryLog)[:10]            
        
        # Calculate dev time : days from in progress to resolved
        if (data.get('Date in progress') or data.get('Date in progress adj')) and (data.get('Date resolved') or data.get('Date closed') or data.get('Date released')): 
            msg += ' -> Calculate Dev Time'
            date_in_progress_final = min(list(map(lambda x : x if x else '9999-12-31', [data.get('Date in progress'), data.get('Date in progress adj')]))) 
            date_in_resolved_final = min(list(map(lambda x : x if x else '9999-12-31', [data.get('Date resolved'), data.get('Date closed'), data.get('Date released')])))
            data['Dev Time'] = np.busday_count(date_in_progress_final, date_in_resolved_final)

        # Calculate cycle time
        if (data.get('Date in progress') or data.get('Date in progress adj')) and data.get('Date released'): 
            msg += ' -> Calculate Cycle Time'
            date_in_progress_final = min(list(map(lambda x : x if x else '9999/12/31', [data.get('Date in progress'), data.get('Date in progress adj')]))) 
            data['Cycle Time'] = np.busday_count(date_in_progress_final, data['Date released'])
        
        msg += ' -> [Complete]'
        logger.info(msg)
        result.append(data)
        if n == len(issueList): print('2. Data processing completed')       
    
    # Export to xlsx
    print('3. Data export to excel')
    df = pd.DataFrame(result)
    df.fillna('', inplace = True)
    df.to_excel('{}.xlsx'.format(fileName))
else:
    logger.info('No keys found')

logger.handlers.clear()
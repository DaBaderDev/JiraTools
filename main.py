import requests
import json
import datetime

from openpyxl import Workbook
from openpyxl.chart import Reference, PieChart, BarChart

from config import *

workbook = Workbook()

sheet = workbook.active

project_completion = 0

# Create session
session = requests.Session()
session.auth = (JIRA_EMAIL, JIRA_API_TOKEN)
response = session.get(JIRA_API_ENDPOINT)
project_name = response.json()['values'][0]['location']['projectName']
board_id = response.json()['values'][0]['id']

# Get sprints
response = session.get(f'{JIRA_BOARD_ENDPOINT}/{board_id}/sprint')
sprints = response.json()['values']

sprint_data = [
    ['Sprint', 'Completion', 'Time needed']
]

for sprint in sprints:
    if sprint['state'] == 'closed':
        begin_date = datetime.datetime.fromisoformat(sprint['startDate'])
        complete_date = datetime.datetime.fromisoformat(sprint['completeDate'])
        
        sprint_time = complete_date - begin_date
        
        print(f"Sprint \"{sprint['name']}\" completion: 100.00% | Time needed: {sprint_time}")
        
        project_completion = project_completion + 100
        
        sprint_data.append([sprint['name'], 100.0, sprint_time])
    else:
        now = datetime.datetime.utcnow()
        
        begin_date = now
        
        if sprint['state'] != 'future':
            begin_date = datetime.datetime.fromisoformat(sprint['startDate']).replace(tzinfo=None)
        
        sprint_time = now - begin_date
        
        # Get issues in sprint
        payload = {
            "jql": f"Sprint={sprint['id']}"
        }
        headers = {
            "Content-Type": "application/json"
        }
        
        response = session.post(JIRA_SEARCH_ENDPOINT, data=json.dumps(payload), headers=headers)
        # print(response.json())
        issues = response.json()['issues']

        # Calculate sprint completion percentage
        total_issues = len(issues)
        
        completed_issues = len([issue for issue in issues if issue['fields']['status']['statusCategory']['key'].lower() == 'done'])
        completion_percentage = (completed_issues / total_issues) * 100 if total_issues != 0 else 0
        print(f"Sprint \"{sprint['name']}\" completion: {completion_percentage:.2f}% | Time needed: {sprint_time}")
        project_completion = project_completion + completion_percentage
        
        sprint_data.append([sprint['name'], completion_percentage, sprint_time])

project_completion = project_completion / len(sprints)

print(f"Project completion: {project_completion:.2f}%")

for row in sprint_data:
    sheet.append(row)

# Setup Time chart
sprint_time_chart = BarChart()

categories = Reference(sheet, min_col=1, min_row=1, max_row=len(sprint_data))
values = Reference(sheet, min_col=3, min_row=1, max_row=len(sprint_data))
sprint_time_chart.add_data(values, titles_from_data=True)
sprint_time_chart.set_categories(categories)

project_completion_chart = PieChart()

sheet.add_chart(sprint_time_chart, "E2")

workbook.save('chart.xlsx')
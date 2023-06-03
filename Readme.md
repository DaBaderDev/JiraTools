# Jira Tools

A script that gets a specified projects sprints and output some data based on them.

## Requirements:
- openpyxl is needed `pip install openpyxl`

## Getting started:
- Rename the file `config.py.example` to `config.py`
- Change `JIRA_SERVER_URL` to your Jira Server url
- Change `JIRA_EMAIL` to the email you use to login
- Change `JIRA_API_TOKEN` to your API token (Get API token [here](https://id.atlassian.com/manage-profile/security/api-tokens))
- To run the script: `python main.py`
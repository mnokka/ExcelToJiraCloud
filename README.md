# ExcelToJiraCloud
Tool to parse excel data and create issues to Jira Cloud (server/DC)
<br>
<br>
Just create an issue (Epic) , check that auth and creation works:
<br>
<br>
python3 CreateIssue.py -u "EMAIL_FOR_ACCOUNT_USED_FOR_TOKEN" -w JIRA_CLOUD_TOKEN -s JIRA_CLOUD_ADDRESS -y "Summary text" -d "Description text" -k PROJECT_KEY
<br>
<br>
Included example excel sheet is parsed using ReadExcel.py tool. Target Jira requires have defined custom fields in the project screens and their IDs added tool code (Jira has custom fields URL,POINTS,VALUE added to test project screens and their IDs has been added to IssueCreator.py new issue dict definition)
<br>
<br>
USAGE:
<br>

python3 ReadExcel.py  -f . -n testexcel.xlsx -w JIRA_CLOUD_TOKEN   -u "EMAIL_FOR_ACCOUNT_USED_FOR_TOKEN" -s JIRA_CLOUD_ADDRESS -p PROJECT_KEY

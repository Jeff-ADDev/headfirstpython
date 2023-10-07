import os, requests, sys
import openpyxl
from typing import List
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from colorama import init, Fore, Back, Style
from dotenv import load_dotenv
from datetime import datetime
from epic import Epic
from issue import Issue
from sprint import Sprint

start_time = datetime.now()
start_time_format = start_time.strftime("%m/%d/%Y, %H:%M:%S")

load_dotenv()
init() # Colorama   

jirakey = os.getenv("JIRA_API_KEY")
url_location = os.getenv("JIRA_REV_LOCATION")
url_search = os.getenv("JIRA_SEARCH")
url_board = os.getenv("JIRA_BOARD")
sheets_location = os.getenv("SHEETS_LOCATION")
project_label = os.getenv("PROJECT_LABEL")
jira_issue_link = os.getenv("JIRA_ISSUE_LINK")
project_figma_link = os.getenv("FIGMA_LINK")

main_serach = f"{url_location}/{url_search}"
header = {"Authorization": "Basic " + jirakey}
baord_issues = f"{url_location}/{url_board}"

epics: List[Epic] = []

# Retrieve all epics from main project label
# Get Sub labels to help break down the epics
all_epics = main_serach + "'issuetype'='Epic' AND ('Status'='FUTURE' OR 'Status'='NEXT' OR 'Status'='Now') AND 'labels' in ('" + project_label + "')"
response = requests.get(all_epics, headers=header)
if response.status_code == 200:
    data = response.json()
    for epicitem in data["issues"]:
        epic_add = Epic(epicitem["id"], epicitem["key"], epicitem["fields"]["summary"], epicitem["fields"]["created"])
        epic_add.set_team(epicitem["fields"]["project"]["name"])
        epic_add.set_estimate(epicitem["fields"]["customfield_10032"])
        for label in epicitem["fields"]["labels"]:
            if (label != project_label):
                epic_add.add_sub_label(label)
        epics.append(epic_add)
        
    print(Fore.GREEN + f"Success! - All Epics {len(epics)}" + Style.RESET_ALL)
else:
    print(Fore.RED + "Failed - All Epics" + Style.RESET_ALL)

# Retrieve all the issues attached to the epics
issues_with_points = 0
issues_points = 0
issues_with_no_points = 0
for epicitem in epics:
    epic_issues = main_serach + "'Epic Link'='" + epicitem.key + "' and STATUS != Cancelled"
    response = requests.get(epic_issues, headers=header)
    if response.status_code == 200:
        data = response.json()
        for issue in data["issues"]:
            assigned_points = issue["fields"]["customfield_10032"]
            issue_add = Issue(issue["id"], issue["key"], issue["fields"]["summary"], assigned_points)
            if assigned_points == None:
                issues_with_no_points += 1
            else:
                issues_with_points += 1
                try: 
                    issues_points += assigned_points
                except:
                    pass
            try:
                for item in issue["fields"]["customfield_10010"]:
                    add_sprint = Sprint(item["id"], item["name"], item["boardId"], item["state"])
                    if "completeDate" in item:
                        add_sprint.set_completeDate(item["completeDate"])
                    issue_add.add_sprint(add_sprint)
            except:
                pass
            epicitem.add_issue(issue_add)
            epicitem.set_issues_with_points(issues_with_points)
            epicitem.set_issues_points(issues_points)
            epicitem.set_issues_with_no_points(issues_with_no_points)

# Console Output Information
for epicitem in epics:
    epicitem.print_Epic()
    for issueitem in epicitem.issues:
        issueitem.print_Issue()
        for sprintitem in issueitem.sprint:
            sprintitem.print_Sprint()

# Create new workbook
# Epic Tab
# Key(link) | Summary | Team | Estimate | Issues with Points | Issues with No Points | Issues Points | Sub Labels
workbook = openpyxl.Workbook()
worksheet_epics = workbook.active
worksheet_epics.title = "Epics"
worksheet_issues = workbook.create_sheet("All Issues")

worksheet_epics.column_dimensions["A"].width = 16
worksheet_epics.column_dimensions["B"].width = 50
worksheet_epics.column_dimensions["C"].width = 30
worksheet_epics.column_dimensions["D"].width = 18
worksheet_epics.column_dimensions["E"].width = 18
worksheet_epics.column_dimensions["F"].width = 18
worksheet_epics.column_dimensions["G"].width = 18
worksheet_epics.column_dimensions["H"].width = 40

table = Table(displayName="TableEpics", ref="A1:H" + str(len(epics) + 1))
style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                    showLastColumn=False, showRowStripes=True, showColumnStripes=False)
table.tableStyleInfo = style

worksheet_epics.append(["Epic", "Summary", "Team", "Estimate", "Issues w Points", "Issues w No Points ", "Issues Total Points", "Sub Labels"])
for epicitem in epics:
    sub_labels = epicitem.get_sublevles()
    worksheet_epics.append([epicitem.key, epicitem.summary, epicitem.team, epicitem.estimate, epicitem.issues_with_points, epicitem.issues_with_no_points, epicitem.issues_points, sub_labels])    

worksheet_epics.add_table(table)

for row in worksheet_epics[2:worksheet_epics.max_row]:  # Exclude The Header
    cell = row[0] # zeor based index
    value_use = cell.value
    cell.hyperlink = f"{jira_issue_link}{value_use}"
    cell.value = value_use
    cell.style = "Hyperlink"

for row in worksheet_epics[1:worksheet_epics.max_row]:  # Include The Header
    cell = row[0] # zeor based index
    cell.alignment = Alignment(horizontal="center", vertical="center")

for row in worksheet_epics[2:worksheet_epics.max_row]:  # skip the header
    cell = row[1] # zeor based index
    cell.alignment = Alignment(wrap_text=True)
    cell.number_format = "text"

for row in worksheet_epics[1:worksheet_epics.max_row]:  # Include The Header
    cell = row[2] # zeor based index
    cell.alignment = Alignment(horizontal="center", vertical="center")

for row in worksheet_epics[1:worksheet_epics.max_row]:  # Include The Header
    cell = row[3] # zeor based index
    cell.alignment = Alignment(horizontal="center", vertical="center")

for row in worksheet_epics[1:worksheet_epics.max_row]:  # Include The Header
    cell = row[4] # zeor based index
    cell.alignment = Alignment(horizontal="center", vertical="center")

for row in worksheet_epics[1:worksheet_epics.max_row]:  # Include The Header
    cell = row[5] # zeor based index
    cell.alignment = Alignment(horizontal="center", vertical="center")

for row in worksheet_epics[1:worksheet_epics.max_row]:  # Include The Header
    cell = row[6] # zeor based index
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Add Figma Plan Link to the bottom of the Epics Sheet
worksheet_epics["A" + str(len(epics) + 3)].hyperlink = project_figma_link
worksheet_epics["A" + str(len(epics) + 3)].value = "Figma Plan"
worksheet_epics["A" + str(len(epics) + 3)].style = "Hyperlink"

# Start Building Issues Tab
worksheet_issues.column_dimensions["A"].width = 16
worksheet_issues.column_dimensions["B"].width = 50
worksheet_issues.column_dimensions["C"].width = 30
worksheet_issues.column_dimensions["D"].width = 18
worksheet_issues.column_dimensions["E"].width = 18
worksheet_issues.column_dimensions["F"].width = 18
worksheet_issues.column_dimensions["G"].width = 18
worksheet_issues.column_dimensions["H"].width = 40

# Get all Issue Count
all_issues = 0
for epicitem in epics:
    all_issues += (len(epicitem.issues))
table_issues = Table(displayName="TableIssues", ref="A1:E" + str(all_issues + 1))
style_issues = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                    showLastColumn=False, showRowStripes=True, showColumnStripes=False)
table_issues.tableStyleInfo = style_issues

worksheet_issues.append(["Issue", "Summary", "Team", "Estimate", "Size"])
for epicitem in epics:
    for issueitem in epicitem.issues:
        worksheet_issues.append([issueitem.key, issueitem.summary, epicitem.team, issueitem.size])

worksheet_issues.add_table(table_issues)

# Handle Directory
if os.path.exists(sheets_location):
    saveexcelfile = sheets_location + "Project " + project_label + " Details.xlsx"
else:
    os.makedirs(sheets_location)

# Check For File Existence - Delete if exists
if os.path.exists(saveexcelfile):
    os.remove(saveexcelfile)

workbook.save(saveexcelfile)
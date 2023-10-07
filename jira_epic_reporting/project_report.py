import os, requests, sys, argparse
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
date_file_info = start_time.strftime("%Y_%m_%d")
create_date = start_time.strftime("%m/%d/%Y")

load_dotenv()
init() # Colorama   

jirakey = os.getenv("JIRA_API_KEY")
url_location = os.getenv("JIRA_REV_LOCATION")
url_search = os.getenv("JIRA_SEARCH")
url_board = os.getenv("JIRA_BOARD")
path_location = os.getenv("PATH_LOCATION")
project_label = os.getenv("PROJECT_LABEL")
jira_issue_link = os.getenv("JIRA_ISSUE_LINK")
project_figma_link = os.getenv("FIGMA_LINK")

main_serach = f"{url_location}/{url_search}"
header = {"Authorization": "Basic " + jirakey}
baord_issues = f"{url_location}/{url_board}"
epics: List[Epic] = []  

# Retrieve all epics from main project label
# Get Sub labels to help break down the epics
def get_epics(label, con_out):
    all_epics = main_serach + "'issuetype'='Epic' AND ('Status'='FUTURE' OR 'Status'='NEXT' OR 'Status'='Now') AND 'labels' in ('" + label + "')"
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
        if con_out:
            print(Fore.GREEN + f"Success! - All Epics {len(epics)}" + Style.RESET_ALL)
    else:
        if con_out:
            print(Fore.RED + "Failed - All Epics" + Style.RESET_ALL)

# Retrieve all the issues attached to the epics
# JSON Issue
# issues
#      id - 12345                                           id
#      key - ARR-2392                                       key
#      fields
#            summary - "Create a new project report"        summary
#            customfield_10032 - 5 Points                   size  
#            customfield_10010 - Sprint[]                   sprint[]
#            status
#                  name - "In Progress"                     status
#            priority
#                  name - "Medium"                          priority
#            issuetype
#                  name - "Story"                           issuetype
#            project
#                  name - "Agile RevSite Raider$"           project_name
#                  key - "ARR"                              project_key
#            assignee
#                  displayName - "John Doe"                 assignee_displayName
#            created - "2021-03-01T15:00:00.000-0400        created
#            updated - "2021-03-01T15:00:00.000-0400        updated
#            description - "This is a description"          description

def get_issues():
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
                issue_add.set_status(issue["fields"]["status"]["name"])
                issue_add.set_priority(issue["fields"]["priority"]["name"])
                issue_add.set_issuetype(issue["fields"]["issuetype"]["name"])
                issue_add.set_project_name(issue["fields"]["project"]["name"])
                issue_add.set_project_key(issue["fields"]["project"]["key"])
                if issue["fields"]["assignee"] is not None:
                    if "displayName" in issue["fields"]["assignee"]:
                        issue_add.set_assignee_displayName(issue["fields"]["assignee"]["displayName"])
                issue_add.set_created(issue["fields"]["created"])
                issue_add.set_updated(issue["fields"]["updated"])
                issue_add.set_description(issue["fields"]["description"])
                
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
def output_console():
    for epicitem in epics:
        epicitem.print_Epic()
        for issueitem in epicitem.issues:
            issueitem.print_Issue()
            for sprintitem in issueitem.sprint:
                sprintitem.print_Sprint()

# Create new workbook
# --- Summary Tab ---
# A1           B1       C1            D1            E1            F1            G1            H1
# Label                 Created
#                      
# [Epics]              [Issues Overall]
# Count                 Count
# Total Estimate        Total Estimate
# With Estimates        With Estimates
# Percent               Percent
# Average Estimate      Average Estimate
# Max Estimate          Max Estimate              
# Min Estimate          Min Estimate
# Issues in Epic
# [Count]
# Total Estimate
# With Estimates
# Percent
# Average Estimate
# Max Estimate
# Min Estimate
#
#
#
# --- Epic Tab ---
# Key(link) | Summary | Team | Estimate | Issues with Points | Issues with No Points | Issues Points | Sub Labels
#
# --- Issue Tab ---
# Define
#
def create_excel(label):
    workbook = openpyxl.Workbook()
    worksheet_summary = workbook.active
    worksheet_summary.title = "Summary"
    worksheet_epics = workbook.create_sheet("Epics")
    worksheet_issues = workbook.create_sheet("Issues")

    # Create the Summary Tab
    Epic.excel_worksheet_summary(worksheet_summary, epics, label, create_date)

    # Create the Epic Tab
    Epic.excel_worksheet_create(worksheet_epics, epics, jira_issue_link, project_figma_link,label)

    # Create the Issue Tab
    Issue.excel_worksheet_create(worksheet_issues, epics, jira_issue_link)

    return workbook

def save_file(path, filename, wb):
    # Handle Directory
    if os.path.exists(path):
        saveexcelfile = path + date_file_info + " Project " + filename + " Details.xlsx"
    else:
        os.makedirs(path)

    # Check For File Existence - Delete if exists
    if os.path.exists(saveexcelfile):
        os.remove(saveexcelfile)

    # Save Workbook
    wb.save(saveexcelfile)

def main(args):
    if args.label:
        project_label = args.label

    con_out = False
    if args.console:
        con_out = True

    get_epics(project_label, con_out)
    get_issues()
    wb = create_excel(project_label)
    save_file(path_location,project_label,wb)

    if con_out:
        output_console()  

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Create Excel Sheet for Project Reporting")
    parser.add_argument("-l", "--label", help="Label for the project")
    parser.add_argument("-c", "--console", help="Enable Console Output", action="store_true")
    args = parser.parse_args()
    main(args)

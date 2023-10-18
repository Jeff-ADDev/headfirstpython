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
from objects.issue import Issue
from objects.sprint import Sprint
import utils.excel_util as excel_util
import utils.claude_util as claude_util
import console_util
from objects.epic import Epic

def test_zero_value(value, cell):
    if value == 0:
        cell.value = " - "
    else:
        cell.value = value

# Retrieve all epics from main project label
# Get Sub labels to help break down the epics
def get_epics(project_label, con_out, main_search, header):
    epics = []
    console_util.terminal_update("Retrieving Epics", " - ", False)
    all_epics = main_search + "'issuetype'='Epic' AND 'labels' in ('" + project_label + "')"
    response = requests.get(all_epics, headers=header)
    if response.status_code == 200:
        data = response.json()
        for epicitem in data["issues"]:
            epic_add = Epic(epicitem["id"], epicitem["key"], epicitem["fields"]["summary"], epicitem["fields"]["created"])
            epic_add.set_team(epicitem["fields"]["project"]["name"])
            epic_add.set_estimate(epicitem["fields"]["customfield_10032"])
            epic_add.set_description(epicitem["fields"]["description"])
            for label in epicitem["fields"]["labels"]:
                if (label != project_label):
                    epic_add.add_sub_label(label)
            epics.append(epic_add)
        if con_out:
            print(Fore.GREEN + f"Success! - All Epics {len(epics)}" + Style.RESET_ALL)
        return epics
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
def get_issues(epics, main_search, header):
    issues_with_points = 0
    issues_points = 0
    issues_with_no_points = 0
    count_epics = (len(epics)+1)
    count_epics_current = 1
    for epicitem in epics:
        console_util.terminal_update("Retrieving Issues on Epics", f"{count_epics_current}/{count_epics}", False)
        count_epics_current += 1
        epic_issues = main_search + "'Epic Link'='" + epicitem.key + "' and STATUS != Cancelled"
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

def get_comments(epics, con_out,url_location, url_issue, header):
    console_util.terminal_update("Retrieving Comments", " - ", False)
    for epicitem in epics:
        #https://revlocaldev.atlassian.net/rest/api/2/issue/ARR-2392/comment
        all_comments = f"{url_location}/{url_issue}{epicitem.key}/comment"
        response = requests.get(all_comments, headers=header)
        if response.status_code == 200:
            data = response.json()
            for commentitem in data["comments"]:
                epicitem.add_comment(commentitem["body"])

# Console Output Information
def output_console(epics):
    for epicitem in epics:
        epicitem.print_epic()
        for issueitem in epicitem.issues:
            issueitem.print_issue()
            for sprintitem in issueitem.sprint:
                sprintitem.print_sprint()
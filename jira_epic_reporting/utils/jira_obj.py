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
from objects.board import Board
from objects.user import User
import console_util
from objects.epic import Epic

class Jira:
    def __init__(self, project_label, jirakey, url_location, url_search, url_board, url_issue, url_users, header, con_out):
        self.project_label = project_label
        self.jirakey = jirakey
        self.url_location = url_location
        self.url_search = url_search
        self.url_board = url_board
        self.url_issue = url_issue
        self.url_users = url_users
        self.header = header
        self.con_out = con_out

    def test_zero_value(value, cell):
        if value == 0:
            cell.value = " - "
        else:
            cell.value = value

    def get_epics(self):
        """
        Retrieve all epics from main project label
        """
        main_search = f"{self.url_location}/{self.url_search}"
        epics = []
        console_util.terminal_update("Retrieving Epics", " - ", False)
        all_epics = main_search + "'issuetype'='Epic' AND 'labels' in ('" + self.project_label + "')"
        response = requests.get(all_epics, headers=self.header)
        if response.status_code == 200:
            data = response.json()
            for epicitem in data["issues"]:
                epic_add = Epic(epicitem["id"], epicitem["key"], epicitem["fields"]["summary"], epicitem["fields"]["created"])
                epic_add.set_team(epicitem["fields"]["project"]["name"])
                epic_add.set_estimate(epicitem["fields"]["customfield_10032"])
                epic_add.set_description(epicitem["fields"]["description"])
                for label in epicitem["fields"]["labels"]:
                    if (label != self.project_label):
                        epic_add.add_sub_label(label)
                epics.append(epic_add)
            if self.con_out:
                print(Fore.GREEN + f"Success! - All Epics {len(epics)}" + Style.RESET_ALL)
            return epics
        else:
            if self.con_out:
                print(Fore.RED + "Failed - All Epics" + Style.RESET_ALL)

    def get_issues(self, epics):
        """
        Get all issues that are part of the epics
        """
        main_search = f"{self.url_location}/{self.url_search}"
        issue_comments = f"{self.url_location}/{self.url_issue}"
        issues_with_points = 0
        issues_points = 0
        issues_with_no_points = 0
        count_epics = (len(epics)+1)
        count_epics_current = 1
        for epicitem in epics:
            console_util.terminal_update("Retrieving Issues on Epics", f"{count_epics_current}/{count_epics}", False)
            count_epics_current += 1
            epic_issues = main_search + "'Epic Link'='" + epicitem.key + "' and STATUS != Cancelled"
            response = requests.get(epic_issues, headers=self.header)
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
                        
                    self.get_changelogs(issue_add)

                    epicitem.add_issue(issue_add)
                    epicitem.set_issues_with_points(issues_with_points)
                    epicitem.set_issues_points(issues_points)
                    epicitem.set_issues_with_no_points(issues_with_no_points)

    def get_comments(self, epics):
        """
        Get Epic Item Comments
        """
        console_util.terminal_update("Retrieving Comments", " - ", False)
        for epicitem in epics:
            #https://revlocaldev.atlassian.net/rest/api/2/issue/ARR-2392/comment
            all_comments = f"{self.url_location}/{self.url_issue}{epicitem.key}/comment"
            response = requests.get(all_comments, headers=self.header)
            if response.status_code == 200:
                data = response.json()
                for commentitem in data["comments"]:
                    epicitem.add_comment(commentitem["body"])

    def output_console(self, epics):
        """
        Output all information to console
        """
        for epicitem in epics:
            epicitem.print_epic()
            for issueitem in epicitem.issues:
                issueitem.print_issue()
                for sprintitem in issueitem.sprint:
                    sprintitem.print_sprint()

    def get_boards(self, board_info):
        """
        Get all boards from Jira
        """
        console_util.terminal_update("Retrieving Boards", " - ", False)
        temp_boards = []
        board_info = f"{self.url_location}/{self.url_board}"
        # https://revlocaldev.atlassian.net/rest/agile/1.0/board
        response = requests.get(board_info, headers=self.header)
        if response.status_code == 200:
            data = response.json()
            for boarditem in data["values"]:
                temp_boards.append(Board(boarditem["id"], boarditem["name"], boarditem["type"]))
                #temp_boards[boarditem["id"]] = f"{boarditem['name']}|{boarditem['type']}"
                # print(f"ID - {boarditem['id']} - Name - {boarditem['name']} - Type - {boarditem['type']}")
                # location
                #           projectId
                #           name
                #           projectKey
                #           projectTypeKey
                #           displayName
                #           projectNmae
            if self.con_out: 
                print(Fore.GREEN + f"Success! - Board Info" + Style.RESET_ALL)
            return temp_boards
        else:
            if self.con_out:
                print(Fore.RED + "Failed - All Boards Info" + Style.RESET_ALL)

    def get_sprints(self, temp_boards):
        """
        Get all sprints from Jira
        """
        count = 0
        all_sprints = []
        for boarditem in temp_boards:
            start_location = 0
            count_boarditem_sprints = 0
            has_more_sprints = True
            while (has_more_sprints):
                console_util.terminal_busy("Retrieving Sprints", count)
                count += 1
                if count > 3:
                    count = 0
                sprint_info = f"{self.url_location}/rest/agile/1.0/board/{boarditem.id}/sprint?maxResults=50&startAt={start_location}"
                response = requests.get(sprint_info, headers=self.header)
                if response.status_code == 200:
                    data = response.json()
                    for sprintitem in data["values"]:
                        new_sprint = Sprint(sprintitem["id"], sprintitem["name"], boarditem.id, sprintitem["state"])
                        # TODO: Check that these dates are getting set and working oaky, move date handling to console_util
                        try: 
                            new_sprint.set_start_date(sprintitem["startDate"])
                        except:
                            pass
                        try:
                            new_sprint.set_end_date(sprintitem["endDate"])
                        except:
                            pass
                        try: 
                            new_sprint.set_complete_date(sprintitem["completeDate"])
                        except:
                            pass
                        all_sprints.append(new_sprint)
                        #temp_sprints[sprintitem["id"]] = f"{boarditem}|{count_boarditem_sprints}|{sprintitem['name']}|{sprintitem['state']}|{sprintitem['startDate']}|{sprintitem['endDate']}"
                        count_boarditem_sprints += 1

                    if len(data["values"]) == 50:
                        has_more_sprints = True
                        start_location += 50
                    else:
                        has_more_sprints = False
                else:
                    has_more_sprints = False
        if self.con_out:
            print(Fore.GREEN + f"Success! - Sprint Info " + Style.RESET_ALL)
        else:
            if self.con_out:
                print(Fore.RED + "Failed - All Sprint Info" + Style.RESET_ALL)
        
        return all_sprints

    def get_users(self):
        """
        Get All Users from Jira
        """
        url = f"{self.url_location}/{self.url_users}"
        console_util.terminal_update("Retrieving Users", " - ", False)
        all_users = []
        response = requests.get(url, headers=self.header)
        if response.status_code == 200:
            data = response.json()
            for useritem in data:
                email = ""
                try:
                    email = useritem["emailAddress"]
                except:
                    pass
                all_users.append(User(useritem["accountId"], email, useritem["displayName"], useritem["active"]))
            if self.con_out:
                print(Fore.GREEN + f"Success! - Users Info" + Style.RESET_ALL)
            return all_users
        else:
            if self.con_out:
                print(Fore.RED + "Failed - Users Info" + Style.RESET_ALL)

    def get_changelogs(self, issue):
        """
        Get All Change Logs from Jira
        """
        url = f"{self.url_location}/{self.url_issue}"
        all_logs = f"{url}{issue.key}/changelog"
        response = requests.get(all_logs, headers=self.header)
        if response.status_code == 200:
            data = response.json()
            for change in data["values"]:
                print(f"{change['id']} - {change['created']}")

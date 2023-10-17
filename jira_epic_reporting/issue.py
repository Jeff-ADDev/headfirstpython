from typing import List
from colorama import init, Fore, Back, Style
from sprint import Sprint
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

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

class Issue:
    def __init__(self, id, key, summary, size):
        self.id = id
        self.key = key
        self.summary = summary
        self.size = size
        self.sprint: List[Sprint] = []
        self.status = ""
        self.priority = ""
        self.issuetype = ""
        self.project_name = ""
        self.project_key = ""
        self.assignee_displayName = ""
        self.created = ""
        self.updated = ""
        self.description = ""
    
    def set_status(self, status):
        self.status = status
    
    def set_priority(self, priority):
        self.priority = priority
    
    def set_issuetype(self, issuetype):
        self.issuetype = issuetype
    
    def set_project_name(self, project_name):
        self.project_name = project_name
    
    def set_project_key(self, project_key):
        self.project_key = project_key

    def set_assignee_displayName(self, assignee_displayName):
        self.assignee_displayName = assignee_displayName

    def set_created(self, created):
        self.created = datetime.strptime(created, "%Y-%m-%dT%H:%M:%S.%f%z")
    
    def set_updated(self, updated):
        self.updated = datetime.strptime(updated, "%Y-%m-%dT%H:%M:%S.%f%z")

    def set_description(self, description):
        self.description = description

    def add_sprint(self, sprint):
        self.sprint.append(sprint)

    def set_sprint_name(self, sprint_name):
        self.sprint_name = sprint_name

    def set_sprint_state(self, sprint_state):
        self.sprint_state = sprint_state

    def set_boardID(self, boardID):
        self.boardID = boardID

    def set_completeDate(self, completeDate):
        self.completeDate = completeDate

    def print_issue(self):
        print(Fore.CYAN + Style.BRIGHT + 
              "  Issue-" + self.key + ": " + Fore.CYAN + Style.NORMAL + self.summary 
              + Fore.BLUE + Style.BRIGHT + " (" + str(self.size) + ")" + Style.RESET_ALL)
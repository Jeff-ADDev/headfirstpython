from typing import List
from colorama import init, Fore, Back, Style
from objects.sprint import Sprint
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from objects.changelog import Changelog

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
        self.times_to_dev = 0
        self.first_time_to_dev = ""
        self.hours_in_dev = 0
        self.times_to_qa = 0
        self.first_time_to_qa = ""
        self.hours_in_qa = 0
        self.times_to_uat = 0
        self.first_time_to_uat = ""
        self.hours_in_uat = 0
        self.changelogs: List[Changelog] = []
        self.date_done = ""
        self.date_ready_dev = ""
        self.total_hours = 0
        self.total_days = 0
        self.last_pointchange_date = ""

    
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

    def set_times_to_dev(self, times_to_dev):
        self.times_to_dev = times_to_dev
    
    def set_first_time_to_dev(self, first_time_to_dev):
        self.first_time_to_dev = first_time_to_dev
    
    def set_hours_in_dev(self, hours_in_dev):
        self.hours_in_dev = hours_in_dev
    
    def set_times_to_qa(self, times_to_qa):
        self.times_to_qa = times_to_qa

    def set_first_time_to_qa(self, first_time_to_qa):
        self.first_time_to_qa = first_time_to_qa

    def set_hours_in_qa(self, hours_in_qa):
        self.hours_in_qa = hours_in_qa

    def set_times_to_uat(self, times_to_uat):
        self.times_to_uat = times_to_uat
    
    def set_first_time_to_uat(self, first_time_to_uat):
        self.first_time_to_uat = first_time_to_uat

    def set_hours_in_uat(self, hours_in_uat):
        self.hours_in_uat = hours_in_uat

    def set_changelogs(self, changelogs):
        self.changelogs = changelogs

    def set_date_done(self, date_done):
        self.date_done = date_done

    def set_date_ready_dev(self, date_ready_dev):
        self.date_ready_dev = date_ready_dev

    def set_total_hours(self, total_hours):
        self.total_hours = total_hours

    def set_total_days(self, total_days):
        self.total_days = total_days

    def set_last_pointchange_date(self, last_pointchange_date):
        self.last_pointchange_date = last_pointchange_date

    def print_issue(self):
        print(Fore.CYAN + Style.BRIGHT + 
              "  Issue-" + self.key + ": " + Fore.CYAN + Style.NORMAL + self.summary 
              + Fore.BLUE + Style.BRIGHT + " (" + str(self.size) + ")" + Style.RESET_ALL)
        print(Fore.WHITE + Style.BRIGHT + 
              "    -------------------   Stats   ------------------ " + Style.RESET_ALL)
        print(Fore.WHITE + Style.NORMAL + 
              "    Points Added: " + str(self.last_pointchange_date) + " Size: " + str(self.size) +
              Style.RESET_ALL)        
        print(Fore.WHITE + Style.NORMAL + 
              "    Ready For Dev: " + str(self.date_ready_dev) +
              Style.RESET_ALL)        
        print(Fore.WHITE + Style.NORMAL + 
              "    Times to Dev: " + str(self.times_to_dev) + " First to Dev: " + self.first_time_to_dev + 
              " Hours in Dev: " + str(self.hours_in_dev) +
              Style.RESET_ALL)
        print(Fore.WHITE + Style.NORMAL + 
              "    Times to QA: " + str(self.times_to_qa) + " First to QA: " + self.first_time_to_qa + 
              " Hours in QA: " + str(self.hours_in_qa) +
              Style.RESET_ALL)
        print(Fore.WHITE + Style.NORMAL + 
              "    Times to UAT: " + str(self.times_to_uat) + " First to UAT: " + self.first_time_to_uat + 
              " Hours in UAT: " + str(self.hours_in_uat) +
              Style.RESET_ALL)
        print(Fore.WHITE + Style.NORMAL + 
              "    Date To Done: " + str(self.date_done) + " Total Days: " + str(self.total_days) + 
              Style.RESET_ALL)        
        print(Fore.WHITE + Style.BRIGHT + 
              "    ------------------------------------------------ " + Style.RESET_ALL)
        
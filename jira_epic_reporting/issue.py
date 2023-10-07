from typing import List
from colorama import init, Fore, Back, Style
from sprint import Sprint
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

class Issue:
    def __init__(self, id, key, summary, size):
        self.id = id
        self.key = key
        self.summary = summary
        self.size = size
        self.sprint: List[Sprint] = []

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

    def print_Issue(self):
        print(Fore.CYAN + Style.BRIGHT + 
              "  Issue-" + self.key + ": " + Fore.CYAN + Style.NORMAL + self.summary 
              + Fore.BLUE + Style.BRIGHT + " (" + str(self.size) + ")" + Style.RESET_ALL)
        
    def excel_worksheet_create(ws, epics, jira_issue_link):
        # Start Building Issues Tab
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 50
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 18
        ws.column_dimensions["G"].width = 18
        ws.column_dimensions["H"].width = 40

        all_issues = 0
        for epicitem in epics:
            all_issues += (len(epicitem.issues))
        table_issues = Table(displayName="TableIssues", ref="A1:E" + str(all_issues + 1))
        style_issues = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                            showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        table_issues.tableStyleInfo = style_issues

        ws.append(["Issue", "Summary", "Team", "Estimate", "Size"])
        for epicitem in epics:
            for issueitem in epicitem.issues:
                ws.append([issueitem.key, issueitem.summary, epicitem.team, issueitem.size])

        ws.add_table(table_issues)

        
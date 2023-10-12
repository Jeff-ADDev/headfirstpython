from typing import List
from issue import Issue
from colorama import init, Fore, Back, Style
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import numbers
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

class Jirainfo:
    def excel_boards(ws, boards):
        ws.append(["Board ID", "Board Name", "Board Type"])
        for board in boards:
            name, type = boards[board].split("|")
            ws.append([board,name,type])
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 35
        ws.column_dimensions["C"].width = 25

    def excel_sprints(ws, sprints):
        ws.append(["Sprint ID", "Board ID", "Number", "Name", "State"])
        for sprint in sprints:
            board_id, sprint_number, sprint_name, sprint_state = sprints[sprint].split("|")
            ws.append([sprint, board_id, sprint_number, sprint_name, sprint_state])
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 20
        ws.column_dimensions["D"].width = 45
        ws.column_dimensions["E"].width = 25
        
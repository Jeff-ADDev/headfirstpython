from typing import List
from colorama import init, Fore, Back, Style
from objects.sprint import Sprint
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from objects.changelog import Changelog
from objects.status import Status

class IssueType:
    def __init__(self, id, name):
        self.id = id
        self.name = name
        self.statuses: List[Status] = []

    def add_status(self, status):
        self.statuses.append(status)
    
    def add_statuses(self, statuses):
        self.statuses = statuses
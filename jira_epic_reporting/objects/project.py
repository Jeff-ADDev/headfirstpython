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

class Project:
    def __init__(self, id, key, name):
        self.id = id
        self.key = key
        self.name = name
        self.issues: List[Status] = []

    def add_issue(self, issue):
        self.issues.append(issue)
    
    def get_issues(self, issues):
        self.issues = issues
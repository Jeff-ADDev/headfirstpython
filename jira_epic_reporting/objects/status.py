from typing import List
from colorama import init, Fore, Back, Style
from objects.sprint import Sprint
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from objects.changelog import Changelog

# https://revlocaldev.atlassian.net/rest/api/2/project/project_key/statuses

class Status:
    def __init__(self, id, name, description):
        self.id = id
        self.name = name
        self.description = description

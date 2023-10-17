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
from issue import Issue
from sprint import Sprint
import excel_util
import claude_util
import console_util
from epic import Epic

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
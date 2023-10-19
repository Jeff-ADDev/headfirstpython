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
from objects.epic import Epic
from objects.issue import Issue
from objects.sprint import Sprint
import utils.excel_util as excel_util
import utils.claude_util as claude_util
import utils.jira_utils as jira_utils
import console_util

load_dotenv()
init() # Colorama   

jirakey = os.getenv("JIRA_API_KEY")
claudekey = os.getenv("CLAUDE_API_KEY")
url_location = os.getenv("JIRA_REV_LOCATION")
url_search = os.getenv("JIRA_SEARCH")
url_board = os.getenv("JIRA_BOARD")
url_issue = os.getenv("JIRA_ISSUE")
url_users = os.getenv("JIRA_USERS")
path_location = os.getenv("PATH_LOCATION")
project_label = os.getenv("PROJECT_LABEL")
jira_issue_link = os.getenv("JIRA_ISSUE_LINK")

def jira_project_label_reporting(args):
    epics: List[Epic] = []  
    main_search = f"{url_location}/{url_search}"
    header = {"Authorization": "Basic " + jirakey}
    baord_issues = f"{url_location}/{url_board}"
    start_time = datetime.now()
    start_time_format = start_time.strftime("%m/%d/%Y, %H:%M:%S")
    date_file_info = start_time.strftime("%Y_%m_%d")
    create_date = start_time.strftime("%m/%d/%Y")

    if args.label:
        project_label = args.label

    con_out = False
    if args.console:
        con_out = True
    
    ai_out = False
    if args.ai:
        ai_out = True

    other_links = {}
    if args.file:
        other_links = console_util.get_links(args.file)

    epics = jira_utils.get_epics(project_label, con_out, main_search, header)
    jira_utils.get_issues(epics, main_search, header)
    
    if ai_out:
        jira_utils.get_comments(epics, con_out, url_location, url_issue, header)

    wb = excel_util.create_label_excel_report(epics, project_label, other_links, 
        ai_out, create_date, jira_issue_link, claudekey)
    save_excel_file = date_file_info + " Project " + project_label + " Details.xlsx"
    console_util.save_excel_file(path_location, save_excel_file, wb)
    
    if con_out:
        jira_utils.output_console(epics)  

def jira_boards_sprint_reporting(args):
    con_out = False
    if args.console:
        con_out = True
    start_time = datetime.now()
    start_time_format = start_time.strftime("%m/%d/%Y, %H:%M:%S")
    date_file_info = start_time.strftime("%Y_%m_%d")
    create_date = start_time.strftime("%m/%d/%Y")
    
    board_issues = f"{url_location}/{url_board}"
    users_issues = f"{url_location}/{url_users}"
    header = {"Authorization": "Basic " + jirakey}

    all_boards = jira_utils.get_boards(con_out, board_issues, header)
    all_sprints = jira_utils.get_sprints(con_out, all_boards, url_location, header)
    all_users = jira_utils.get_users(con_out, users_issues, header)

    wb = excel_util.create_jira_info_report(all_boards, all_sprints, all_users)
    save_excel_file = date_file_info + " Jira Info.xlsx"
    console_util.save_excel_file(path_location, save_excel_file, wb)

def jira_people_reporting(args):
    print("Getting People Information")

def main(args):
    if args.label:
        jira_project_label_reporting(args)
    elif args.info:
        jira_boards_sprint_reporting(args)
    elif args.people:
        jira_people_reporting(args)
    else:  
        print("Please provide a proper argument for the program")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Create Excel Sheet for Project Reporting")
    parser.add_argument("-l", "--label", help="Label for the project")
    parser.add_argument("-c", "--console", help="Enable Console Output", action="store_true")
    parser.add_argument("-f", "--file", help="File name for reporting links")
    parser.add_argument("-a", "--ai", help="Use Description and Comments for Epic Health", action="store_true")
    parser.add_argument("-i", "--info", help="Get Boards and Sprints Information", action="store_true")
    parser.add_argument("-p", "--people", help="Get information on people", action="store_true")
    args = parser.parse_args()
    main(args)

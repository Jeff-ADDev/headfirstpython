import os, argparse, json
from typing import List
from colorama import init, Fore, Back, Style
from dotenv import load_dotenv
from datetime import datetime
from objects.epic import Epic
from utils.jira_obj import Jira
from utils.excel_obj import Excel
import console_util
from objects.issue_type import IssueType
from objects.project import Project
from objects.status import Status
from objects.calendar_sprint import CalendarSprint

load_dotenv()
init() # Colorama   

def jira_project_label_reporting(jira, excel, ai_out, date_file_info, path_location, 
                                 con_out, project_label, all_reporting_sprints):
    epics: List[Epic] = []  

    epics = jira.get_epics()
    jira.get_issues(epics)
    
    # Place Points into Sprints
    CalendarSprint.points_in_sprints(epics, all_reporting_sprints)

    # Display Console prior to Excel creation
    if con_out:
        jira.output_console(epics)
    
    if ai_out:
        jira.get_comments(epics)

    wb = excel.create_label_excel_report(epics, all_reporting_sprints)
    save_excel_file = date_file_info + " Project " + project_label + " Details.xlsx"
    console_util.save_excel_file(path_location, save_excel_file, wb)

def jira_boards_sprint_reporting(jira, excel, path_location):

    start_time = datetime.now()
    #start_time_format = start_time.strftime("%m/%d/%Y, %H:%M:%S")
    date_file_info = start_time.strftime("%Y_%m_%d")
    #create_date = start_time.strftime("%m/%d/%Y")

    all_boards = jira.get_boards()
    all_sprints = jira.get_sprints(all_boards)
    all_users = jira.get_users()
    all_projects = jira.get_projects()

    wb = excel.create_jira_info_report(all_boards, all_sprints, all_users, all_projects)
    save_excel_file = date_file_info + " Jira Info.xlsx"
    console_util.save_excel_file(path_location, save_excel_file, wb)

def jira_people_reporting(args):
    print("Getting People Information")

def project_json_load(filelabel):
    begin_sprint = 97
    end_sprint = 200
    json_file = filelabel + ".json"
    with open(json_file) as f:
        data = json.load(f)
    begin_sprint = int(data['start_sprint'])
    end_sprint = int(data['end_sprint'])
    return begin_sprint, end_sprint


def main(args):
    start_time = datetime.now()
    start_time_format = start_time.strftime("%m/%d/%Y, %H:%M:%S")
    date_file_info = start_time.strftime("%Y %m %d")
    create_date = start_time.strftime("%m/%d/%Y")

    jirakey = os.getenv("JIRA_API_KEY")
    claudekey = os.getenv("CLAUDE_API_KEY")
    url_location = os.getenv("JIRA_REV_LOCATION")
    url_search = os.getenv("JIRA_SEARCH")
    url_board = os.getenv("JIRA_BOARD")
    url_issue = os.getenv("JIRA_ISSUE")
    url_users = os.getenv("JIRA_USERS")
    project_label = os.getenv("PROJECT_LABEL")
    path_location = os.getenv("PATH_LOCATION")
    jira_issue_link = os.getenv("JIRA_ISSUE_LINK")
    header = {"Authorization": "Basic " + jirakey}

    ai_out = False
    if args.ai:
        ai_out = True

    # Generate the reporting sprints from the json file
    begin_sprint = 97
    end_sprint = 200
    if args.label:
        project_label = args.label
        begin_sprint, end_sprint = project_json_load(project_label.lower())

    all_reporting_sprints = CalendarSprint.create_calendar_sprints(begin_sprint, end_sprint)

    con_out = False
    if args.console:
        con_out = True

    other_links = {}
    if args.file:
        other_links = console_util.get_links(args.file)
    jira = Jira(project_label, jirakey, url_location, url_search, url_board, url_issue, url_users, header, con_out)
    excel = Excel(claudekey, project_label, jira_issue_link, create_date, ai_out, other_links)
    
    if args.label:
        jira_project_label_reporting(jira, excel, ai_out, date_file_info, path_location, 
                                     con_out, project_label, all_reporting_sprints)
    elif args.info:
        jira_boards_sprint_reporting(jira, excel, path_location)
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

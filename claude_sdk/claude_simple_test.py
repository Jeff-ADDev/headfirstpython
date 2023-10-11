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
from anthropic import Anthropic, HUMAN_PROMPT, AI_PROMPT

start_time = datetime.now()
start_time_format = start_time.strftime("%m/%d/%Y, %H:%M:%S")
date_file_info = start_time.strftime("%Y_%m_%d")
create_date = start_time.strftime("%m/%d/%Y")

load_dotenv()
init() # Colorama   

claudekey = os.getenv("CLAUDE_KEY")

def terminal_update(message, data, bold):
    if bold:
        print(Back.GREEN + Fore.BLACK + Style.BRIGHT + f"  {message}: " + Back.BLUE + Fore.BLACK + Style.BRIGHT + f" {data} " + Style.RESET_ALL, end="\r")
    else:
        print(Fore.GREEN + Style.BRIGHT + f"  {message}: " + Fore.BLUE + Style.NORMAL + f" {data} " + Style.RESET_ALL, end="\r")

def try_claude():
    terminal_update("Trying Cluade", " - ", False)
    anthropic = Anthropic(
        # defaults to os.environ.get("ANTHROPIC_API_KEY")
        api_key=claudekey,
    )
    completion = anthropic.completions.create(
        model="claude-2",
        max_tokens_to_sample=3000,
        prompt=f"{HUMAN_PROMPT}" + 
        """
        I would like to develop a python code that will create an excel sheet for project reporting. 

        """ 
        + f"{AI_PROMPT}",
    )
    print(completion.completion)

def count_tokens(text):
    terminal_update("Claude Counting Tokens", " - ", False)
    anthropic = Anthropic(
        # defaults to os.environ.get("ANTHROPIC_API_KEY")
        api_key=claudekey,
    )
    print(f"\nTokens - {anthropic.count_tokens(text)}")


def main(args):
    #if args.label:
    #    project_label = args.label 
    
    try_claude()

    #count_tokens("how does a court case get to the Supreme Court?")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Create Excel Sheet for Project Reporting")
    parser.add_argument("-l", "--label", help="Label for the project")
    parser.add_argument("-c", "--console", help="Enable Console Output", action="store_true")
    args = parser.parse_args()
    main(args)
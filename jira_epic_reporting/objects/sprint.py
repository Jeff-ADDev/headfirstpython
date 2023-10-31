from typing import List
from colorama import init, Fore, Back, Style
from datetime import datetime

class Sprint:
    """
        Sprint object that represents the sprint in Jira that is attached to a board.

        The Calendar Sprint object is used to represent the sprint encompasing all boardss and
        the universal dates that are used as a company for a sprint.
    """
    def __init__(self, id, name, boardID, state):
        self.id = id
        self.name = name
        self.boardID = boardID
        self.state = state
        self.complete_date = ""
        self.start_date = ""
        self.end_date = ""
        self.story_points_created = 0
        self.story_points_completed = 0

    def set_complete_date(self, complete_date):
        self.complete_date = datetime.strptime(complete_date, "%Y-%m-%dT%H:%M:%S.%fZ")

    def set_start_date(self, start_date):
        self.start_date = datetime.strptime(start_date, "%Y-%m-%dT%H:%M:%S.%fZ")

    def set_end_date(self, end_date):
        self.end_date = datetime.strptime(end_date, "%Y-%m-%dT%H:%M:%S.%fZ")

    def set_story_points_created(self, story_points_created):
        self.story_points_created = story_points_created
    
    def set_story_points_completed(self, story_points_completed):
        self.story_points_completed = story_points_completed

    def get_is_sprint(self, compare_date):
        if self.start_date != "" and self.end_date != "" and compare_date == "":
            if self.start_date <= compare_date and self.end_date >= compare_date:
                return True
        return False
        
    def print_sprint(self):
        print_completeDate = ""
        if self.complete_date != "":
            print_completeDate = "Complete: " + str(self.complete_date.month) + "/" + str(self.complete_date.day) + "/" + str(self.complete_date.year)
        
        print(Fore.GREEN + Style.BRIGHT + 
              "    " + Fore.WHITE + self.name + " " + Style.DIM + str(self.id) + Fore.BLUE + Style.BRIGHT + " Board: " + str(self.boardID) + 
              " " + Fore.WHITE + Style.NORMAL + self.state + 
              " " + Fore.BLUE + print_completeDate +
              Style.RESET_ALL)
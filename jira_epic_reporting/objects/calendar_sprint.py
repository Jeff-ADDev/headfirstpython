from typing import List
from colorama import init, Fore, Back, Style
from datetime import datetime, timedelta
from objects.epic import Epic
from objects.issue import Issue

class CalendarSprint:
    """
        The Calendar Sprint object is used to represent the sprint encompasing all boardss and
        the universal dates that are used as a company for a sprint.

        The Sprint object represents the sprint in Jira that is attached to a board.
    """
    def __init__(self, sprint_number):
        self.name = sprint_number
        self.start_date = ""
        self.end_date = ""
        self.story_points_created = 0
        self.story_points_completed = 0
        self.stories_created = 0

    def set_start_date(self, start_date):
        self.start_date = datetime.strptime(start_date, "%Y-%m-%d")
    
    def set_end_date(self, end_date):
        self.end_date = datetime.strptime(end_date, "%Y-%m-%d")

    def add_story_points_created(self, story_points_created):
        self.story_points_created += story_points_created
    
    def add_story_points_completed(self, story_points_completed):
        self.story_points_completed += story_points_completed

    def add_stories_created(self, stories_created):
        self.stories_created += stories_created

    def get_is_sprint(self, compare_date):
        if self.start_date != "" and self.end_date != "" and compare_date != "":
            # Compare Date is datetimne
            # print(f"{self.start_date} - {self.end_date} - {compare_date}")
            if self.start_date <= compare_date and self.end_date >= compare_date:
                return True
        return False
    
    def points_in_sprints(epics, all_sprints):
        """
            Place Points into Sprints
        """
        for epic in epics:
            for issue in epic.issues:
                # Handle Adding the Points Create Data
                if issue.last_pointchange_date != "":
                    sprint_item = CalendarSprint.get_sprint(all_sprints, datetime.strptime(issue.last_pointchange_date, "%m/%d/%Y"))
                    if issue.size != None and issue.size != "":
                        if sprint_item != None:    
                            sprint_item.add_story_points_created(issue.size)

                # Handle Addint the Points For Done
                if issue.date_done != "":
                    sprint_item = CalendarSprint.get_sprint(all_sprints, datetime.strptime(issue.date_done, "%m/%d/%Y"))
                    if issue.size != None and issue.size != "": 
                        if sprint_item != None:
                            sprint_item.add_story_points_completed(issue.size)
            
                if issue.created != "":
                    sprint_item = CalendarSprint.get_sprint(all_sprints, datetime.strptime(issue.created.strftime("%m/%d/%Y"), "%m/%d/%Y"))
                    if sprint_item != None:
                        sprint_item.add_stories_created(1)
    
    def get_sprint(all_sprints, compare_date):
        for sprint in all_sprints:
            if sprint.get_is_sprint(compare_date):
                return sprint
        return None
    
    def create_calendar_sprints(start_sprint, end_sprint):
        """
            Example of creating calendar sprints
        """
        
        if start_sprint < 97:
            start_sprint = 97
        
        if end_sprint < 98:
            end_sprint = 98

        # 97 5/26/2021 6/8/2021  is the fist sprint
        start_date_97 = datetime.strptime("2021-05-26", "%Y-%m-%d")
        end__date_97 = datetime.strptime("2021-06-08", "%Y-%m-%d")
        all_sprints = []
        return_sprints = []

        for i in range(97, end_sprint):
            sprint = CalendarSprint(i)
            sprint.set_start_date(start_date_97.strftime("%Y-%m-%d"))
            sprint.set_end_date(end__date_97.strftime("%Y-%m-%d"))
            all_sprints.append(sprint)
            start_date_97 += timedelta(days=14)
            end__date_97 += timedelta(days=14)
           
        for i in range(start_sprint, end_sprint):
            for sprint in all_sprints:
                if sprint.name == i:
                    return_sprints.append(sprint)
                    break

        return return_sprints    
        
from typing import List
from datetime import datetime
from colorama import init, Fore, Back, Style

    # Values
    #   id
    #   author
    #       accountId
    #       displayName
    #       emailAddress
    #   created : "2023-09-27T10:16:41.965-0400"
    #   items []
    #       field
    #       fieldtype
    #       fieldId
    #       from
    #       fromString
    #       to
    #       toString

    # For each item we create Changelog

    # field: status - description - IssueParentAssociation - project - Key - link 
    #        Epic Link - issuetype - Sprint - Work Categorization - Story Points 
    #        assignee - resolution - Rank - summary 
    # 
    # field:status - fromString: Ready for Dev - toString: In Development

class Changelog:
    def __init__(self, id, accountId, displayName, emailAddress, created,
                 field, fieldtype, fieldId, fromInfo, fromString, toInfo, toString):
        self.id = id
        self.accountId = accountId
        self.displayName = displayName
        self.emailAddress = emailAddress
        self.created = datetime.strptime(created, "%Y-%m-%dT%H:%M:%S.%f%z")
        self.field = field
        self.fieldtype = fieldtype
        self.fieldId = fieldId
        self.fromInfo = fromInfo
        self.fromString = fromString
        self.toInfo = toInfo
        self.toString = toString

        if self.field is None:
            self.field = "None"
        if self.fieldtype is None:
            self.fieldtype = "None"
        if self.fieldId is None:
            self.fieldId = "None"
        if self.fromInfo is None:
            self.fromInfo = "None"        
        if self.fromString is None:
            self.fromString = "None"
        if self.toString is None:
            self.toString = "None"
        if self.toInfo is None:
            self.toInfo = "None"

    
    def get_create_date(self):
        date = ""
        try:
           date = self.created.strftime("%m/%d/%Y")
        except:
            date = "None"
        return date

    # Utility Functions for ChangeLogs

    def get_first_time(changelogs, fieldvalue, toItem):
        """
        Get the first time to dev date
        """
        first_time = None

        for changelog in changelogs:
            if changelog.field.lower() == fieldvalue.lower():
                if changelog.toString.lower() == toItem.lower():
                    if first_time is None:
                        first_time = changelog.created
                    else:
                        if changelog.created < first_time:
                            first_time = changelog.created
        if first_time is not None:                              
            return datetime.strftime(first_time, "%m/%d/%Y") 
        else:
            return ""
        
    def get_last_time(changelogs, fieldvalue, toItem):
        """
        Get the last time to dev date
        """
        last_time = None

        for changelog in changelogs:
            if changelog.field.lower() == fieldvalue.lower():
                if changelog.toString.lower() == toItem.lower():
                    if last_time is None:
                        last_time = changelog.created
                    else:
                        if changelog.created > last_time:
                            last_time = changelog.created
        if last_time is not None:                              
            return datetime.strftime(last_time, "%m/%d/%Y") 
        else:
            return ""
        
    def get_total_hours(changelogs, fieldvalue, toStartItem, toEndItem):
        """
        Get the total time between
          The first toStartItem
          The last toEndItem 
        """
        first = None
        last = None
        # Find all logs
        all_start_logs: List[Changelog] = []
        all_end_logs: List[Changelog] = []
        for changelog in changelogs:
            if changelog.field.lower() == fieldvalue.lower():
                if changelog.toString.lower() == toStartItem.lower():
                    all_start_logs.append(changelog)

        for changelog in changelogs:
            if changelog.field.lower() == fieldvalue.lower():
                if changelog.toString.lower() == toEndItem.lower():
                    all_end_logs.append(changelog)
        
        # Get the first and last log
        for log in all_start_logs:
            if first is None:
                first = log.created
            else:
                if log.created < first:
                    first = log.created

        for log in all_end_logs:
            if last is None:
                last = log.created
            else:
                if log.created > last:
                    last = log.created

        # differnece 
        if first is not None and last is not None:
            x = (last - first).total_seconds() / 3600
            formatted = f"{x:.2f}"

            return formatted
        else:
            return 0
        
    def get_last_date_point_change(changelogs):
        """
        Get the last date the point change occured
        """
        last_time = None

        for changelog in changelogs:
            if changelog.field.lower() == "story points":
                if last_time is None:
                    last_time = changelog.created
                else:
                    if changelog.created > last_time:
                        last_time = changelog.created
        if last_time is not None:                              
            return datetime.strftime(last_time, "%m/%d/%Y") 
        else:
            return ""

    def get_total_times(changelogs, fieldvalue, toItem):
        """
        Get the total time to dev date
        """
        count = 0
        for changelog in changelogs:
            if changelog.field.lower() == fieldvalue.lower():
                if changelog.toString.lower() == toItem.lower():
                    count += 1  
        return count

    def print_logs(issue, changelogs):
        for changelog in changelogs:
            print(Fore.CYAN + Style.BRIGHT + 
                "   Log ID - " + changelog.id + " - " + changelog.displayName + " - " + changelog.get_create_date() +
                Style.RESET_ALL)
            if changelog.field.lower() == "status":
                print(Fore.YELLOW + Style.BRIGHT + 
                    "     Status - " + changelog.fromString + " --> " + changelog.toString +
                    Style.RESET_ALL)
            elif changelog.field.lower() == "assignee":
                print(Fore.YELLOW + Style.NORMAL + 
                    "     Assignee - " + changelog.fromString + " --> " + changelog.toString +
                    Style.RESET_ALL)
            elif changelog.field.lower() == "story points":
                print(Fore.YELLOW + Style.NORMAL + 
                    "     Story Points - " + changelog.fromString + " --> " + changelog.toString +
                    Style.RESET_ALL)
            else:
                print(Fore.RED + Style.BRIGHT + 
                    "     " + changelog.field +
                    Style.RESET_ALL)
            
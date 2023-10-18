from typing import List
from colorama import init, Fore, Back, Style
from datetime import datetime

class Sprint:
    def __init__(self, id, name, boardID, state):
        self.id = id
        self.name = name
        self.boardID = boardID
        self.state = state
        self.completeDate = ""

    def set_completeDate(self, completeDate):
        self.completeDate = datetime.strptime(completeDate, "%Y-%m-%dT%H:%M:%S.%fZ")

    def print_sprint(self):
        print_completeDate = ""
        if self.completeDate != "":
            print_completeDate = "Complete: " + str(self.completeDate.month) + "/" + str(self.completeDate.day) + "/" + str(self.completeDate.year)
        
        print(Fore.GREEN + Style.BRIGHT + 
              "    " + Fore.WHITE + self.name + " " + Style.DIM + str(self.id) + Fore.BLUE + Style.BRIGHT + " Board: " + str(self.boardID) + 
              " " + Fore.WHITE + Style.NORMAL + self.state + 
              " " + Fore.BLUE + print_completeDate +
              Style.RESET_ALL)
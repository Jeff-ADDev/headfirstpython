from typing import List
from colorama import init, Fore, Back, Style
from sprint import Sprint

class Issue:
    def __init__(self, id, key, summary, size):
        self.id = id
        self.key = key
        self.summary = summary
        self.size = size
        self.sprint: List[Sprint] = []

    def add_sprint(self, sprint):
        self.sprint.append(sprint)

    def set_sprint_name(self, sprint_name):
        self.sprint_name = sprint_name

    def set_sprint_state(self, sprint_state):
        self.sprint_state = sprint_state

    def set_boardID(self, boardID):
        self.boardID = boardID

    def set_completeDate(self, completeDate):
        self.completeDate = completeDate

    def print_Issue(self):
        print(Fore.CYAN + Style.BRIGHT + 
              "  Issue-" + self.key + ": " + Fore.CYAN + Style.NORMAL + self.summary 
              + Fore.BLUE + Style.BRIGHT + " (" + str(self.size) + ")" + Style.RESET_ALL)
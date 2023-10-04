import time, os, requests
from colorama import init, Fore, Back, Style
from dotenv import load_dotenv

start_time = time.time()

load_dotenv()
init() # Colorama   

jirakey = os.getenv("JIRA_API_KEY")
url_location = os.getenv("JIRA_REV_LOCATION")
url_search = os.getenv("JIRA_SEARCH")

main_serach = f"https://{url_location}/{url_search}"
header = {"Authorization": "Basic " + jirakey}

epic_info = main_serach + "'issue'='ARR-2392'"
response = requests.get(epic_info, headers=header)
if response.status_code == 200:
    print(Fore.GREEN + "Success! - Epic Item" + Style.RESET_ALL)
    data = response.json()
    for issue in data["issues"]:
        print(issue["fields"]["summary"])
else:
    print(Fore.RED + "Failed!")

epic_issues = main_serach + "'Epic Link'='ARR-2392' and STATUS != Cancelled"
response = requests.get(epic_issues, headers=header)
if response.status_code == 200:
    print(Fore.GREEN + "Success! - Epic Issues" + Style.RESET_ALL)
    data = response.json()
    for issue in data["issues"]:
        print(Fore.YELLOW + Style.BRIGHT 
              + issue["key"] + " - "
              + issue["fields"]["summary"] + Style.RESET_ALL)
        
        try:
            for item in issue["fields"]["customfield_10010"]:
                cf_10010_name = item["name"]
                cf_10010_state = item["state"]
        except:
                cf_10010_name = ""
                cf_10010_state = ""
        
        print(Fore.LIGHTYELLOW_EX + Style.DIM
              + issue["fields"]["issuetype"]["name"] 
              + " : " + str(issue["fields"]["customfield_10032"])
              + " : " + issue["fields"]["project"]["name"] 
              + " : " + cf_10010_name
              + " : " + cf_10010_state
              + Style.RESET_ALL)

else:
    print(Fore.RED + "Failed!")
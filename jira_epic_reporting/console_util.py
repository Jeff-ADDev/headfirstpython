from colorama import init, Fore, Back, Style
import os
import openpyxl

def terminal_update(message, data, bold):
    print(Style.RESET_ALL + "                                                                                               " + Style.RESET_ALL, end="\r")
    if bold:
        print(Back.GREEN + Fore.BLACK + Style.BRIGHT + f"  {message}: " + Back.BLUE + Fore.BLACK + Style.BRIGHT + f" {data} " + Style.RESET_ALL, end="\r")
    else:
        print(Fore.GREEN + Style.BRIGHT + f"  {message}: " + Fore.BLUE + Style.NORMAL + f" {data} " + Style.RESET_ALL, end="\r")

def terminal_busy(message, count):
    if count > 3:
        count = 0
    data = ""
    if count == 0:
        data = "|"
    elif count == 1:
        data = "/"
    elif count == 2:
        data = "-"
    elif count == 3:
        data = "\\"
    print(Fore.GREEN + Style.BRIGHT + f"  {message}: " + Fore.BLUE + Style.NORMAL + f" {data} " + Style.RESET_ALL, end="\r")

def save_excel_file(path, filename, wb):
    """
    Save the Excel File given the path and filename, along with the workbook
    """
    if os.path.exists(path):
        saveexcelfile = path + filename
    else:
        os.makedirs(path)

    # Check For File Existence - Delete if exists
    if os.path.exists(saveexcelfile):
        os.remove(saveexcelfile)

    # Save Workbook
    wb.save(saveexcelfile)

def get_links(file):
    config = {}
    try:
        with open(file) as f:
            for line in f:
                key, value = line.split("|")
                config[key] = value
        return config
    except:
        return config
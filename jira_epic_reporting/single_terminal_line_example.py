import os, requests, sys, argparse
import time
from colorama import init, Fore, Back, Style

init() # Colorama

def terminal_update(message, data, bold):
    if bold:
        print(Back.GREEN + Fore.BLACK + Style.BRIGHT + f"  {message}: " + Back.BLUE + Fore.BLACK + Style.BRIGHT + f" {data} " + Style.RESET_ALL, end="\r")
    else:
        print(Fore.GREEN + Style.BRIGHT + f"  {message}: " + Fore.BLUE + Style.NORMAL + f" {data} " + Style.RESET_ALL, end="\r")

def test_app():
    for i in range(10):
        terminal_update("Processing Example of Line Data", f"{i}/10", True)
        time.sleep(1)

def main(args):
    test_app()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Create Excel Sheet for Project Reporting")
    parser.add_argument("-l", "--label", help="Label for the project")
    parser.add_argument("-c", "--console", help="Enable Console Output", action="store_true")
    args = parser.parse_args()
    main(args)
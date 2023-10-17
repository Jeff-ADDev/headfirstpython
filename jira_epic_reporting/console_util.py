from colorama import init, Fore, Back, Style

def terminal_update(message, data, bold):
    print(Style.RESET_ALL + "                                                                                               " + Style.RESET_ALL, end="\r")
    if bold:
        print(Back.GREEN + Fore.BLACK + Style.BRIGHT + f"  {message}: " + Back.BLUE + Fore.BLACK + Style.BRIGHT + f" {data} " + Style.RESET_ALL, end="\r")
    else:
        print(Fore.GREEN + Style.BRIGHT + f"  {message}: " + Fore.BLUE + Style.NORMAL + f" {data} " + Style.RESET_ALL, end="\r")

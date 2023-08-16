import os
import webbrowser
import hfpy_utils
from swimclub import get_swim_data, folder

charts = "charts/"

def produce_bar_chart(fn):
    """Given the name of the swimmers file, produce a HTML/SVG based bar chart
    , saving it to the charts folder.  Return the path to the bar chart file."""
    swimmer, age, distance, stroke, _, average_str, times, converts = get_swim_data(fn)
    from_max = max(converts)
    times.reverse()
    converts.reverse()
    title = f"{swimmer} (Under {age}) {distance} {stroke}"
    header = f"""<!DOCTYPE html>
        <html>
            <head>
                <title>{title}</title>
            </head>
            <body>
                <h3>{title}</h3>
    """
    body = ""
    for n, t in enumerate(times):
        bar_width = hfpy_utils.convert2range(converts[n], 0, from_max, 0, 350)
        body += f"""
                    <svg width="400" height="30">
                        <rect height="30" width="{bar_width}" style="fill:rgb(0,0,255);stroke-width:3;stroke:rgb(0,0,0)" />
                    </svg>{t}<br />
                """
    footer = f"""
            <p>Average time: {average_str}</p>
        </body>
    </html>
    """
    image = header + body + footer
    page = header + body + footer
    save_to=f"{folder + charts}{fn.removesuffix('.txt')}.html"
#    with open(save_to, "w") as sf:
#        print(page, file=sf)
    print(f"------ {save_to} -------")
    with open(save_to, "w") as file:
        file.write(page)
    return save_to

import create_chart
import webbrowser

chart = produce_bar_chart("Darius-13-100m-Fly.txt")

webbrowser.open("file://" + os.path.realpath(chart))
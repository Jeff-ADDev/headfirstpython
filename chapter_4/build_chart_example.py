import os
import webbrowser
import hfpy_utils
from swimclub import get_swim_data

fn = "Darius-13-100m-Fly.txt"

swimmer, age, distance, stroke, _, average_str, times, converts = get_swim_data(fn)
from_max = max(converts)

title = f"{swimmer} (Under {age}) {distance} {stroke}"

html = f"""<!DOCTYPE html>
    <html>
        <head>
            <title>{title}</title>
        </head>
        <body>
            <h3>{title}</h3>
"""

svgs = ""
for n, t in enumerate(times):
    bar_width = hfpy_utils.convert2range(converts[n], 0, from_max, 0, 350)
    svgs += f"""
                <svg width="400" height="30">
                    <rect height="30" width="{bar_width}" style="fill:rgb(0,0,255);stroke-width:3;stroke:rgb(0,0,0)" />
                </svg>{t}<br />
            """

footer = f"""
        <p>Average time: {average_str}</p>
    </body>
</html>
"""

page = html + svgs + footer

#webbrowser.open('file://' + os.path.realpath('chapter_4/bar.html'))
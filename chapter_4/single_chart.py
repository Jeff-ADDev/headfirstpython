import os
import create_chart
import webbrowser

chart = create_chart.produce_bar_chart("Lizzie-14-100m-Back.txt")

webbrowser.open("file://" + os.path.realpath(chart))

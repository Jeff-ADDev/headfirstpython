import os
import create_chart
import webbrowser

chart = create_chart.produce_bar_chart("Darius-13-100m-Fly.txt")

webbrowser.open("file://" + os.path.realpath(chart))

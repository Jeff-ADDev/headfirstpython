from swimclub import get_swim_data, folder
import os

swim_files = os.listdir(folder)
swim_files.remove(".DS_Store")

for n, s in enumerate(swim_files, 1):
    print(n, "Processing: " + s)
    get_swim_data(s)

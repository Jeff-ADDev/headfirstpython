import os
import swimclub

swim_files = os.listdir(swimclub.folder)
swim_files.remove('.DS_Store')
swim_files.remove('charts')

swimmers = []
for file in swim_files:
    name = swimclub.get_swim_data(file)[0]
    if name not in swimmers:
        swimmers.append(name)

print(sorted(swimmers))



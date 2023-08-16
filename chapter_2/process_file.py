import statistics

fn = "Darius-13-100m-Fly.txt"
swimmer, age, distance, stroke = fn.removesuffix(".txt").split("-")
folder = "./chapter_2/swimdata/"

with open(folder+fn) as df:
    data = df.readlines()

times = data[0].strip().split(",")

converts = []
for t in times:
    first = t
    minutes, rest = t.split(":")
    seconds, hundredths = rest.split(".")
    converted_time = int(minutes)*60*100 + int(seconds)*100 + int(hundredths)
    converts.append(converted_time)
    print(f"{first} -> {converted_time}")

average = statistics.mean(converts)

min_secs, min_hund = str(round(average / 100, 2)).split(".")
min_secs = int(min_secs)
minutes = min_secs // 60

average_time = f"{minutes}:{min_secs % 60}.{min_hund}"

print(average_time)
import statistics

folder = "./chapter_2/swimdata/"

def get_swim_data(fn):
    """Given the name of a swimmers file, extract all the required data,
    then return it to the caller as a tuple."""
    swimmer, age, distance, stroke = fn.removesuffix(".txt").split("-")
    with open(folder+fn) as df:
        data = df.readlines()
    times = data[0].strip().split(",")
    converts = []
    for t in times:
        if ":" in t:
            minutes, rest = t.split(":")
            seconds, hundredths = rest.split(".")
        else:
            minutes = 0
            seconds, hundredths = t.split(".")
        converts.append(int(minutes)*60*100 + int(seconds)*100 + int(hundredths))
    average = statistics.mean(converts)
    min_secs, min_hund = f"{(average / 100):.2f}".split(".")
    min_secs = int(min_secs)
    minutes = min_secs // 60
    seconds = min_secs % 60
    average_time = f"{minutes}:{seconds}.{min_hund}"
    return swimmer, age, distance, stroke, average, average_time, times, converts

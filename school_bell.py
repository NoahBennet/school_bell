import time
import schedule
import xlrd
import audioplayer


def read_from_excel(path):
    wb = xlrd.open_workbook(path)
    rozvrh = wb.sheet_by_name("Rozvrh")
    zvonenie = wb.sheet_by_name("Zvonenie")
    test = wb.sheet_by_name("Test")

    days = []
    for i in range(3, 8):
        one_day = []
        for j in range(rozvrh.ncols):
            one_day.append(rozvrh.cell_value(i, j))
        days.append(one_day)

    # gets the time data from excel sheet in a format "07:50"
    lesson_bell_times = []
    for i in range(1, 10):
        start_end_time = [str(xlrd.xldate_as_datetime(zvonenie.cell_value(i, 1), wb.datemode)).split()[1][0:-3],
                          str(xlrd.xldate_as_datetime(zvonenie.cell_value(i, 2), wb.datemode)).split()[1][0:-3]]
        lesson_bell_times.append(start_end_time)

    use_first_bell = rozvrh.cell_value(9, 3)
    first_bell_before_start = int(rozvrh.cell_value(10, 3))
    first_bell_before_end = int(rozvrh.cell_value(11, 3))
    bell_volume = int(zvonenie.cell_value(11, 2))
    test_the_bell = test.cell_value(0, 3)
    test_bell_time = str(xlrd.xldate_as_datetime(test.cell_value(1, 5), wb.datemode)).split()[1][0:-3]

    return days, lesson_bell_times, use_first_bell, first_bell_before_start, first_bell_before_end, bell_volume, \
           test_the_bell, test_bell_time


def ring_the_bell(volume):
    # Playback stops when the object is destroyed (GC'ed), so save a reference to the object for non-blocking playback.
    p = audioplayer.AudioPlayer("school_bell_sound.mp3")
    p.volume = volume
    p.play(block=True)


# subtracts from a time given in string (like '10:11') the amount of minutes (15), return a string ('09:56')
def subtract_from_time(given_time, minutes_to_subtract):
    hours = int(given_time.split(":")[0])
    minutes = int(given_time.split(":")[1])
    return convert_min_to_h_and_min((convert_h_and_min_to_min(hours, minutes) - minutes_to_subtract) % 1440)


def convert_h_and_min_to_min(hours, minutes):
    return 60 * hours + minutes


def convert_min_to_h_and_min(minutes):
    return str(minutes // 60).zfill(2) + ":" + str(minutes % 60).zfill(2)


def schedule_bell_times(data_from_excel):
    days = data_from_excel[0]
    lesson_bell_times = data_from_excel[1]
    use_first_bell = data_from_excel[2]
    first_bell_before_start = data_from_excel[3]
    first_bell_before_end = data_from_excel[4]
    bell_volume = data_from_excel[5]
    test_the_bell = data_from_excel[6]
    test_bell_time = data_from_excel[7]

    print("Program zazvoní v týchto časoch:")

    for j in range(5):
        times_at_given_day = []
        for i in range(1, 10):
            if days[j][i] != "nezvoniť":
                if days[j][i] == "zvoniť" or days[j][i] == "zvoniť iba na začiatku":
                    if use_first_bell == "áno pred začiatkom a koncom hodiny" or use_first_bell == "áno iba pred " \
                                                                                                   "začiatkom hodiny":
                        times_at_given_day.append(subtract_from_time(lesson_bell_times[i - 1][0],
                                                                     first_bell_before_start))
                    times_at_given_day.append(lesson_bell_times[i - 1][0])
                if days[j][i] == "zvoniť" or days[j][i] == "zvoniť iba na konci":
                    if use_first_bell == "áno pred začiatkom a koncom hodiny" or use_first_bell == "áno iba pred " \
                                                                                                   "koncom hodiny":
                        times_at_given_day.append(subtract_from_time(lesson_bell_times[i - 1][1],
                                                                     first_bell_before_end))
                    times_at_given_day.append(lesson_bell_times[i - 1][1])
        if j == 0:
            print("Pondelok: \t" + str(times_at_given_day))
            for t in times_at_given_day:
                schedule.every().monday.at(t).do(lambda: ring_the_bell(bell_volume))
        elif j == 1:
            print("Utorok: \t" + str(times_at_given_day))
            for t in times_at_given_day:
                schedule.every().tuesday.at(t).do(lambda: ring_the_bell(bell_volume))
        elif j == 2:
            print("Streda: \t" + str(times_at_given_day))
            for t in times_at_given_day:
                schedule.every().wednesday.at(t).do(lambda: ring_the_bell(bell_volume))
        elif j == 3:
            print("Štvrtok: \t" + str(times_at_given_day))
            for t in times_at_given_day:
                schedule.every().thursday.at(t).do(lambda: ring_the_bell(bell_volume))
        else:
            print("Piatok: \t" + str(times_at_given_day))
            for t in times_at_given_day:
                schedule.every().friday.at(t).do(lambda: ring_the_bell(bell_volume))

    if test_the_bell == "áno":
        schedule.every().day.at(test_bell_time).do(lambda: ring_the_bell(bell_volume))
        print("Každý deň o: " + str(test_bell_time))

    while True:
        schedule.run_pending()
        time.sleep(1)


def main():
    data_from_excel = read_from_excel("skolske zvonenie.xlsx")
    schedule_bell_times(data_from_excel)


if __name__ == '__main__':
    main()

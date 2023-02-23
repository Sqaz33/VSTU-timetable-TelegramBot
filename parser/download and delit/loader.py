import urllib.request, time


def excel_loader(links):
    """загружает таблицы excel"""
    """C:/Users/mset6/OneDrive/Рабочий стол/VSTU-timetable-TelegramBot/data/{file_name}/{excel_name} on pc""" #убрать пр
    """C:/Users/Степан/Documents/GitHub/VSTU-timetable-TelegramBot/data/{file_name}/{excel_name} on laptop"""

    passes = []
    program_start = time.time()
    for i in links:
        for key in i[1]:
            URL = i[1][key]
            file_name = i[0]
            excel_name = key
            destination = f'C:/Users/Степан/Documents/GitHub/VSTU-timetable-TelegramBot/data/{file_name}/{excel_name}' #заменить полный путь на относительный
            passes.append(destination)

            urllib.request.urlretrieve(URL, destination)
            print(f'{excel_name} загружен') #заменить на логгер
    program_stop = time.time()

    return passes, program_stop - program_start
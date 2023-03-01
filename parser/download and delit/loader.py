import urllib.request, time


def excel_loader(links, programm_path):
    """загружает таблицы excel"""


    passes = []
    program_start = time.time()
    for i in links:
        for key in i[1]:
            URL = i[1][key]
            file_name = i[0]
            excel_name = key
            destination = f'{programm_path}VSTU-timetable-TelegramBot/data/{file_name}/{excel_name}' #заменить полный путь на относительный
            passes.append(destination)

            urllib.request.urlretrieve(URL, destination)
            print(f'{excel_name} загружен') #заменить на логгер
    program_stop = time.time()

    return passes, program_stop - program_start


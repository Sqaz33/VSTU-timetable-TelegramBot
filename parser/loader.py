import urllib.request


def excel_loader(links):

    for i in links:
        for key in i[1]:
            URL = i[1][key]
            file_name = i[0]
            excel_name = key
            destination = f'C:/Users/mset6/OneDrive/Рабочий стол/VSTU-timetable-TelegramBot/data/{file_name}/{excel_name}'

            urllib.request.urlretrieve(URL, destination)
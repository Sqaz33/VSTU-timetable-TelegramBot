import urllib.request
file_name = 'timetable.xls'
destination = f'C:/Users/Степан/Desktop/VSTU_timetable_Telegram-bot/timetables/{file_name}'
url = 'https://www.vstu.ru/upload/raspisanie/z/%D0%9E%D0%9D_%D0%A4%D0%90%D0%A1%D0%A2%D0%98%D0%92_1%20%D0%BA%D1%83%D1%80%D1%81%20(%D0%B3%D1%80.%20100%20-%20101).xls'

urllib.request.urlretrieve(url, destination)
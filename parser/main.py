"""
VSTU TimeTableParser v 0.9FFevtO*
================================
*FFevtO - For Fevt Only, Только для ФэВт
"""




import sys, time
sys.path.insert(0, 'C:/Users/Степан/Documents/GitHub/VSTU-timetable-TelegramBot/parser/download and delit')

from loader import excel_loader
from datetime import datetime
from url_parser import links_parser
from del_data import delet_excel
from format_change import changef

sys.path.insert(0, 'C:/Users/Степан/Documents/GitHub/VSTU-timetable-TelegramBot/parser/excel analysis')

from excel_analysis import get_timetable

flag, passes = 1, []
faculties = ['fevt']

if __name__ == '__main__':
    while True:
        if flag:
            program_start = time.time()

            delet_excel(passes)
            print("таблицы Excel удалены")

            links = links_parser(faculties)
            print('ссылки на таблицы с расписание получены')

            passes, program_time = excel_loader(links)
            print('Все файлы загруженны')

            passes = changef(passes)
            print('Все файлы расширения xls заменены на файлы расширения xlsx')
            program_stop = time.time()
            print(f'Лоадер завершил работу за {program_stop - program_start}')

            time, program_time = get_timetable(passes)

            flag = 0
        elif str(datetime.now().minute) == 10:
            flag = 1


"""VSTU TimeTableParser v 0.5 beta"""

import sys, time
sys.path.insert(0, 'C:/Users/mset6/OneDrive/Рабочий стол/VSTU-timetable-TelegramBot/parser/download and delit')

from loader import excel_loader
from datetime import datetime
from url_parser import links_parser
from del_data import delet_excel
from format_change import changef


flag, passes = 1, []
faculties = ['fastiv', 'fat', 'ftkm', 'ftpp', 'feu', 'fevt', 'htf', 'mag']

if __name__ == '__main__':
    while True:
        if str(datetime.now().hour) == '11' or flag:  #заменить после окончания разработки:
            program_start = time.time()

            #удаление предидущих таблиц
            delet_excel(passes)


            #загрузка таблиц
            links = links_parser(faculties)
            passes, program_time = excel_loader(links)

            passes = changef(passes)

            if len(passes) == 104:
                print(f'ВСЕ файлы ЗАГРУЖЕННЫ за {program_time} секунд')  #заменить на логгер
            else:
                print(f'НЕ ВСЕ файлы загружены. Загружены {len(passes)} файлов за {program_time} секунд')  #заменить на логер

            program_stop = time.time()


            flag = 0
            print(f'Лоадер завершил работу за {program_stop-program_start}') #заменить на логгер
        elif str(datetime.now().hour) == '19' and False:
            flag = 1


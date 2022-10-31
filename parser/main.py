import sys
sys.path.insert(0, 'C:/Users/Степан/Documents/GitHub/VSTU-timetable-TelegramBot/parser/download and delit/')

from loader import excel_loader
from datetime import datetime
from url_parser import links_parser
from del_data import delet_excel

flag, passes = 1, []
faculties = ['fastiv', 'fat', 'ftkm', 'ftpp', 'feu', 'fevt', 'htf', 'vkf', 'mmf', 'fpik']

if __name__ == '__main__':
    while True:
        if str(datetime.now().hour) == '15' or flag: #заменить после пр:

            #удаление предидущих таблиц
            delet_excel(passes)


            #загрузка таблиц
            links = links_parser(faculties)
            passes = excel_loader(links)



            flag = 0
        elif str(datetime.now().hour) == '14':
            flag = 1


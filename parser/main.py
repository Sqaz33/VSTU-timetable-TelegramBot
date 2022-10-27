from loader import excel_loader
from datetime import datetime
from url_parser import links_parser
from del_data import delet_excel

flag, passes = 1, []


if __name__ == '__main__':
    while True:
        if str(datetime.now().hour) == '14' and flag:

            #удаление предидущих таблиц
            delet_excel(passes)

            #загрузка таблиц
            links = links_parser()
            passes = excel_loader(links)

            flag = 0
        elif str(datetime.now().hour) == '15':
            flag = 1


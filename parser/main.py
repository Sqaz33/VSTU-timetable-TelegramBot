from loader import excel_loader
from datetime import datetime
from url_parser import links_parser

flag = 1

if __name__ == '__main__':
    while True:
        if str(datetime.now().hour) == '20' and flag:
            links = links_parser()
            excel_loader(links)

            flag = 0
        elif str(datetime.now().hour) == '21':
            flag = 1


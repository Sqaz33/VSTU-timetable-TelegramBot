import os


def delet_excel(passes):
    if len(passes) > 0:

        for i in passes:
            path = os.path.join(os.path.abspath(os.path.dirname(__file__)), i)
            os.remove(path)

        print('таблицы удалены') #заменить на логгер
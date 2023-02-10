import win32com.client as win32
import os


def xls_to_xlsx(fname):
    """не знаю, что он делает: украл с стаковерфлоу"""

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)

    new_path = fname+'x'
    wb.SaveAs(new_path, FileFormat=51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                             #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

    return new_path


def delf(fname):
    """Удаляет предыдущий файл excel"""

    path = os.path.join(os.path.abspath(os.path.dirname(__file__)), fname)
    os.remove(path)


def changef(passes):
    """заменяет в списке passes название пути с xls расширением а новый путь с xlsx расширением"""
    new_passes = passes
    for path in passes:
        if not 'xlsx' in path:
            new_path = xls_to_xlsx(path)
            delf(path)
            new_passes.insert(passes.index(path), new_path)
            new_passes.pop(passes.index(path))

    return new_passes


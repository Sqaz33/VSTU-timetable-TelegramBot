import win32com.client as win32
import os
from xls2xlsx import XLS2XLSX


def xls_to_xlsx(fname):
    """не знаю, что он делает: украл с стаковерфлоу"""

    new_path = fname+'x'
    x2x = XLS2XLSX(fname)
    wb = x2x.to_xlsx(new_path)

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


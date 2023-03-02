import os, pandas




def xls_to_xlsx(fname):
    """не знаю, что он делает: украл с стаковерфлоу"""
    new_path = fname + 'x'
    new_excle = pandas.read_excel(fname)
    new_excle.to_excel(new_path)

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


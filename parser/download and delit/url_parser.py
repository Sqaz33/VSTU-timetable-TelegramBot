import requests, urllib
from bs4 import BeautifulSoup as beati


def links_loader(faculti):
    """Модуль выкачивает ССЫЛКИ на загрузку Excel-таблиц расписаний"""
    ru_letters = 'йцукенгшщзхъфывапролджэячсмитьбюё'

    url = f'https://www.vstu.ru/student/raspisaniya/zanyatiy/index.php?dep={faculti}'
    page = requests.get(url)

    if str(page.status_code) == '200':
        soup = beati(page.text, 'html.parser')
        href = [a['href'] for a in soup.find_all('a', href=True)]
        clean_href = [i[29:] for i in href if any(j.lower() in ru_letters for j in i)]

        return clean_href


def links_packer(faculties):
    """Упаковывает факультеты"""
    load_links = []

    for faculti in faculties:
        links_now = []
        links_now.append(faculti)
        links_now.append(links_loader(faculti))
        load_links.append(links_now)

    return load_links


def links_parser(faculties): #функция высшего порядка
    """модуль переделывающий url"""
    load_links = links_packer(faculties)
    links = []

    for i in load_links:
        faculti = i[0]
        URLs = [f'https://www.vstu.ru/upload/raspisanie/z/{urllib.parse.quote(j)}' for j in i[1]]

        links.append([faculti, dict(zip(i[1], URLs))])

    return links










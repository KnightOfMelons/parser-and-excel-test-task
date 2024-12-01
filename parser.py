import os
import requests
import warnings # эта библиотека у меня используется для игнорирования одной ошибки
from urllib.parse import urljoin
from bs4 import BeautifulSoup


save_folder = "static/excel" # Сюда буду сохранять Excel файлы с сайта (либо можете выбрать свою директорию)
os.makedirs(save_folder, exist_ok=True) # Проверка на то, что есть ли эта папка. Если нет, то скрипт создаст сам
start_date, end_date = 20240902, 20240909 # Это дата с какого периода скачивать и до какаого (20240902: с начала 2024 г, 09 м, 02 д)

# Постоянно в логах ругался на отсутствие SSL сертификата, посему отключил вывод этой ошибки
warnings.simplefilter('ignore', requests.packages.urllib3.exceptions.InsecureRequestWarning)


def download_excel_files(start_date, end_date, save_folder):
    """
    Скачивает Excel файлы с https://www.atsenergo.ru/nreport за указанный период (регион ЕВРОПА).

    Функция проходит по всем датам в диапазоне от `start_date` до `end_date` включительно,
    запрашивает страницу с данными для каждой даты и ищет на ней ссылку на Excel файл.
    Если ссылка найдена, файл скачивается и сохраняется в указанную директорию `save_folder`.

    Аргументы:
    ----------
    start_date : int
        Начальная дата для скачивания файлов в формате YYYYMMDD.
    end_date : int
        Конечная дата для скачивания файлов в формате YYYYMMDD.
    save_folder : str
        Путь к директории, в которую будут сохранены скачанные файлы.

    Примечания:
    -----------
    - Функция игнорирует ошибки SSL сертификата, чтобы избежать предупреждений.
    - Если директория `save_folder` не существует, она будет создана автоматически.

    Пример использования:
    ---------------------
    >>> download_excel_files(20240902, 20240909, "static/excel")
    Файл 20240902_eur_big_nodes_prices_pub.xls скачан
    Файл 20240903_eur_big_nodes_prices_pub.xls скачан
    """
    for date in range(start_date, end_date + 1):
        url = f"https://www.atsenergo.ru/nreport?rname=big_nodes_prices_pub&region=eur&rdate={date}"
        response = requests.get(url, verify=False) # Тут я опять игнорирую SSL сертификат, ибо выдавало ошибки
        soup = BeautifulSoup(response.text, 'html.parser')
        
        link_tag = soup.find('a', title=lambda t: t and t.startswith('Опубликовано'))
        if link_tag:
            file_url = urljoin(url, link_tag['href'])
            file_name = link_tag.get_text()
            
            with open(os.path.join(save_folder, file_name), 'wb') as f:
                f.write(requests.get(file_url, verify=False).content)
            print(f"Файл {file_name} скачан")

download_excel_files(start_date, end_date, save_folder)

import pandas as pd # Не забыть установить (pip install xlrd, pip install openpyxl)
import openpyxl
import logging
from datetime import datetime


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')


def process_price_data(files, output_file="static/result/average_prices.xlsx"):
    """
    Обрабатывает Excel файлы с ценовыми данными для Республики Бурятия,
    рассчитывает средние значения для трёх ценовых столбцов и сохраняет результаты в новый файл Excel.

    Функция проходит по всем переданным файлам, фильтрует данные для Республики Бурятия, 
    рассчитывает средние значения для столбцов: "Равновесная узловая цена, руб./МВт∙ч", 
    "Цена в расчёте без учёта ЦЗСП, руб./МВт∙ч*" и "Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*", 
    а затем сохраняет результат в файл Excel.

    Аргументы:
    ----------
    files : list
        Список путей к Excel файлам, которые необходимо обработать.
    output_file : str, optional
        Путь к файлу, в который будут сохранены результаты. По умолчанию "static/result/average_prices.xlsx".

    Примечания:
    -----------
    - Для каждого файла извлекается дата из имени, и результаты сохраняются в формате "ДД.ММ.ГГГГ".
    - Функция проверяет наличие столбца "Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*", если его нет, выводится предупреждение.
    """
    results = []

    for file in files:
        for sheet_number in range(2, 15):
            df = pd.read_excel(file, sheet_name=sheet_number, skiprows=2)
            df_buryatia = df[df['Субъект РФ'] == 'Республика Бурятия']

            if 'Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*' in df_buryatia.columns: # Решил сделать простую проверку на присутствие этого столбца
                first_price = df_buryatia['Равновесная узловая цена, руб./МВт∙ч'].mean() # С помощью .mean() просто считаем среднее значение для столбцов
                second_price = df_buryatia['Цена в расчёте без учёта ЦЗСП, руб./МВт∙ч*'].mean()
                third_price = df_buryatia['Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*'].mean()

                # Решил использовать небольшой костыль. То есть он берет 20240902 из пути от символа / и до символа _, а затем преобразует в объект 
                # datetime и в формат ДД.ММ.ГГГГ
                date_str = file.split('/')[-1].split('_')[0]
                date_obj = datetime.strptime(date_str, '%Y%m%d')
                formatted_date = date_obj.strftime('%d.%m.%Y')

                # На всякий случай добавил логи, чтобы было понятно, что делает программа (она чуть дольше теперь выполняется к тому же)
                logging.info(f"Обработка файла {file} для даты {formatted_date}")

                results.append([formatted_date, first_price, second_price, third_price])
            else:
                print(f"Столбец 'Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*' отсутствует в файле {file}")

    # Тут собираются все значения по страничкам, получается какая-то такая структура (покажу с 02.09.2024)
    # 02.09.2024	1343,846505	1343,846505	1337,433069
    # 02.09.2024	1364,554526	1364,554526	1361,788641
    # 02.09.2024	1391,828953	1391,828953	1386,12471
    # 02.09.2024	1421,947016	1421,947016	1420,460402
    result_df = pd.DataFrame(results, columns=[
        'Дата',
        'Равновесная узловая цена, руб./МВт∙ч',
        'Цена в расчёте без учёта ЦЗСП, руб./МВт∙ч*',
        'Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*'
    ])

    # А тут уже группирую по столбику Дата и ищу средние значения среди них (вроде как у вас по ТЗ так нужно было)
    final_result = result_df.groupby('Дата').agg({
        'Равновесная узловая цена, руб./МВт∙ч': 'mean',
        'Цена в расчёте без учёта ЦЗСП, руб./МВт∙ч*': 'mean',
        'Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*': 'mean'
    }).reset_index()

    final_result.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Результаты сохранены в файл: {output_file}")


# Я оставил это в таком ужасном виде на случай, если будете пробовать с другими файлами (просто вставьте путь к ним в " ")
files = [
    "static/excel/20240902_eur_big_nodes_prices_pub.xls",
    "static/excel/20240903_eur_big_nodes_prices_pub.xls",
    "static/excel/20240904_eur_big_nodes_prices_pub.xls",
    "static/excel/20240905_eur_big_nodes_prices_pub.xls",
    "static/excel/20240906_eur_big_nodes_prices_pub.xls",
    "static/excel/20240907_eur_big_nodes_prices_pub.xls",
    "static/excel/20240908_eur_big_nodes_prices_pub.xls",
    "static/excel/20240909_eur_big_nodes_prices_pub.xls"
]

process_price_data(files, "static/result/average_prices.xlsx")

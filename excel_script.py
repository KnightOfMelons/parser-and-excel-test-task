import pandas as pd # Не забыть установить (pip install xlrd)

# Я оставил это в таком ужасном виде на случай, если будете пробовать с другими файлами
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


for file in files:
    df = pd.read_excel(file, skiprows=2)
    
    df_buryatia = df[df['Субъект РФ'] == 'Республика Бурятия']
    

    if 'Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*' in df_buryatia.columns: # Решил сделать простую проверку на присутствие этого столбца
        first_price = df_buryatia['Равновесная узловая цена, руб./МВт∙ч'].mean() # С помощью .mean() просто считаем среднее значение для столбцов
        second_price = df_buryatia['Цена в расчёте без учёта ЦЗСП, руб./МВт∙ч*'].mean() 
        third_price = df_buryatia['Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*'].mean()
        
        print(f"Среднее значение для файла {file}:\nРавновесная узловая цена, руб./МВт∙ч: {first_price}\nЦена в расчёте без учёта ЦЗСП: {second_price}\nЦена в расчёте с учётом ЦЗСП: {third_price}\n")
    else:
        print(f"Столбец 'Цена в расчёте с учётом ЦЗСП, руб./МВт∙ч*' отсутствует в файле {file}")

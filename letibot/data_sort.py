from imports import *

old_file_path = r'.\.output\applicants.xlsx'
new_file_path = 'результат_распределения.xlsx'

def data_sort():
    keys = ['Конкурсный балл','∑ балл','Предмет 1','Предмет 2','Предмет 3','ИД']
    values = [False, False, False, False, False, False]
    sheets = pd.ExcelFile(old_file_path).sheet_names
    true_sheets = ["Целевое", "Бюджет", "Контракт"]
    for i in true_sheets:
        for sheet in sheets:
            if i == sheet:
                if sheet == 'Целевое':
                    keys.append('Приоритет №')
                    values.append(True)
                    df = pd.read_excel(old_file_path, sheet_name='Целевое')
                    df2 = df.sort_values(by=keys,ascending=values)
                    free_places = {
                        "01.03.02": 10,
                        "09.03.01": 15,
                        "09.03.01_alt": 11,
                        "09.03.02": 16,
                        "09.03.04": 10,
                        "10.05.01": 30,
                        "11.03.01": 19,
                        "11.03.01_alt": 19,
                        "11.03.02": 7,
                        "11.03.03": 23,
                        "11.03.04": 32,
                        "11.05.01": 28,
                        "12.03.01": 25,
                        "12.03.04": 2,
                        "13.03.02": 21,
                        "15.03.06": 1,
                        "20.03.01": 2,
                        "27.03.02": 3,
                        "27.03.03": 1,
                        "27.03.04": 2,
                        "27.03.04_alt": 4,
                        "27.03.05": 1,
                        "28.03.01": 3,
                        "45.03.02": 2
                    }

                    df2['Результат'] = 0
                    result_df = df2.copy()
                    for index, row in df2.iterrows():
                        if index not in result_df.index:
                            continue
                            
                        unique_code = row['Уникальный код поступающего']
                        direction = row['Код направления']
                        consent = row['Согласие на зачисление']

                        if (direction in free_places and free_places[direction] > 0 and consent in ['Бумажное', 'Электронное']):
                            result_df['Результат'] = result_df['Результат'].astype(str)
                            result_df.at[index, 'Результат'] = direction
                            free_places[direction] -= 1
                            result_df = result_df[(result_df['Уникальный код поступающего'] != unique_code) | 
                                                (result_df.index == index)]

                    file_exists = Path(new_file_path).exists()
                    mode = 'a' if file_exists else 'w'
                    with pd.ExcelWriter(new_file_path,engine='openpyxl',mode=mode,if_sheet_exists='replace') as writer:
                        result_df.to_excel(writer, sheet_name='Целевое', index=False)
                    keys = ['Конкурсный балл','∑ балл','Предмет 1','Предмет 2','Предмет 3','ИД']
                    values = [False, False, False, False, False, False]
                    print(f"Лист 'Целевое' обновлен")
                elif sheet == 'Бюджет':
                    keys.append('ПП')
                    values.append(True)
                    keys.append('Приоритет №')
                    values.append(True)
                    df = pd.read_excel(old_file_path, sheet_name='Бюджет')
                    filtered_df = df[df['Условия зачисления'] == 'Основные места']
                    df2 = filtered_df.sort_values(by=keys,ascending=values)
                    free_places = {
                        "01.03.02": 100,
                        "09.03.01": 100,
                        "09.03.01_alt": 70,
                        "09.03.02": 104,
                        "09.03.04": 68,
                        "10.05.01": 75,
                        "11.03.01": 38,
                        "11.03.01_alt": 54,
                        "11.03.02": 73,
                        "11.03.03": 50,
                        "11.03.04": 222,
                        "11.05.01": 40,
                        "12.03.01": 126,
                        "12.03.04": 55,
                        "13.03.02": 105,
                        "15.03.06": 4,
                        "20.03.01": 30,
                        "27.03.02": 25,
                        "27.03.03": 20,
                        "27.03.04": 80,
                        "27.03.04_alt": 43,
                        "27.03.05": 14,
                        "28.03.01": 63,
                        "42.03.01": 1,
                        "45.03.02": 12
                    }

                    df2['Результат'] = 0
                    result_df = df2.copy()
                    for index, row in df2.iterrows():
                        if index not in result_df.index:
                            continue
                            
                        unique_code = row['Уникальный код поступающего']
                        direction = row['Код направления']
                        consent = row['Согласие на зачисление']

                        if (direction in free_places and free_places[direction] > 0 and consent in ['Бумажное', 'Электронное']):
                            result_df['Результат'] = result_df['Результат'].astype(str)
                            result_df.at[index, 'Результат'] = direction
                            free_places[direction] -= 1
                            result_df = result_df[(result_df['Уникальный код поступающего'] != unique_code) | 
                                                (result_df.index == index)]

                    file_exists = Path(new_file_path).exists()
                    mode = 'a' if file_exists else 'w'
                    with pd.ExcelWriter(new_file_path,engine='openpyxl',mode=mode,if_sheet_exists='replace') as writer:
                        result_df.to_excel(writer, sheet_name='Бюджет', index=False)
                    keys = ['Конкурсный балл','∑ балл','Предмет 1','Предмет 2','Предмет 3','ИД']
                    values = [False, False, False, False, False, False]
                    print(f"Лист 'Бюджет' обновлен")
                elif sheet == 'Контракт':
                    keys.append('ПП')
                    values.append(True)
                    keys.append('Приоритет №')
                    values.append(True)
                    df = pd.read_excel(old_file_path, sheet_name='Контракт')
                    filtered_df = df[df['Условия зачисления'] == 'Контракт']
                    df2 = df.sort_values(by=keys,ascending=values)
                    contract_places = {
                        "01.03.02": 80,
                        "09.03.01": 80,
                        "09.03.01_alt": 80,
                        "09.03.02": 106,
                        "09.03.04": 82,
                        "10.05.01": 75,
                        "11.03.01": 15,
                        "11.03.01_alt": 6,
                        "11.03.02": 20,
                        "11.03.03": 10,
                        "11.03.04": 40,
                        "11.05.01": 40,
                        "12.03.01": 10,
                        "12.03.04": 25,
                        "13.03.02": 25,
                        "15.03.06": 50,
                        "20.03.01": 20,
                        "27.03.02": 35,
                        "27.03.03": 10,
                        "27.03.04": 30,
                        "27.03.04_alt": 47,
                        "27.03.05": 45,
                        "28.03.01": 10,
                        "42.03.01": 130,
                        "45.03.02": 110
                    }
                    df2['Результат'] = 0
                    result_df = df2.copy()
                    for index, row in df2.iterrows():
                        if index not in result_df.index:
                            continue
                            
                        unique_code = row['Уникальный код поступающего']
                        direction = row['Код направления']
                        consent = row['Статус договора']

                        if (direction in contract_places and contract_places[direction] > 0 and consent in ['Заключён', 'Оплачен']):
                            result_df['Результат'] = result_df['Результат'].astype(str)
                            result_df.at[index, 'Результат'] = direction
                            contract_places[direction] -= 1
                            result_df = result_df[(result_df['Уникальный код поступающего'] != unique_code) | 
                                                (result_df.index == index)]

                    file_exists = Path(new_file_path).exists()
                    mode = 'a' if file_exists else 'w'
                    with pd.ExcelWriter(new_file_path,engine='openpyxl',mode=mode,if_sheet_exists='replace') as writer:
                        result_df.to_excel(writer, sheet_name='Контракт', index=False)
                    keys = ['Конкурсный балл','∑ балл','Предмет 1','Предмет 2','Предмет 3','ИД']
                    values = [False, False, False, False, False, False]
                    print(f"Лист 'Контракт' обновлен")
                else:
                    print(f"Обнаружен неизвестный лист {sheet}")
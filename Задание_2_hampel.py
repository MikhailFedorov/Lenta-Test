import pandas as pd
import numpy as np

# Загрузка Excel-файла
file_path = 'Лента_тестовое_задание.xlsx'
file1 = pd.read_excel(file_path, sheet_name='файл1', dtype={'Товар': str})
file2 = pd.read_excel(file_path, sheet_name='файл2', dtype={'Товар': str})
file3 = pd.read_excel(file_path, sheet_name='файл3', dtype={'Товар': str})


# Функция для чтения последних 6 цифр кода товара
def last_digits(code):
    """Возвращает последние 6 цифр кода товара

        Аргументы:
            code: изначальный код товара

        Возвращает:
            Строку, содержащую последние 6 цифр кода товара
        """
    return str(code)[-6:]


# Создание временных столбцов с унифицированным кодом в 6 символов
file1['Нормализованный_Товар'] = file1['Товар'].apply(last_digits)
file2['Нормализованный_Товар'] = file2['Товар'].apply(last_digits)
file3['Нормализованный_Товар'] = file3['Товар'].apply(last_digits)

# Объединение таблиц методом left-join по городу и коду
merged_df = file1.merge(file2[['Кластер', 'Нормализованный_Товар']],
                        how='left',
                        left_on=['Кластер', 'Нормализованный_Товар'],
                        right_on=['Кластер', 'Нормализованный_Товар'],
                        indicator='file2_merge')
merged_df = merged_df.merge(file3[['Кластер', 'Нормализованный_Товар']],
                            how='left',
                            left_on=['Кластер', 'Нормализованный_Товар'],
                            right_on=['Кластер', 'Нормализованный_Товар'],
                            indicator='file3_merge')


# Функция для определения типа мониторинга
def monitoring_type(row):
    """Получает значения типов мониторинга в зависимости от значения индикатора.
       indicator = both в том случае, если при присоединении таблицы произошло слияние как левого,
       так и правого значения, что означает, что код товара из файла 1 совпал с кодом из файла 2 или 3

            Аргументы:
                row: строка Excel

            Возвращает:
                Тип мониторинга (Тип_1, Тип_2 или Тип_3)
            """
    if row['file2_merge'] == 'both':
        return 'Тип_1'
    elif row['file3_merge'] == 'both':
        return 'Тип_2'
    else:
        return 'Тип_3'


# Применение функции и добавление столбца "Тип мониторинга"
merged_df['Тип мониторинга'] = merged_df.apply(monitoring_type, axis=1)

# Удаление временных столбцов
merged_df.drop(columns=['Нормализованный_Товар', 'file2_merge', 'file3_merge'], inplace=True)

# Сохранение результата в новый Excel-файл
output_file_path = 'Лента_тестовое_задание_1.xlsx'
merged_df.to_excel(output_file_path, index=False)

type_2_df = merged_df[merged_df['Тип мониторинга'] == 'Тип_2'].copy()
merged_df_2 = merged_df


# Функция для очистки выбросов методом фильтра Хэмпеля
def hampel_filter(data, column, window_size=7, n_sigmas=3):
    """Удаляет выбросы из выборки методом фильтра Хэмпеля. Так как при удалении выбросов методом IQR не зафиксировалось
       ни одного выброса, была совершена еще одна проверка другим методом, более чувствительным к данным
       содержащим импульсные шумы.

            Аргументы:
                data: файл данных
                column: конкретная колонка в файле
                window_size, n_sigmas: параметры фильтра

            Возвращает:
                Датафрейм с удаленными выбросами
            """
    new_data = data[column].copy()
    k = 1.4826  # Константа для нормального распределения

    for i in range(window_size, len(data) - window_size):
        window = data[column][i - window_size:i + window_size + 1]
        median = window.median()
        mad = k * np.median(np.abs(window - median))
        threshold = n_sigmas * mad
        if np.abs(data[column].iloc[i] - median) > threshold:
            new_data.iloc[i] = median

    return new_data


# Применение фильтра Хэмпеля для столбцов Цена_мин и Цена_ср
type_2_df['Цена_мин'] = hampel_filter(type_2_df, 'Цена_мин')
type_2_df['Цена_ср'] = hampel_filter(type_2_df, 'Цена_ср')

# Обновление основной таблицы с очищенными данными для Тип_2
merged_df.update(type_2_df)

# Сохранение результата в новый Excel-файл
output_file_path = 'Лента_тестовое_задание_2_hampel.xlsx'
merged_df.to_excel(output_file_path, index=False)

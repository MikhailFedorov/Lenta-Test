import pandas as pd

# Загрузка Excel-файла
file_path = 'Лента_тестовое_задание.xlsx'
file1 = pd.read_excel(file_path, sheet_name='файл1', dtype={'Товар': object, 'Цена_мин': int})
file2 = pd.read_excel(file_path, sheet_name='файл2', dtype={'Товар': object, 'Цена_мин': int})
file3 = pd.read_excel(file_path, sheet_name='файл3', dtype={'Товар': object, 'Цена_мин': int})


# -------------------------Задание 1------------------------------

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
merged_df.drop_duplicates(subset=None, keep="first", inplace=True)

# Сохранение результата в новый Excel-файл
output_file_path = 'Лента_тестовое_задание_1.xlsx'
merged_df.to_excel(output_file_path, index=False)

# -------------------------Задание 2------------------------------

type_2_df = merged_df[merged_df['Тип мониторинга'] == 'Тип_2'].copy()
merged_df_2 = merged_df


# Функция для очистки выбросов методом IQR
def IQR(data, column):
    """Удаляет выбросы из выборки методом Межквартильного размаха. Был выбран именно этот метод
       из-за его надежности и простоты работы (не нужно самому определять порог, как в Z-score или считать много разных
       характеристик, как для метода Фишера.

            Аргументы:
                data: файл данных
                column: конкретная колонка в файле

            Возвращает:
                Датафрейм с удаленными выбросами
            """
    Q1 = data[column].quantile(0.25)
    Q3 = data[column].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    print(lower_bound, upper_bound)
    return data[(data[column] >= lower_bound) & (data[column] <= upper_bound)]


# Применение функции для столбцов Цена_мин и Цена_ср
cleaned_type_2_df = IQR(type_2_df, 'Цена_мин')
cleaned_type_2_df = IQR(cleaned_type_2_df, 'Цена_ср')

# Обновление основной таблицы с очищенными данными для Тип_2
merged_df_2.update(cleaned_type_2_df)


# -------------------------Задание 3------------------------------

# Выгрузка в новый файл
output_file_path_2 = 'Лента_тестовое_задание_2.xlsx'
merged_df_2.to_excel(output_file_path_2, index=False)

# выгрузка в новый файл с разбиением таблиц по городам
with pd.ExcelWriter('Лента_тестовое_задание_3.xlsx') as writer:
    for city in merged_df_2['Кластер'].unique():
        city_df = merged_df_2[merged_df_2['Кластер'] == city]
        city_df.to_excel(writer, sheet_name=city, index=False)

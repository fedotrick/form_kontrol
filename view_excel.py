import pandas as pd

# Просмотр plavka.xlsx
try:
    df_plavka = pd.read_excel('plavka.xlsx')
    print("== Структура plavka.xlsx ==")
    print(df_plavka.columns.tolist())
    print("\n== Проверка наличия столбца 'Номер_кластера' ==")
    if 'Номер_кластера' in df_plavka.columns:
        print("Столбец 'Номер_кластера' существует в файле plavka.xlsx")
        
        # Добавим тестовое значение для первой записи, чтобы проверить работу
        print("\n== Добавление тестового значения для проверки ==")
        df_plavka.loc[0, 'Номер_кластера'] = "K-001"
        print(f"Учетный номер: {df_plavka.loc[0, 'Учетный_номер']}")
        print(f"Номер кластера: {df_plavka.loc[0, 'Номер_кластера']}")
        
        # Сохраняем обратно файл с изменением для тестирования
        df_plavka.to_excel('plavka.xlsx', index=False)
        print("Данные сохранены для тестирования")
    else:
        print("Столбец 'Номер_кластера' отсутствует в файле plavka.xlsx")
        print("Доступные столбцы:")
        for col in df_plavka.columns:
            print(f"- {col}")
except Exception as e:
    print(f"Ошибка при работе с plavka.xlsx: {e}")

# Просмотр control.xlsx
try:
    df_control = pd.read_excel('control.xlsx')
    print("\n\n== Структура control.xlsx ==")
    print(df_control.columns.tolist())
    print("\n== Первые 5 строк control.xlsx ==")
    print(df_control.head())
except Exception as e:
    print(f"Ошибка при чтении control.xlsx: {e}") 
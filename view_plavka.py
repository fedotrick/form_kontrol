import pandas as pd

# Проверяем загрузку конкретной плавки
try:
    # Загрузка данных из plavka.xlsx
    df_plavka = pd.read_excel('plavka.xlsx')
    
    # Вывод колонок и первой строки
    print("Колонки файла:")
    print(df_plavka.columns.tolist())
    print("\nПервая строка данных:")
    print(df_plavka.iloc[0])
    
    # Поиск плавки с номером 9-1/23
    search_number = '9-1/23'
    mask = df_plavka['Учетный_номер'].astype(str) == search_number
    
    if mask.any():
        row = df_plavka.loc[mask].iloc[0]
        print(f"\nДанные плавки {search_number}:")
        print(f"  Учетный номер: {row['Учетный_номер']}")
        
        if 'Номер_кластера' in df_plavka.columns:
            print(f"  Номер кластера: {row['Номер_кластера']} (тип: {type(row['Номер_кластера'])})")
            if pd.isna(row['Номер_кластера']):
                print("  Номер кластера является NaN значением")
            
        print(f"  Наименование отливки: {row['Наименование_отливки']}")
    else:
        print(f"Плавка с номером {search_number} не найдена")
    
    # Добавляем тестовое значение снова
    df_plavka.loc[df_plavka['Учетный_номер'].astype(str) == '9-1/23', 'Номер_кластера'] = "K-001"
    df_plavka.to_excel('plavka.xlsx', index=False)
    print("\nТестовое значение 'K-001' для номера кластера добавлено заново")
    
except Exception as e:
    print(f"Ошибка при чтении plavka.xlsx: {e}") 
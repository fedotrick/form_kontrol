import pandas as pd

# Просмотр plavka.xlsx
try:
    df_plavka = pd.read_excel('plavka.xlsx')
    print("== Структура plavka.xlsx ==")
    print(df_plavka.columns.tolist())
    print("\n== Первые 5 строк plavka.xlsx ==")
    print(df_plavka.head())
except Exception as e:
    print(f"Ошибка при чтении plavka.xlsx: {e}")

# Просмотр control.xlsx
try:
    df_control = pd.read_excel('control.xlsx')
    print("\n\n== Структура control.xlsx ==")
    print(df_control.columns.tolist())
    print("\n== Первые 5 строк control.xlsx ==")
    print(df_control.head())
except Exception as e:
    print(f"Ошибка при чтении control.xlsx: {e}") 
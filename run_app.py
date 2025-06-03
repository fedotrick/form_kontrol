import sys
import os
from kontrol import ControlForm, QApplication

def main():
    """
    Главная функция для запуска приложения
    """
    # Проверяем наличие необходимых файлов
    required_files = ['Отчёт.xlsx', 'Сводный.xlsx']
    for file in required_files:
        if not os.path.exists(file):
            print(f"Ошибка: Файл {file} не найден!")
            input("Нажмите Enter для выхода...")
            return

    # Запускаем приложение
    try:
        app = QApplication(sys.argv)
        form = ControlForm()
        form.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"Ошибка при запуске приложения: {str(e)}")
        input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main() 
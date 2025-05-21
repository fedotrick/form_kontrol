import sys
import os
from datetime import datetime
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                              QHBoxLayout, QPushButton, QLabel, QDateEdit, 
                              QLineEdit, QFileDialog, QMessageBox, QGroupBox,
                              QListWidget, QComboBox, QCheckBox)
from PySide6.QtCore import Qt, QDate

from report_generator_helper import ReportHelper

class ReportGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Генератор отчетов")
        self.setMinimumSize(600, 400)
        
        # Создаем центральный виджет и компоновку
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        
        # Группа настроек отчета
        settings_group = QGroupBox("Настройки отчета")
        settings_layout = QVBoxLayout(settings_group)
        
        # Выбор даты
        date_layout = QHBoxLayout()
        date_label = QLabel("Дата отчета:")
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        date_layout.addWidget(date_label)
        date_layout.addWidget(self.date_edit)
        settings_layout.addLayout(date_layout)
        
        # Контролеры с опцией авто-извлечения
        controllers_group = QVBoxLayout()
        controllers_auto_layout = QHBoxLayout()
        self.auto_controllers_checkbox = QCheckBox("Автоматически извлечь контролеров из файла")
        self.auto_controllers_checkbox.setChecked(True)
        self.auto_controllers_checkbox.stateChanged.connect(self.toggle_controllers_input)
        controllers_auto_layout.addWidget(self.auto_controllers_checkbox)
        controllers_group.addLayout(controllers_auto_layout)
        
        controllers_layout = QHBoxLayout()
        controllers_label = QLabel("Контролеры:")
        self.controllers_edit = QLineEdit()
        self.controllers_edit.setPlaceholderText("Введите имена контролеров через запятую")
        self.controllers_edit.setEnabled(False)  # По умолчанию отключено
        controllers_layout.addWidget(controllers_label)
        controllers_layout.addWidget(self.controllers_edit)
        controllers_group.addLayout(controllers_layout)
        
        settings_layout.addLayout(controllers_group)
        
        # Выбор шаблона отчета
        template_layout = QHBoxLayout()
        template_label = QLabel("Шаблон отчета:")
        self.template_edit = QLineEdit()
        self.template_edit.setText("Отчёт.xlsx")
        self.template_edit.setReadOnly(True)
        template_button = QPushButton("Обзор...")
        template_button.clicked.connect(self.select_template)
        template_layout.addWidget(template_label)
        template_layout.addWidget(self.template_edit)
        template_layout.addWidget(template_button)
        settings_layout.addLayout(template_layout)
        
        # Выбор файла контроля
        control_layout = QHBoxLayout()
        control_label = QLabel("Файл контроля:")
        self.control_edit = QLineEdit()
        self.control_edit.setText("control.xlsx")
        self.control_edit.setReadOnly(True)
        control_button = QPushButton("Обзор...")
        control_button.clicked.connect(self.select_control_file)
        control_layout.addWidget(control_label)
        control_layout.addWidget(self.control_edit)
        control_layout.addWidget(control_button)
        settings_layout.addLayout(control_layout)
        
        # Добавляем настройки в основную компоновку
        main_layout.addWidget(settings_group)
        
        # Группа кнопок генерации отчета
        buttons_layout = QHBoxLayout()
        
        # Кнопка генерации отчета
        generate_button = QPushButton("Сгенерировать отчет")
        generate_button.clicked.connect(self.generate_report)
        generate_button.setMinimumHeight(40)
        buttons_layout.addWidget(generate_button)
        
        # Кнопка просмотра отчета
        view_button = QPushButton("Просмотреть отчет")
        view_button.clicked.connect(self.view_report)
        view_button.setMinimumHeight(40)
        buttons_layout.addWidget(view_button)
        
        # Добавляем кнопки в основную компоновку
        main_layout.addLayout(buttons_layout)
        
        # Статус
        self.status_label = QLabel("Готов к работе")
        main_layout.addWidget(self.status_label)
        
        # Устанавливаем центральный виджет
        self.setCentralWidget(central_widget)
        
        # Последний сгенерированный отчет
        self.last_report_path = None
    
    def toggle_controllers_input(self, state):
        """Переключает доступность поля ввода контролеров"""
        self.controllers_edit.setEnabled(not state)
    
    def select_template(self):
        """Выбор файла шаблона отчета"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите шаблон отчета", "", "Excel файлы (*.xlsx *.xls)")
        
        if file_path:
            self.template_edit.setText(file_path)
    
    def select_control_file(self):
        """Выбор файла контроля"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл контроля", "", "Excel файлы (*.xlsx *.xls)")
        
        if file_path:
            self.control_edit.setText(file_path)
    
    def get_controllers(self, helper, control_path, date):
        """Получает список контролеров автоматически или из поля ввода"""
        if self.auto_controllers_checkbox.isChecked():
            # Автоматическое извлечение из файла control.xlsx
            controllers = helper.get_controllers_from_excel(control_path, date)
            if not controllers:
                reply = QMessageBox.question(
                    self, "Контролеры не найдены", 
                    "В файле не найдены контролеры. Хотите ввести их вручную?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                
                if reply == QMessageBox.Yes:
                    # Временно включаем поле ввода
                    self.controllers_edit.setEnabled(True)
                    QMessageBox.information(self, "Внимание", 
                                           "Введите имена контролеров через запятую и нажмите 'Сгенерировать отчет' снова.")
                    return None
                else:
                    return []
            return controllers
        else:
            # Ручной ввод
            return [c.strip() for c in self.controllers_edit.text().split(",") if c.strip()]
    
    def generate_report(self):
        """Генерация отчета"""
        try:
            # Получаем настройки
            date = self.date_edit.date().toString("dd.MM.yyyy")
            template_path = self.template_edit.text()
            control_path = self.control_edit.text()
            
            # Проверяем наличие файлов
            if not os.path.exists(template_path):
                QMessageBox.warning(self, "Ошибка", f"Файл шаблона не найден: {template_path}")
                return
            
            if not os.path.exists(control_path):
                QMessageBox.warning(self, "Ошибка", f"Файл контроля не найден: {control_path}")
                return
            
            # Создаем помощник
            helper = ReportHelper(template_path)
            
            # Получаем контролеров
            controllers = self.get_controllers(helper, control_path, date)
            if controllers is None:
                # Пользователь будет вводить контролеров вручную
                return
            
            # Формируем имя выходного файла
            output_date = self.date_edit.date().toString("dd-MM-yyyy")
            output_path = f"Отчет_{output_date}.xlsx"
            
            # Обновляем статус
            self.status_label.setText(f"Генерация отчета для даты {date}...")
            QApplication.processEvents()
            
            # Устанавливаем дату и контролеров
            helper.set_date(date)
            helper.set_controllers(controllers)
            
            # Загружаем данные из файла контроля
            success = helper.load_from_control_xlsx(control_path, date)
            
            if not success:
                QMessageBox.warning(self, "Предупреждение", 
                                   "Возникли проблемы при загрузке данных из файла контроля. "
                                   "Отчет может быть неполным.")
            
            # Сохраняем отчет
            self.last_report_path = helper.save(output_path)
            
            # Обновляем статус
            controllers_str = ", ".join(controllers)
            self.status_label.setText(f"Отчет успешно сгенерирован с контролерами: {controllers_str}")
            
            # Показываем сообщение об успехе
            QMessageBox.information(self, "Успех", f"Отчет успешно сгенерирован и сохранен:\n{self.last_report_path}")
            
        except Exception as e:
            # В случае ошибки показываем сообщение
            QMessageBox.critical(self, "Ошибка", f"Ошибка при генерации отчета: {str(e)}")
            self.status_label.setText(f"Ошибка: {str(e)}")
    
    def view_report(self):
        """Открытие последнего сгенерированного отчета"""
        if self.last_report_path and os.path.exists(self.last_report_path):
            # Открываем файл с помощью системного приложения по умолчанию
            os.startfile(self.last_report_path)
        else:
            # Предлагаем сгенерировать отчет
            reply = QMessageBox.question(
                self, "Отчет не найден", 
                "Отчет еще не сгенерирован или был перемещен. Сгенерировать новый отчет?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                self.generate_report()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ReportGeneratorApp()
    window.show()
    sys.exit(app.exec()) 
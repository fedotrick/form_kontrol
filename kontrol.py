import sys
import os
import pandas as pd
import traceback
from datetime import datetime, timedelta
import shutil
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QFormLayout, QLineEdit,
    QDateEdit, QPushButton, QMessageBox, QGroupBox, QLabel, QScrollArea, QComboBox, QHBoxLayout, QGraphicsDropShadowEffect,
    QDialog, QCalendarWidget, QRadioButton, QButtonGroup, QFileDialog, QTextBrowser, QTableWidget, QTableWidgetItem,
    QHeaderView, QTabWidget, QCheckBox, QMainWindow
)
from PySide6.QtCore import QDate, Qt, QPropertyAnimation, QEasingCurve, QEvent, QDateTime, QRegularExpression, QTimer
from PySide6 import QtGui
from PySide6.QtGui import QFont, QColor, QRegularExpressionValidator, QIcon, QTextCursor
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from functools import partial

class ControlForm(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Электронный журнал контроля")
        self.setGeometry(0, 0, 800, 1100)

        layout = QVBoxLayout()
        layout.setSpacing(5)
        layout.setContentsMargins(5, 5, 5, 5)

        # Добавляем общие стили в стиле Dracula
        self.setStyleSheet("""
            QWidget {
                background-color: #282a36;
                color: #f8f8f2;
                font-family: 'Segoe UI', 'Aptos';
                font-size: 11px;
            }
            
            QGroupBox {
                border: 1px solid #44475a;
                border-radius: 4px;
                margin-top: 0.5em;
                padding: 8px;
            }
            
            QLineEdit, QComboBox, QDateEdit {
                background-color: #44475a;
                color: #f8f8f2;
                border: 1px solid #6272a4;
                border-radius: 2px;
                padding: 3px;
                min-width: 80px;
                max-width: 120px;
                height: 20px;
            }
            
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus {
                border: 2px solid #bd93f9;
            }
            
            QLabel {
                color: #f8f8f2;
                padding: 2px;
            }
            
            QPushButton {
                background-color: #6272a4;
                color: #f8f8f2;
                border: none;
                padding: 5px 10px;
                border-radius: 2px;
                height: 25px;
            }
            
            QPushButton:hover {
                background-color: #bd93f9;
            }
            
            QPushButton:pressed {
                background-color: #ff79c6;
            }
            
            QScrollArea {
                border: none;
            }
            
            QScrollBar:vertical {
                background-color: #282a36;
                width: 14px;
                margin: 15px 0;
            }
            
            QScrollBar::handle:vertical {
                background-color: #44475a;
                min-height: 30px;
                border-radius: 7px;
            }
            
            QScrollBar::handle:vertical:hover {
                background-color: #6272a4;
            }
            
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            
            QComboBox QAbstractItemView {
                background-color: #44475a;
                color: #f8f8f2;
                selection-background-color: #6272a4;
            }
        """)

        # Создание области прокрутки
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        
        # Основной виджет для прокрутки
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        # Блок 1: Основные данные
        group_box1 = QGroupBox("Основные данные")
        group_box1.setStyleSheet("""
            QGroupBox::title {
                color: #bd93f9;
            }
        """)
        form_layout1 = QFormLayout()
        form_layout1.setSpacing(3)
        form_layout1.setContentsMargins(5, 5, 5, 5)
        
        # Выпадающий список для номера плавки
        self.номер_плавки_input = QComboBox(self)
        self.load_plavka_numbers()
        
        # Добавляем поле для отображения наименования отливки (только для чтения)
        self.наименование_отливки_input = QLineEdit(self)
        self.наименование_отливки_input.setReadOnly(True)  # Только для чтения
        self.наименование_отливки_input.setStyleSheet("""
            QLineEdit {
                background-color: #44475a;
                color: #f8f8f2;
                border: 1px solid #6272a4;
                border-radius: 2px;
                padding: 3px;
                min-width: 200px;
                max-width: 300px;
                height: 20px;
            }
        """)
        
        # Добавляем поле для отображения номера кластера (только для чтения)
        self.номер_кластера_input = QLineEdit(self)
        self.номер_кластера_input.setReadOnly(True)  # Только для чтения
        self.номер_кластера_input.setStyleSheet("""
            QLineEdit {
                background-color: #44475a;
                color: #f8f8f2;
                border: 1px solid #6272a4;
                border-radius: 2px;
                padding: 3px;
                min-width: 120px;
                max-width: 150px;
                height: 20px;
            }
        """)
        
        # Список участников
        persons = [
            "Елхова", "Лабуткина", "Рябова", "Улитина"            
        ]
        # Сортировка списка участников по возрастанию
        persons.sort()
        
        # Создаем основные поля ввода
        self.контроль_отлито_input = QLineEdit(self)
        self.контроль_принято_input = QLineEdit(self)
        self.контроль_принято_input.setReadOnly(True)  # Только для чтения
        
        # Создаем комбобоксы для контролеров
        self.контролер1_input = QComboBox(self)
        self.контролер1_input.addItems(persons)  # Добавляем участников в комбобокс
        self.контролер1_input.setFont(QtGui.QFont("Aptos", 12, QtGui.QFont.Bold))
        self.контролер1_input.setStyleSheet("color: white;")
        self.контролер1_input.setCurrentIndex(-1)  # Ничего не выбрано по умолчанию
        
        self.контролер2_input = QComboBox(self)
        self.контролер2_input.addItems(persons)  # Добавляем участников в комбобокс
        self.контролер2_input.setFont(QtGui.QFont("Aptos", 12, QtGui.QFont.Bold))
        self.контролер2_input.setStyleSheet("color: white;")
        self.контролер2_input.setCurrentIndex(-1)

        self.контролер3_input = QComboBox(self)
        self.контролер3_input.addItems(persons)  # Добавляем участников в комбобокс
        self.контролер3_input.setFont(QtGui.QFont("Aptos", 12, QtGui.QFont.Bold))
        self.контролер3_input.setStyleSheet("color: white;")
        self.контролер3_input.setCurrentIndex(-1)
        
        # Настраиваем поле даты
        self.контроль_дата_приемки_input = QDateEdit(self)
        self.контроль_дата_приемки_input.setDisplayFormat("dd.MM.yyyy")
        self.контроль_дата_приемки_input.setDate(QDate.currentDate())  # Текущая дата по умолчанию
        self.контроль_дата_приемки_input.setCalendarPopup(True)  # Разрешаем всплывающий календарь
        self.контроль_дата_приемки_input.setStyleSheet("""
            QDateEdit {
                padding: 8px;
                font-size: 12px;
                min-width: 120px;
                border-radius: 4px;
                background-color: #44475a;
                color: #f8f8f2;
            }
        """)
        
        # Добавляем остальные поля в форму
        form_layout1.addRow(QLabel("Дата приемки:"), self.контроль_дата_приемки_input)
        form_layout1.addRow(QLabel("Номер плавки:"), self.номер_плавки_input)
        form_layout1.addRow(QLabel("Наименование отливки:"), self.наименование_отливки_input)
        form_layout1.addRow(QLabel("Номер кластера:"), self.номер_кластера_input)
        form_layout1.addRow(QLabel("Отлито, шт.:"), self.контроль_отлито_input)
        form_layout1.addRow(QLabel("Принято, шт.:"), self.контроль_принято_input)
        form_layout1.addRow(QLabel("Контролер 1:"), self.контролер1_input)
        form_layout1.addRow(QLabel("Контролер 2:"), self.контролер2_input)
        form_layout1.addRow(QLabel("Контролер 3:"), self.контролер3_input)
        
        group_box1.setLayout(form_layout1)

        # Блок 2: Второй сорт
        group_box2 = QGroupBox("Второй сорт")
        group_box2.setStyleSheet("""
            QGroupBox::title {
                color: #50fa7b;
            }
        """)
        form_layout2 = QFormLayout()
        form_layout2.setSpacing(3)
        form_layout2.setContentsMargins(5, 5, 5, 5)
        
        self.второй_сорт_раковины_input = QLineEdit(self)
        # Удаляем старое поле второй_сорт_зарез_input и создаем два новых поля
        self.второй_сорт_зарез_литейный_input = QLineEdit(self)
        self.второй_сорт_зарез_пеномодельный_input = QLineEdit(self)

        form_layout2.addRow(QLabel("Раковины:"), self.второй_сорт_раковины_input)
        form_layout2.addRow(QLabel("Зарез литейный:"), self.второй_сорт_зарез_литейный_input)
        form_layout2.addRow(QLabel("Зарез пеномодельный:"), self.второй_сорт_зарез_пеномодельный_input)

        group_box2.setLayout(form_layout2)

        # Блок 3: Доработка
        group_box3 = QGroupBox("Доработка")
        group_box3.setStyleSheet("""
            QGroupBox::title {
                color: #ffb86c;
            }
        """)
        form_layout3 = QFormLayout()
        form_layout3.setSpacing(3)
        form_layout3.setContentsMargins(5, 5, 5, 5)
        
        self.доработка_раковины_input = QLineEdit(self)
        self.доработка_раковины_input.hide()  # Hide the field
        self.доработка_зарез_input = QLineEdit(self)
        self.доработка_зарез_input.hide()  # Hide the field
        self.доработка_несоответствие_размеров_input = QLineEdit(self)
        self.доработка_несоответствие_внешнего_вида_input = QLineEdit(self)
        self.доработка_наплыв_металла_input = QLineEdit(self)
        self.доработка_прорыв_металла_input = QLineEdit(self)
        self.доработка_вырыв_input = QLineEdit(self)
        self.доработка_облой_input = QLineEdit(self)
        self.доработка_песок_на_поверхности_input = QLineEdit(self)
        self.доработка_песок_в_резьбе_input = QLineEdit(self)
        # Удаляем старое поле клея и создаем два новых
        self.доработка_клей_подтёк_input = QLineEdit(self)
        self.доработка_клей_по_шву_input = QLineEdit(self)
        self.доработка_коробление_input = QLineEdit(self)
        self.доработка_дефект_пеномодели_input = QLineEdit(self)
        self.доработка_лапы_input = QLineEdit(self)
        self.доработка_питатель_input = QLineEdit(self)
        self.доработка_корона_input = QLineEdit(self)
        self.доработка_смещение_input = QLineEdit(self)

        # Add rows without the hidden fields
        form_layout3.addRow(QLabel("Несоответствие размеров:"), self.доработка_несоответствие_размеров_input)
        form_layout3.addRow(QLabel("Несоответствие внешнего вида:"), self.доработка_несоответствие_внешнего_вида_input)
        form_layout3.addRow(QLabel("Наплыв металла:"), self.доработка_наплыв_металла_input)
        form_layout3.addRow(QLabel("Прорыв металла:"), self.доработка_прорыв_металла_input)
        form_layout3.addRow(QLabel("Вырыв:"), self.доработка_вырыв_input)
        form_layout3.addRow(QLabel("Облой:"), self.доработка_облой_input)
        form_layout3.addRow(QLabel("Песок на поверхности:"), self.доработка_песок_на_поверхности_input)
        form_layout3.addRow(QLabel("Песок в резьбе:"), self.доработка_песок_в_резьбе_input)
        form_layout3.addRow(QLabel("Клей подтёк:"), self.доработка_клей_подтёк_input)
        form_layout3.addRow(QLabel("Клей по шву:"), self.доработка_клей_по_шву_input)
        form_layout3.addRow(QLabel("Коробление:"), self.доработка_коробление_input)
        form_layout3.addRow(QLabel("Дефект пеномодели:"), self.доработка_дефект_пеномодели_input)
        form_layout3.addRow(QLabel("Лапы:"), self.доработка_лапы_input)
        form_layout3.addRow(QLabel("Питатель:"), self.доработка_питатель_input)
        form_layout3.addRow(QLabel("Корона:"), self.доработка_корона_input)
        form_layout3.addRow(QLabel("Смещение:"), self.доработка_смещение_input)

        group_box3.setLayout(form_layout3)

        # Блок 4: Окончательный брак
        group_box4 = QGroupBox("Окончательный брак")
        group_box4.setStyleSheet("""
            QGroupBox::title {
                color: #ff5555;
            }
        """)
        form_layout4 = QFormLayout()
        form_layout4.setSpacing(3)
        form_layout4.setContentsMargins(5, 5, 5, 5)
        
        self.окончательный_брак_недолив_input = QLineEdit(self)
        self.окончательный_брак_раковины_input = QLineEdit(self)
        self.окончательный_брак_коробление_input = QLineEdit(self)
        self.окончательный_брак_спай_input = QLineEdit(self)
        self.окончательный_брак_трещины_input = QLineEdit(self)
        self.окончательный_брак_пригар_песка_input = QLineEdit(self)
        self.окончательный_брак_пористость_input = QLineEdit(self)
        self.окончательный_брак_вырыв_input = QLineEdit(self)
        self.окончательный_брак_скол_input = QLineEdit(self)
        self.окончательный_брак_слом_input = QLineEdit(self)
        # Заменяем старое поле зареза на два новых
        self.окончательный_брак_зарез_литейный_input = QLineEdit(self)
        self.окончательный_брак_зарез_пеномодельный_input = QLineEdit(self)
        self.окончательный_брак_нарушение_геометрии_input = QLineEdit(self)
        self.окончательный_брак_рыхлота_input = QLineEdit(self)
        self.окончательный_брак_непроклей_input = QLineEdit(self)
        self.окончательный_брак_пеномодель_input = QLineEdit(self)
        self.окончательный_брак_наплыв_металла_input = QLineEdit(self)
        self.окончательный_брак_несоответствие_размеров_input = QLineEdit(self)
        self.окончательный_брак_несоответствие_внешнего_вида_input = QLineEdit(self)
        self.окончательный_брак_нарушение_маркировки_input = QLineEdit(self)
        self.окончательный_брак_неслитина_input = QLineEdit(self)
        self.окончательный_брак_прочее_input = QLineEdit(self)

        form_layout4.addRow(QLabel("Недолив:"), self.окончательный_брак_недолив_input)
        form_layout4.addRow(QLabel("Раковины:"), self.окончательный_брак_раковины_input)
        form_layout4.addRow(QLabel("Коробление:"), self.окончательный_брак_коробление_input)
        form_layout4.addRow(QLabel("Спай:"), self.окончательный_брак_спай_input)
        form_layout4.addRow(QLabel("Трещины:"), self.окончательный_брак_трещины_input)
        form_layout4.addRow(QLabel("Пригар песка:"), self.окончательный_брак_пригар_песка_input)
        form_layout4.addRow(QLabel("Пористость:"), self.окончательный_брак_пористость_input)
        form_layout4.addRow(QLabel("Вырыв:"), self.окончательный_брак_вырыв_input)
        form_layout4.addRow(QLabel("Скол:"), self.окончательный_брак_скол_input)
        form_layout4.addRow(QLabel("Слом:"), self.окончательный_брак_слом_input)
        form_layout4.addRow(QLabel("Зарез литейный:"), self.окончательный_брак_зарез_литейный_input)
        form_layout4.addRow(QLabel("Зарез пеномодельный:"), self.окончательный_брак_зарез_пеномодельный_input)
        form_layout4.addRow(QLabel("Нарушение геометрии:"), self.окончательный_брак_нарушение_геометрии_input)
        form_layout4.addRow(QLabel("Рыхлота:"), self.окончательный_брак_рыхлота_input)
        form_layout4.addRow(QLabel("Непроклей:"), self.окончательный_брак_непроклей_input)
        form_layout4.addRow(QLabel("Пеномодель:"), self.окончательный_брак_пеномодель_input)
        form_layout4.addRow(QLabel("Наплыв металла:"), self.окончательный_брак_наплыв_металла_input)
        form_layout4.addRow(QLabel("Несоответствие размеров:"), self.окончательный_брак_несоответствие_размеров_input)
        form_layout4.addRow(QLabel("Несоответствие внешнего вида:"), self.окончательный_брак_несоответствие_внешнего_вида_input)
        form_layout4.addRow(QLabel("Нарушение маркировки:"), self.окончательный_брак_нарушение_маркировки_input)
        form_layout4.addRow(QLabel("Неслитина:"), self.окончательный_брак_неслитина_input)
        form_layout4.addRow(QLabel("Прочее:"), self.окончательный_брак_прочее_input)

        group_box4.setLayout(form_layout4)

        # Создаем горизонтальные layout для группировки полей
        h_layout = QHBoxLayout()
        
        # Группируем GroupBox'ы по два в ряд
        left_column = QVBoxLayout()
        left_column.addWidget(group_box1)
        left_column.addWidget(group_box3)
        
        right_column = QVBoxLayout()
        right_column.addWidget(group_box2)
        right_column.addWidget(group_box4)
        
        h_layout.addLayout(left_column)
        h_layout.addLayout(right_column)
        
        scroll_layout.addLayout(h_layout)

        # Установка виджета прокрутки
        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)

        # Кнопка для сохранения данных
        self.save_button = QPushButton("Сохранить", self)

        # Установка параметров кнопки
        self.save_button.setStyleSheet("""
            QPushButton {
                background-color: #50fa7b;
                color: #282a36;
                font-size: 14px;
                padding: 12px 30px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #69ff94;
            }
            QPushButton:pressed {
                background-color: #41d66b;
            }
        """)

        # Подключение сигнала
        self.save_button.clicked.connect(self.save_data)

        # Добавление кнопки в layout
        layout.addWidget(self.save_button)
        
        # Добавляем кнопку для формирования отчетов
        self.report_button = QPushButton("Сформировать отчет", self)
        self.report_button.setStyleSheet("""
            QPushButton {
                background-color: #bd93f9;
                color: #282a36;
                font-size: 14px;
                padding: 12px 30px;
                font-weight: bold;
                margin-top: 5px;
            }
            QPushButton:hover {
                background-color: #d6acff;
            }
            QPushButton:pressed {
                background-color: #a775f0;
            }
        """)
        self.report_button.clicked.connect(self.show_report_dialog)
        layout.addWidget(self.report_button)

        self.setLayout(layout)

        # Подключение события изменения для расчета контроль_принято
        self.контроль_отлито_input.textChanged.connect(self.calculate_control_prinato)
        self.второй_сорт_раковины_input.textChanged.connect(self.calculate_control_prinato)
        self.второй_сорт_зарез_литейный_input.textChanged.connect(self.calculate_control_prinato)
        self.второй_сорт_зарез_пеномодельный_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_раковины_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_зарез_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_несоответствие_размеров_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_несоответствие_внешнего_вида_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_наплыв_металла_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_прорыв_металла_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_вырыв_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_облой_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_песок_на_поверхности_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_песок_в_резьбе_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_клей_подтёк_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_клей_по_шву_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_коробление_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_дефект_пеномодели_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_лапы_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_питатель_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_корона_input.textChanged.connect(self.calculate_control_prinato)
        self.доработка_смещение_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_недолив_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_раковины_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_коробление_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_спай_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_трещины_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_пригар_песка_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_пористость_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_вырыв_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_скол_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_слом_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_зарез_литейный_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_зарез_пеномодельный_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_нарушение_геометрии_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_рыхлота_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_непроклей_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_пеномодель_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_наплыв_металла_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_несоответствие_размеров_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_несоответствие_внешнего_вида_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_нарушение_маркировки_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_неслитина_input.textChanged.connect(self.calculate_control_prinato)
        self.окончательный_брак_прочее_input.textChanged.connect(self.calculate_control_prinato)

        # Добавляем анимацию при наведении на группы
        for group in [group_box1, group_box2, group_box3, group_box4]:
            group.enterEvent = lambda e, g=group: self.animate_group_hover(g, True)
            group.leaveEvent = lambda e, g=group: self.animate_group_hover(g, False)

        # Добавляем тени для групп
        for group in [group_box1, group_box2, group_box3, group_box4]:
            shadow = QGraphicsDropShadowEffect()
            shadow.setBlurRadius(15)
            shadow.setColor(QColor(0, 0, 0, 30))
            shadow.setOffset(0, 2)
            group.setGraphicsEffect(shadow)

        # Добавляем валидацию для числовых полей
        numeric_inputs = [
            self.контроль_отлито_input,
            self.второй_сорт_раковины_input,
            self.второй_сорт_зарез_литейный_input,
            self.второй_сорт_зарез_пеномодельный_input,
            self.доработка_раковины_input,
            self.доработка_зарез_input,
            self.доработка_несоответствие_размеров_input,
            self.доработка_несоответствие_внешнего_вида_input,
            self.доработка_наплыв_металла_input,
            self.доработка_прорыв_металла_input,
            self.доработка_вырыв_input,
            self.доработка_облой_input,
            self.доработка_песок_на_поверхности_input,
            self.доработка_песок_в_резьбе_input,
            self.доработка_клей_подтёк_input,
            self.доработка_клей_по_шву_input,
            self.доработка_коробление_input,
            self.доработка_дефект_пеномодели_input,
            self.доработка_лапы_input,
            self.доработка_питатель_input,
            self.доработка_корона_input,
            self.доработка_смещение_input,
            self.окончательный_брак_недолив_input,
            self.окончательный_брак_вырыв_input,
            self.окончательный_брак_зарез_литейный_input,
            self.окончательный_брак_зарез_пеномодельный_input,
            self.окончательный_брак_коробление_input,
            self.окончательный_брак_наплыв_металла_input,
            self.окончательный_брак_нарушение_геометрии_input,
            self.окончательный_брак_нарушение_маркировки_input,
            self.окончательный_брак_непроклей_input,
            self.окончательный_брак_неслитина_input,
            self.окончательный_брак_несоответствие_внешнего_вида_input,
            self.окончательный_брак_несоответствие_размеров_input,
            self.окончательный_брак_пеномодель_input,
            self.окончательный_брак_пористость_input,
            self.окончательный_брак_пригар_песка_input,
            self.окончательный_брак_прочее_input,
            self.окончательный_брак_рыхлота_input,
            self.окончательный_брак_раковины_input,
            self.окончательный_брак_скол_input,
            self.окончательный_брак_слом_input,
            self.окончательный_брак_спай_input,
            self.окончательный_брак_трещины_input
        ]
        
        for input_field in numeric_inputs:
            input_field.textChanged.connect(
                lambda text, field=input_field: field.setText(''.join(filter(str.isdigit, text)))
            )

        # Подключаем обработчик изменения номера плавки
        self.номер_плавки_input.currentTextChanged.connect(self.update_наименование_отливки)

        # Список всех интерактивных виджетов для навигации
        self.focusable_widgets = [
            self.контроль_дата_приемки_input,
            self.номер_плавки_input,
            self.контроль_отлито_input,
            self.контроль_принято_input,
            self.контролер1_input,
            self.контролер2_input,
            self.контролер3_input,
            self.второй_сорт_раковины_input,
            self.второй_сорт_зарез_литейный_input,
            self.второй_сорт_зарез_пеномодельный_input,
            self.доработка_несоответствие_размеров_input,
            self.доработка_несоответствие_внешнего_вида_input,
            self.доработка_наплыв_металла_input,
            self.доработка_прорыв_металла_input,
            self.доработка_вырыв_input,
            self.доработка_облой_input,
            self.доработка_песок_на_поверхности_input,
            self.доработка_песок_в_резьбе_input,
            self.доработка_клей_подтёк_input,
            self.доработка_клей_по_шву_input,
            self.доработка_коробление_input,
            self.доработка_дефект_пеномодели_input,
            self.доработка_лапы_input,
            self.доработка_питатель_input,
            self.доработка_корона_input,
            self.доработка_смещение_input,
            self.окончательный_брак_недолив_input,
            self.окончательный_брак_раковины_input,
            self.окончательный_брак_коробление_input,
            self.окончательный_брак_спай_input,
            self.окончательный_брак_трещины_input,
            self.окончательный_брак_пригар_песка_input,
            self.окончательный_брак_пористость_input,
            self.окончательный_брак_вырыв_input,
            self.окончательный_брак_скол_input,
            self.окончательный_брак_слом_input,
            self.окончательный_брак_зарез_литейный_input,
            self.окончательный_брак_зарез_пеномодельный_input,
            self.окончательный_брак_нарушение_геометрии_input,
            self.окончательный_брак_рыхлота_input,
            self.окончательный_брак_непроклей_input,
            self.окончательный_брак_пеномодель_input,
            self.окончательный_брак_наплыв_металла_input,
            self.окончательный_брак_несоответствие_размеров_input,
            self.окончательный_брак_несоответствие_внешнего_вида_input,
            self.окончательный_брак_нарушение_маркировки_input,
            self.окончательный_брак_неслитина_input,
            self.окончательный_брак_прочее_input
        ]

        # Фильтруем список, исключая скрытые виджеты
        self.focusable_widgets = [widget for widget in self.focusable_widgets if not widget.isHidden()]
        
        # Устанавливаем обработку клавиш для каждого виджета
        for widget in self.focusable_widgets:
            widget.installEventFilter(self)

        # Устанавливаем фокус на первое поле при запуске
        self.контроль_дата_приемки_input.setFocus()

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and obj in self.focusable_widgets:
            key = event.key()
            current_index = self.focusable_widgets.index(obj)
            
            if key == Qt.Key_Up:
                # Переход к предыдущему виджету (циклично)
                next_index = (current_index - 1) % len(self.focusable_widgets)
                self.focusable_widgets[next_index].setFocus()
                return True
                
            elif key == Qt.Key_Down:
                # Переход к следующему виджету (циклично)
                next_index = (current_index + 1) % len(self.focusable_widgets)
                self.focusable_widgets[next_index].setFocus()
                return True
                
        return super().eventFilter(obj, event)

    def load_plavka_numbers(self):
        """Загружает номера плавок из Excel файла"""
        if not os.path.exists('plavka.xlsx'):
            QMessageBox.warning(self, "Ошибка", "Файл plavka.xlsx не найден")
            return
            
        try:
            # Загрузка данных из plavka.xlsx
            self.df_plavka = pd.read_excel('plavka.xlsx')  # Сохраняем DataFrame как атрибут класса
            
            # Фильтрация номеров, содержащих "/25"
            self.df_plavka = self.df_plavka[self.df_plavka['Учетный_номер'].astype(str).str.contains('/25')]
            
            # Загрузка данных из control.xlsx, если файл существует
            try:
                df_control = pd.read_excel('control.xlsx')
                # Получение списка уже использованных номеров плавок
                used_numbers = df_control['Номер_плавки'].astype(str).unique()
                # Фильтрация, исключая использованные номера
                self.df_plavka = self.df_plavka[~self.df_plavka['Учетный_номер'].astype(str).isin(used_numbers)]
            except FileNotFoundError:
                pass
            
            # Очищаем комбобокс перед добавлением новых номеров
            self.номер_плавки_input.clear()
            
            # Добавление отфильтрованных номеров в комбобокс
            available_numbers = self.df_plavka['Учетный_номер'].astype(str).tolist()
            self.номер_плавки_input.addItems(available_numbers)
            
            QMessageBox.information(self, "Информация", 
                                  f"Доступно номеров плавок: {len(available_numbers)}")
            
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Ошибка при загрузке номеров плавок: {str(e)}")

    def update_наименование_отливки(self, select_number=None):
        """Обновляет поле наименования отливки и номера кластера на основе выбранного номера плавки"""
        try:
            if not hasattr(self, 'df_plavka') or self.df_plavka is None:
                if not os.path.exists('plavka.xlsx'):
                    self.наименование_отливки_input.clear()
                    self.номер_кластера_input.clear()
                    return
                # Загружаем данные заново, если они не загружены
                self.df_plavka = pd.read_excel('plavka.xlsx')
            
            if select_number is None:
                select_number = self.номер_плавки_input.currentText()

            # Если номер плавки пустой, очищаем поля
            if not select_number:
                self.наименование_отливки_input.clear()
                self.номер_кластера_input.clear()
                return

            # Ищем соответствующие значения
            mask = self.df_plavka['Учетный_номер'].astype(str) == str(select_number)
            if mask.any():
                # Получаем наименование отливки
                наименование = self.df_plavka.loc[mask, 'Наименование_отливки'].iloc[0]
                self.наименование_отливки_input.setText(str(наименование))
                
                # Получаем номер кластера
                if 'Номер_кластера' in self.df_plavka.columns:
                    номер_кластера = self.df_plavka.loc[mask, 'Номер_кластера'].iloc[0]
                    # Проверяем, что номер кластера не NaN и не None
                    if pd.notna(номер_кластера):
                        self.номер_кластера_input.setText(str(номер_кластера))
                    else:
                        self.номер_кластера_input.setText("")
                else:
                    self.номер_кластера_input.setText("")
            else:
                self.наименование_отливки_input.clear()
                self.номер_кластера_input.clear()
                
        except Exception as e:
            self.наименование_отливки_input.clear()
            self.номер_кластера_input.clear()
            QMessageBox.warning(self, "Ошибка", f"Ошибка при обновлении данных: {str(e)}")

    def calculate_control_prinato(self):
        try:
            контроль_отлито = int(self.контроль_отлито_input.text() or 0)
            второй_сорт_раковины = int(self.второй_сорт_раковины_input.text() or 0)
            второй_сорт_зарез_литейный = int(self.второй_сорт_зарез_литейный_input.text() or 0)
            второй_сорт_зарез_пеномодельный = int(self.второй_сорт_зарез_пеномодельный_input.text() or 0)
            # Убираем скрытые поля из расчетов
            # доработка_раковины = int(self.доработка_раковины_input.text() or 0)
            # доработка_зарез = int(self.доработка_зарез_input.text() or 0)
            доработка_несоответствие_размеров = int(self.доработка_несоответствие_размеров_input.text() or 0)
            доработка_несоответствие_внешнего_вида = int(self.доработка_несоответствие_внешнего_вида_input.text() or 0)
            доработка_наплыв_металла = int(self.доработка_наплыв_металла_input.text() or 0)
            доработка_прорыв_металла = int(self.доработка_прорыв_металла_input.text() or 0)
            доработка_вырыв = int(self.доработка_вырыв_input.text() or 0)
            доработка_облой = int(self.доработка_облой_input.text() or 0)
            доработка_песок_на_поверхности = int(self.доработка_песок_на_поверхности_input.text() or 0)
            доработка_песок_в_резьбе = int(self.доработка_песок_в_резьбе_input.text() or 0)
            доработка_клей_подтёк = int(self.доработка_клей_подтёк_input.text() or 0)
            доработка_клей_по_шву = int(self.доработка_клей_по_шву_input.text() or 0)
            доработка_коробление = int(self.доработка_коробление_input.text() or 0)
            доработка_дефект_пеномодели = int(self.доработка_дефект_пеномодели_input.text() or 0)
            доработка_лапы = int(self.доработка_лапы_input.text() or 0)
            доработка_питатель = int(self.доработка_питатель_input.text() or 0)
            доработка_корона = int(self.доработка_корона_input.text() or 0)
            доработка_смещение = int(self.доработка_смещение_input.text() or 0)
            
            # Добавление окончательных браков в расчет
            окончательный_брак_недолив = int(self.окончательный_брак_недолив_input.text() or 0)
            окончательный_брак_раковины = int(self.окончательный_брак_раковины_input.text() or 0)
            окончательный_брак_коробление = int(self.окончательный_брак_коробление_input.text() or 0)
            окончательный_брак_спай = int(self.окончательный_брак_спай_input.text() or 0)
            окончательный_брак_трещины = int(self.окончательный_брак_трещины_input.text() or 0)
            окончательный_брак_пригар_песка = int(self.окончательный_брак_пригар_песка_input.text() or 0)
            окончательный_брак_пористость = int(self.окончательный_брак_пористость_input.text() or 0)
            окончательный_брак_вырыв = int(self.окончательный_брак_вырыв_input.text() or 0)
            окончательный_брак_скол = int(self.окончательный_брак_скол_input.text() or 0)
            окончательный_брак_слом = int(self.окончательный_брак_слом_input.text() or 0)
            окончательный_брак_зарез_литейный = int(self.окончательный_брак_зарез_литейный_input.text() or 0)
            окончательный_брак_зарез_пеномодельный = int(self.окончательный_брак_зарез_пеномодельный_input.text() or 0)
            окончательный_брак_нарушение_геометрии = int(self.окончательный_брак_нарушение_геометрии_input.text() or 0)
            окончательный_брак_рыхлота = int(self.окончательный_брак_рыхлота_input.text() or 0)
            окончательный_брак_непроклей = int(self.окончательный_брак_непроклей_input.text() or 0)
            окончательный_брак_пеномодель = int(self.окончательный_брак_пеномодель_input.text() or 0)
            окончательный_брак_наплыв_металла = int(self.окончательный_брак_наплыв_металла_input.text() or 0)
            окончательный_брак_несоответствие_размеров = int(self.окончательный_брак_несоответствие_размеров_input.text() or 0)
            окончательный_брак_несоответствие_внешнего_вида = int(self.окончательный_брак_несоответствие_внешнего_вида_input.text() or 0)
            окончательный_брак_нарушение_маркировки = int(self.окончательный_брак_нарушение_маркировки_input.text() or 0)
            окончательный_брак_неслитина = int(self.окончательный_брак_неслитина_input.text() or 0)
            окончательный_брак_прочее = int(self.окончательный_брак_прочее_input.text() or 0)

            # Расчет контроль_принято (убираем скрытые поля из расчета)
            контроль_принято = контроль_отлито - (
                второй_сорт_раковины + второй_сорт_зарез_литейный + второй_сорт_зарез_пеномодельный +
                # доработка_раковины + доработка_зарез +
                доработка_несоответствие_размеров + доработка_несоответствие_внешнего_вида +
                доработка_наплыв_металла + доработка_прорыв_металла +
                доработка_вырыв + доработка_облой +
                доработка_песок_на_поверхности + доработка_песок_в_резьбе +
                доработка_клей_подтёк + доработка_клей_по_шву +
                доработка_коробление +
                доработка_дефект_пеномодели + доработка_лапы +
                доработка_питатель + доработка_корона +
                доработка_смещение + окончательный_брак_недолив + окончательный_брак_раковины +
                окончательный_брак_коробление + окончательный_брак_спай +
                окончательный_брак_трещины + окончательный_брак_пригар_песка +
                окончательный_брак_пористость + окончательный_брак_вырыв +
                окончательный_брак_скол + окончательный_брак_слом +
                окончательный_брак_зарез_литейный + окончательный_брак_зарез_пеномодельный +
                окончательный_брак_нарушение_геометрии +
                окончательный_брак_рыхлота + окончательный_брак_непроклей +
                окончательный_брак_пеномодель + окончательный_брак_наплыв_металла +
                окончательный_брак_несоответствие_размеров + окончательный_брак_несоответствие_внешнего_вида +
                окончательный_брак_нарушение_маркировки + окончательный_брак_неслитина +
                окончательный_брак_прочее
            )
            self.контроль_принято_input.setText(str(контроль_принято))
        except ValueError:
            self.контроль_принято_input.setText("")

    def save_data(self):
        try:
            # Проверка обязательных полей
            if not self.номер_плавки_input.currentText():
                QMessageBox.warning(self, "Ошибка", "Выберите номер плавки")
                return
                
            if not self.контроль_отлито_input.text():
                QMessageBox.warning(self, "Ошибка", "Укажите количество отлитых деталей")
                return
                
            if not self.контролер1_input.currentText() and not self.контролер2_input.currentText() and not self.контролер3_input.currentText():
                QMessageBox.warning(self, "Ошибка", "Укажите хотя бы одного контролера")
                return

            # Словарь соответствия полей окончательного брака
            окончательный_брак_поля = {
                'Окончательный_брак_недолив': self.окончательный_брак_недолив_input.text(),
                'Окончательный_брак_вырыв': self.окончательный_брак_вырыв_input.text(),
                'Окончательный_брак_зарез_литейный': self.окончательный_брак_зарез_литейный_input.text(),
                'Окончательный_брак_зарез_пеномодельный': self.окончательный_брак_зарез_пеномодельный_input.text(),
                'Окончательный_брак_коробление': self.окончательный_брак_коробление_input.text(),
                'Окончательный_брак_наплыв_металла': self.окончательный_брак_наплыв_металла_input.text(),
                'Окончательный_брак_нарушение_геометрии': self.окончательный_брак_нарушение_геометрии_input.text(),
                'Окончательный_брак_нарушение_маркировки': self.окончательный_брак_нарушение_маркировки_input.text(),
                'Окончательный_брак_непроклей': self.окончательный_брак_непроклей_input.text(),
                'Окончательный_брак_неслитина': self.окончательный_брак_неслитина_input.text(),
                'Окончательный_брак_несоответствие_внешнего_вида': self.окончательный_брак_несоответствие_внешнего_вида_input.text(),
                'Окончательный_брак_несоответствие_размеров': self.окончательный_брак_несоответствие_размеров_input.text(),
                'Окончательный_брак_пеномодель': self.окончательный_брак_пеномодель_input.text(),
                'Окончательный_брак_пористость': self.окончательный_брак_пористость_input.text(),
                'Окончательный_брак_пригар_песка': self.окончательный_брак_пригар_песка_input.text(),
                'Окончательный_брак_прочее': self.окончательный_брак_прочее_input.text(),
                'Окончательный_брак_рыхлота': self.окончательный_брак_рыхлота_input.text(),
                'Окончательный_брак_раковины': self.окончательный_брак_раковины_input.text(),
                'Окончательный_брак_скол': self.окончательный_брак_скол_input.text(),
                'Окончательный_брак_слом': self.окончательный_брак_слом_input.text(),
                'Окончательный_брак_спай': self.окончательный_брак_спай_input.text(),
                'Окончательный_брак_трещины': self.окончательный_брак_трещины_input.text(),
            }

            # Заголовки для базовых данных
            headers = [
                'Номер_плавки', 'Контроль_отлито', 'Контроль_принято',
                'Контроль_дата_приемки', 'Контролер1', 'Контролер2', 'Контролер3',
                'Второй_сорт_раковины', 'Второй_сорт_зарез_литейный', 'Второй_сорт_зарез_пеномодельный',
                'Доработка_несоответствие_размеров', 'Доработка_несоответствие_внешнего_вида',
                'Доработка_наплыв_металла', 'Доработка_прорыв_металла',
                'Доработка_вырыв', 'Доработка_облой',
                'Доработка_песок_на_поверхности', 'Доработка_песок_в_резьбе',
                'Доработка_клей_подтёк', 'Доработка_клей_по_шву',
                'Доработка_коробление',
                'Доработка_дефект_пеномодели', 'Доработка_лапы',
                'Доработка_питатель', 'Доработка_корона',
                'Доработка_смещение',
            ]

            # Собираем базовые данные в список
            data = [
                self.номер_плавки_input.currentText(),
                self.контроль_отлито_input.text(),
                self.контроль_принято_input.text(),
                self.контроль_дата_приемки_input.date().toString("dd.MM.yyyy"),
                self.контролер1_input.currentText(),
                self.контролер2_input.currentText(),
                self.контролер3_input.currentText(),
                self.второй_сорт_раковины_input.text(),
                self.второй_сорт_зарез_литейный_input.text(),
                self.второй_сорт_зарез_пеномодельный_input.text(),
                self.доработка_несоответствие_размеров_input.text(),
                self.доработка_несоответствие_внешнего_вида_input.text(),
                self.доработка_наплыв_металла_input.text(),
                self.доработка_прорыв_металла_input.text(),
                self.доработка_вырыв_input.text(),
                self.доработка_облой_input.text(),
                self.доработка_песок_на_поверхности_input.text(),
                self.доработка_песок_в_резьбе_input.text(),
                self.доработка_клей_подтёк_input.text(),
                self.доработка_клей_по_шву_input.text(),
                self.доработка_коробление_input.text(),
                self.доработка_дефект_пеномодели_input.text(),
                self.доработка_лапы_input.text(),
                self.доработка_питатель_input.text(),
                self.доработка_корона_input.text(),
                self.доработка_смещение_input.text(),
            ]

            # Добавляем заголовки окончательного брака в нужном порядке
            headers.extend([
                'Окончательный_брак_недолив', 'Окончательный_брак_вырыв',
                'Окончательный_брак_зарез_литейный', 'Окончательный_брак_зарез_пеномодельный', 
                'Окончательный_брак_коробление',
                'Окончательный_брак_наплыв_металла', 'Окончательный_брак_нарушение_геометрии',
                'Окончательный_брак_нарушение_маркировки', 'Окончательный_брак_непроклей',
                'Окончательный_брак_неслитина', 'Окончательный_брак_несоответствие_внешнего_вида',
                'Окончательный_брак_несоответствие_размеров', 'Окончательный_брак_пеномодель',
                'Окончательный_брак_пористость', 'Окончательный_брак_пригар_песка',
                'Окончательный_брак_прочее', 'Окончательный_брак_рыхлота',
                'Окончательный_брак_раковины', 'Окончательный_брак_скол',
                'Окончательный_брак_слом', 'Окончательный_брак_спай',
                'Окончательный_брак_трещины'
            ])
            
            # Добавляем Наименование_отливки в конец заголовков
            headers.append('Наименование_отливки')

            # Добавляем данные окончательного брака в том же порядке, что и заголовки
            for header in headers[26:48]:  # Используем правильный диапазон для полей окончательного брака
                data.append(окончательный_брак_поля[header])
                
            # Добавляем значение Наименование_отливки в конец списка данных
            data.append(self.наименование_отливки_input.text())

            if os.path.exists('control.xlsx'):
                # Проверяем структуру существующей таблицы
                df_existing = pd.read_excel('control.xlsx')
                existing_columns = df_existing.columns.tolist()
                
                # Проверяем, есть ли дополнительные столбцы в headers, которых нет в existing_columns
                missing_columns = [col for col in headers if col not in existing_columns]
                
                # Загружаем файл для работы
                wb = load_workbook('control.xlsx')
                ws = wb.active
                
                # Добавляем недостающие заголовки, если они есть
                if missing_columns:
                    for i, col_name in enumerate(missing_columns):
                        col_idx = len(existing_columns) + 1 + i
                        ws.cell(row=1, column=col_idx, value=col_name)
            else:
                wb = Workbook()
                ws = wb.active
                # Добавляем заголовки только если это новый файл
                for col, header in enumerate(headers, start=1):
                    ws.cell(row=1, column=col, value=header)

            # Добавляем новую строку данных
            next_row = ws.max_row + 1
            
            # Если файл существует и есть данные
            if os.path.exists('control.xlsx') and next_row > 2:
                # Мэппим данные к существующим заголовкам
                existing_columns = [ws.cell(row=1, column=col).value for col in range(1, ws.max_column + 1)]
                
                for col, header in enumerate(existing_columns, start=1):
                    if header in headers:
                        # Находим индекс этого заголовка в нашем массиве
                        data_idx = headers.index(header)
                        if data_idx < len(data):
                            cell = ws.cell(row=next_row, column=col)
                            cell.value = data[data_idx]
                            if header == 'Контроль_дата_приемки':  # Форматирование даты
                                cell.number_format = 'DD.MM.YYYY'
            else:
                # Для нового файла просто записываем последовательно
                for col, value in enumerate(data, start=1):
                    cell = ws.cell(row=next_row, column=col)
                    cell.value = value
                    if col == 4:  # Колонка D (дата)
                        cell.number_format = 'DD.MM.YYYY'

            # Применяем формат даты ко всем ячейкам в колонке даты
            date_col = None
            for col in range(1, ws.max_column + 1):
                if ws.cell(row=1, column=col).value == 'Контроль_дата_приемки':
                    date_col = col
                    break
                    
            if date_col:
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row, column=date_col)
                    cell.number_format = 'DD.MM.YYYY'

            wb.save('control.xlsx')
            wb.close()

            QMessageBox.information(self, "Успех", "Данные успешно сохранены!")
            
            # Очищаем и обновляем список доступных номеров плавок
            self.номер_плавки_input.clear()
            self.load_plavka_numbers()  # Обновляем список доступных номеров
            
            # Очищаем форму
            self.clear_form()
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении данных: {str(e)}")
        
        
    def clear_form(self):
        """Очистка формы"""
        # Очищаем все поля без дополнительных диалогов
        self.номер_плавки_input.setCurrentIndex(-1)
        self.контроль_отлито_input.setText('')
        self.контроль_принято_input.setText('')
        self.контроль_дата_приемки_input.setDate(QDate.currentDate())
        self.контролер1_input.setCurrentIndex(-1)
        self.контролер2_input.setCurrentIndex(-1)
        self.контролер3_input.setCurrentIndex(-1)

        # Очистка полей второго сорта
        self.второй_сорт_раковины_input.setText('')
        self.второй_сорт_зарез_литейный_input.setText('')
        self.второй_сорт_зарез_пеномодельный_input.setText('')

        # Очистка полей доработки
        self.доработка_раковины_input.setText('')
        self.доработка_зарез_input.setText('')
        self.доработка_несоответствие_размеров_input.setText('')
        self.доработка_несоответствие_внешнего_вида_input.setText('')
        self.доработка_наплыв_металла_input.setText('')
        self.доработка_прорыв_металла_input.setText('')
        self.доработка_вырыв_input.setText('')
        self.доработка_облой_input.setText('')
        self.доработка_песок_на_поверхности_input.setText('')
        self.доработка_песок_в_резьбе_input.setText('')
        self.доработка_клей_подтёк_input.setText('')
        self.доработка_клей_по_шву_input.setText('')
        self.доработка_коробление_input.setText('')
        self.доработка_дефект_пеномодели_input.setText('')
        self.доработка_лапы_input.setText('')
        self.доработка_питатель_input.setText('')
        self.доработка_корона_input.setText('')
        self.доработка_смещение_input.setText('')

        # Очистка полей окончательного брака
        self.окончательный_брак_недолив_input.setText('')
        self.окончательный_брак_раковины_input.setText('')
        self.окончательный_брак_коробление_input.setText('')
        self.окончательный_брак_спай_input.setText('')
        self.окончательный_брак_трещины_input.setText('')
        self.окончательный_брак_пригар_песка_input.setText('')
        self.окончательный_брак_пористость_input.setText('')
        self.окончательный_брак_вырыв_input.setText('')
        self.окончательный_брак_скол_input.setText('')
        self.окончательный_брак_слом_input.setText('')
        self.окончательный_брак_зарез_литейный_input.setText('')
        self.окончательный_брак_зарез_пеномодельный_input.setText('')
        self.окончательный_брак_нарушение_геометрии_input.setText('')
        self.окончательный_брак_рыхлота_input.setText('')
        self.окончательный_брак_непроклей_input.setText('')
        self.окончательный_брак_пеномодель_input.setText('')
        self.окончательный_брак_наплыв_металла_input.setText('')
        self.окончательный_брак_несоответствие_размеров_input.setText('')
        self.окончательный_брак_несоответствие_внешнего_вида_input.setText('')
        self.окончательный_брак_нарушение_маркировки_input.setText('')
        self.окончательный_брак_неслитина_input.setText('')
        self.окончательный_брак_прочее_input.setText('')

        # Обновляем поля с пустым значением
        self.наименование_отливки_input.setText('')
        self.номер_кластера_input.setText('')

    def animate_group_hover(self, group, hover_in):
        if not hasattr(self, 'animations'):
            self.animations = []
        
        shadow = group.graphicsEffect()
        if shadow and shadow.isEnabled():
            # Удаляем завершенные анимации
            self.animations = [a for a in self.animations if a.state() != QPropertyAnimation.Stopped]
            
            animation = QPropertyAnimation(shadow, b"blurRadius")
            animation.setDuration(200)
            animation.setEasingCurve(QEasingCurve.InOutCubic)
            
            if hover_in:
                animation.setStartValue(15)
                animation.setEndValue(25)
            else:
                animation.setStartValue(25)
                animation.setEndValue(15)
            
            # Сохраняем анимацию
            self.animations.append(animation)
            animation.start()

    def show_report_dialog(self):
        """Открывает диалог для формирования отчетов"""
        try:
            dialog = ReportDialog(self)
            if dialog.exec():
                report_type, selected_date = dialog.get_report_data()
                
                # Создаем объект генератора отчетов
                report_generator = ReportGenerator()
                
                try:
                    QMessageBox.information(self, "Информация", 
                                         f"Формирование отчета - {dialog.get_report_type_name(report_type)} за {selected_date}. "
                                         f"Это может занять некоторое время.")
                    
                    if report_type == 'daily':
                        # Сначала генерируем сводный отчет
                        summary_report_file = report_generator.generate_summary_report(selected_date)
                        # Затем создаем ежедневный отчет на его основе
                        daily_report_file = report_generator.generate_daily_report(selected_date, summary_report_file)
                        QMessageBox.information(self, "Успех", f"Отчет сохранен в файл: {daily_report_file}")
                    elif report_type == 'summary':
                        # Генерируем обычный сводный отчет
                        summary_report_file = report_generator.generate_summary_report(selected_date)
                        QMessageBox.information(self, "Успех", f"Отчет сохранен в файл: {summary_report_file}")
                    elif report_type == 'full':
                        # Генерируем полный сводный отчет
                        full_report_file = report_generator.generate_full_report(selected_date)
                        QMessageBox.information(self, "Успех", f"Отчет сохранен в файл: {full_report_file}")
                    elif report_type == 'template':
                        # Создаем новый шаблон отчета
                        template_file = report_generator.create_new_report_template()
                        QMessageBox.information(self, "Успех", f"Шаблон отчета создан в файл: {template_file}")
                    
                except Exception as e:
                    traceback_str = traceback.format_exc()
                    QMessageBox.critical(self, "Ошибка", f"Не удалось сформировать отчет: {str(e)}")
                    with open('report_generator.log', 'a', encoding='utf-8') as f:
                        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - ОШИБКА: {str(e)}\n{traceback_str}\n")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при открытии диалога: {str(e)}")
    
    def debug_report_file(self, file_path):
        """Отладочная функция для анализа структуры отчета"""
        try:
            # Создаем диалог для отображения информации
            debug_dialog = QDialog(self)
            debug_dialog.setWindowTitle(f"Отладка отчета: {os.path.basename(file_path)}")
            debug_dialog.setGeometry(100, 100, 800, 600)
            
            layout = QVBoxLayout(debug_dialog)
            
            # Текстовое поле для отображения информации
            text_browser = QTextBrowser()
            text_browser.setStyleSheet("""
                background-color: #1a1b26;
                color: #f8f8f2;
                font-family: 'Consolas', 'Courier New';
                font-size: 11px;
            """)
            
            layout.addWidget(text_browser)
            
            # Добавляем информацию о размере файла
            file_size = os.path.getsize(file_path)
            text_browser.append(f"Файл: {file_path}")
            text_browser.append(f"Размер: {file_size} байт")
            text_browser.append("-" * 50)
            
            # Анализируем структуру файла Excel
            wb = load_workbook(file_path)
            ws = wb.active
            
            text_browser.append(f"Имя листа: {ws.title}")
            text_browser.append(f"Размеры: {ws.max_row} строк, {ws.max_column} столбцов")
            text_browser.append("-" * 50)
            
            # Ищем заголовки таблицы
            header_rows = []
            for row in range(1, min(30, ws.max_row + 1)):
                has_headers = False
                for col in range(1, min(10, ws.max_column + 1)):
                    cell_value = ws.cell(row=row, column=col).value
                    if cell_value and isinstance(cell_value, str) and any(keyword in cell_value.lower() for keyword in ["дата", "плавк", "отливк", "брак", "сорт"]):
                        has_headers = True
                        break
                if has_headers:
                    header_rows.append(row)
            
            if header_rows:
                text_browser.append(f"Строки с заголовками: {header_rows}")
            else:
                text_browser.append("ВНИМАНИЕ: Не найдены строки с заголовками!")
            
            text_browser.append("-" * 50)
            
            # Проверяем данные в таблице
            data_found = False
            for row in range(1, ws.max_row + 1):
                row_has_data = False
                for col in range(1, ws.max_column + 1):
                    cell_value = ws.cell(row=row, column=col).value
                    if cell_value is not None and str(cell_value).strip():
                        row_has_data = True
                        break
                
                if row_has_data:
                    data_found = True
                    # Выводим первые несколько ячеек строки
                    data_preview = []
                    for col in range(1, min(5, ws.max_column + 1)):
                        cell_value = ws.cell(row=row, column=col).value
                        if cell_value is not None:
                            data_preview.append(str(cell_value))
                    
                    text_browser.append(f"Строка {row}: {', '.join(data_preview)}...")
                    
                    # Ограничиваем вывод первыми 20 строками с данными
                    if len(text_browser.toPlainText().split("\n")) > 40:
                        text_browser.append("...")
                        break
            
            if not data_found:
                text_browser.append("ВНИМАНИЕ: Данные в таблице не найдены!")
            
            # Кнопка "Закрыть"
            close_button = QPushButton("Закрыть")
            close_button.clicked.connect(debug_dialog.accept)
            layout.addWidget(close_button)
            
            # Отображаем диалог
            debug_dialog.exec()
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка отладки", f"Ошибка при анализе файла: {str(e)}")

# Класс диалогового окна для формирования отчетов
class ReportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Формирование отчета")
        self.setGeometry(100, 100, 400, 400)
        
        # Включаем режим отладки
        self.debug_mode = True
        
        # Задаем темную тему как в основном приложении
        self.setStyleSheet("""
            QDialog {
                background-color: #282a36;
                color: #f8f8f2;
                font-family: 'Segoe UI', 'Aptos';
                font-size: 11px;
            }
            
            QLabel {
                color: #f8f8f2;
                padding: 2px;
            }
            
            QRadioButton {
                color: #f8f8f2;
                padding: 5px;
            }
            
            QRadioButton::indicator {
                width: 13px;
                height: 13px;
            }
            
            QRadioButton::indicator:checked {
                background-color: #bd93f9;
                border: 2px solid #f8f8f2;
                border-radius: 6px;
            }
            
            QPushButton {
                background-color: #6272a4;
                color: #f8f8f2;
                border: none;
                padding: 8px 15px;
                border-radius: 2px;
                height: 30px;
            }
            
            QPushButton:hover {
                background-color: #bd93f9;
            }
            
            QPushButton:pressed {
                background-color: #ff79c6;
            }
            
            QCalendarWidget {
                background-color: #282a36;
                color: #f8f8f2;
            }
            
            QCalendarWidget QToolButton {
                color: #f8f8f2;
                background-color: #44475a;
                border: none;
            }
            
            QCalendarWidget QMenu {
                color: #f8f8f2;
                background-color: #44475a;
            }
            
            QCalendarWidget QSpinBox {
                color: #f8f8f2;
                background-color: #44475a;
                selection-background-color: #6272a4;
                selection-color: #f8f8f2;
            }
            
            QCalendarWidget QAbstractItemView:enabled {
                color: #f8f8f2;
                background-color: #44475a;
                selection-background-color: #6272a4;
                selection-color: #f8f8f2;
            }
            
            QCalendarWidget QWidget {
                alternate-background-color: #44475a;
            }
            
            QGroupBox {
                border: 1px solid #44475a;
                border-radius: 4px;
                margin-top: 0.5em;
                padding: 8px;
            }
            
            QGroupBox::title {
                color: #bd93f9;
            }
            
            QTextBrowser {
                background-color: #1a1b26;
                color: #f8f8f2;
                border: 1px solid #6272a4;
                font-family: 'Consolas', 'Courier New';
                font-size: 11px;
            }
        """)
        
        layout = QVBoxLayout(self)
        
        # Группа выбора типа отчета
        report_group = QGroupBox("Тип отчета")
        report_layout = QVBoxLayout()
        
        self.radio_daily = QRadioButton("Ежедневный отчет")
        self.radio_summary = QRadioButton("Стандартный сводный отчет")
        self.radio_full = QRadioButton("Полный сводный отчет")
        self.radio_template = QRadioButton("Создать шаблон полного отчета")
        
        self.radio_daily.setChecked(True)
        
        report_layout.addWidget(self.radio_daily)
        report_layout.addWidget(self.radio_summary)
        report_layout.addWidget(self.radio_full)
        report_layout.addWidget(self.radio_template)
        
        report_group.setLayout(report_layout)
        layout.addWidget(report_group)
        
        # Группа выбора даты
        self.date_group = QGroupBox("Выбор даты")
        date_layout = QVBoxLayout()
        
        self.calendar = QCalendarWidget()
        self.calendar.setSelectedDate(QDate.currentDate())
        self.calendar.setGridVisible(True)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        
        # Соединяем переключатели типа отчета с управлением видимостью блока даты
        self.radio_template.toggled.connect(self.toggle_date_visibility)
        
        date_layout.addWidget(self.calendar)
        self.date_group.setLayout(date_layout)
        layout.addWidget(self.date_group)
        
        # Добавляем справочную информацию
        info_group = QGroupBox("Информация о типах отчетов")
        info_layout = QVBoxLayout()
        
        info_text = QLabel(
            "• Ежедневный отчет - простой отчет с выделением\n"
            "   наименований отливок по дате\n"
            "• Стандартный сводный отчет - сводная таблица\n"
            "   статистики по отливкам с основными показателями\n"
            "• Полный сводный отчет - детальная статистика по\n"
            "   всем типам дефектов, включая редкие варианты\n"
            "• Шаблон полного отчета - создает пустой шаблон\n"
            "   для последующего заполнения"
        )
        info_text.setStyleSheet("color: #8be9fd; padding: 5px;")
        
        info_layout.addWidget(info_text)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # Добавляем лог для отладки
        if self.debug_mode:
            debug_group = QGroupBox("Отладочная информация")
            debug_layout = QVBoxLayout()
            
            self.log_browser = QTextBrowser()
            self.log_browser.setMaximumHeight(100)
            
            debug_layout.addWidget(self.log_browser)
            debug_group.setLayout(debug_layout)
            layout.addWidget(debug_group)
        
        # Кнопки
        buttons_layout = QHBoxLayout()
        
        self.generate_button = QPushButton("Сформировать")
        self.generate_button.setStyleSheet("""
            QPushButton {
                background-color: #50fa7b;
                color: #282a36;
                font-size: 14px;
                padding: 12px 30px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #69ff94;
            }
            QPushButton:pressed {
                background-color: #41d66b;
            }
        """)
        self.generate_button.clicked.connect(self.validate_and_accept)
        
        self.cancel_button = QPushButton("Отмена")
        self.cancel_button.clicked.connect(self.reject)
        
        buttons_layout.addWidget(self.cancel_button)
        buttons_layout.addWidget(self.generate_button)
        
        layout.addLayout(buttons_layout)
    
    def toggle_date_visibility(self, checked):
        """Переключает видимость выбора даты при выборе создания шаблона"""
        self.date_group.setVisible(not checked)
    
    def add_log(self, message):
        """Добавляет сообщение в отладочный лог"""
        if self.debug_mode and hasattr(self, 'log_browser'):
            self.log_browser.append(f"{datetime.now().strftime('%H:%M:%S')} - {message}")
            self.log_browser.verticalScrollBar().setValue(
                self.log_browser.verticalScrollBar().maximum()
            )
            QApplication.processEvents()  # Обновляем UI
    
    def validate_and_accept(self):
        """Проверяет наличие необходимых файлов шаблонов перед формированием отчетов"""
        try:
            # Получаем тип отчета и дату
            report_type = self.get_report_type()
            selected_date = self.calendar.selectedDate().toString("dd.MM.yyyy")
            
            # Проверяем наличие необходимых файлов
            if report_type == 'daily':
                if not os.path.exists('Отчёт.xlsx'):
                    self.add_log(f"ОШИБКА: Шаблон Отчёт.xlsx не найден!")
                    QMessageBox.critical(self, "Ошибка", f"Шаблон Отчёт.xlsx не найден!")
                    return
                    
                if not os.path.exists('Сводный.xlsx'):
                    self.add_log(f"ОШИБКА: Шаблон Сводный.xlsx не найден!")
                    QMessageBox.critical(self, "Ошибка", f"Шаблон Сводный.xlsx не найден!")
                    return
            
            if report_type == 'summary' and not os.path.exists('Сводный.xlsx'):
                self.add_log(f"ОШИБКА: Шаблон Сводный.xlsx не найден!")
                QMessageBox.critical(self, "Ошибка", f"Шаблон Сводный.xlsx не найден!")
                return
            
            if not os.path.exists('control.xlsx'):
                self.add_log("ОШИБКА: Файл control.xlsx не найден!")
                QMessageBox.critical(self, "Ошибка", "Файл control.xlsx не найден!")
                return
            
            self.add_log(f"Выбран тип отчета: {report_type}, дата: {selected_date}")
            
            # Принимаем диалог
            self.accept()
            
        except Exception as e:
            self.add_log(f"ОШИБКА: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при проверке: {str(e)}")
    
    def analyze_template(self, template_file):
        """Анализирует структуру шаблона отчета для отладки"""
        try:
            self.add_log(f"Анализ шаблона {template_file}...")
            
            # Загружаем шаблон
            wb = load_workbook(template_file)
            ws = wb.active
            
            # Проверяем размер файла
            self.add_log(f"Размеры листа: {ws.max_row} строк, {ws.max_column} столбцов")
            
            # Ищем ключевые заголовки
            header_rows = []
            for row in range(1, min(20, ws.max_row + 1)):
                has_headers = False
                for col in range(1, min(10, ws.max_column + 1)):
                    cell_value = ws.cell(row=row, column=col).value
                    if cell_value and isinstance(cell_value, str) and any(keyword in cell_value.lower() for keyword in ["дата", "плавк", "отливк", "брак", "сорт"]):
                        has_headers = True
                        break
                if has_headers:
                    header_rows.append(row)
            
            if header_rows:
                self.add_log(f"Найдены строки с заголовками: {header_rows}")
            else:
                self.add_log("ВНИМАНИЕ: Не найдены строки с заголовками!")
            
        except Exception as e:
            self.add_log(f"Ошибка при анализе шаблона: {str(e)}")
    
    def get_report_type(self):
        """Возвращает тип отчета"""
        if self.radio_daily.isChecked():
            return "daily"
        elif self.radio_summary.isChecked():
            return "summary"
        elif self.radio_full.isChecked():
            return "full"
        elif self.radio_template.isChecked():
            return "template"
        return "summary"  # По умолчанию
    
    def get_report_type_name(self, report_type):
        """Возвращает название типа отчета для отображения"""
        types = {
            "daily": "Ежедневный отчет",
            "summary": "Стандартный сводный отчет",
            "full": "Полный сводный отчет",
            "template": "Шаблон полного отчета"
        }
        return types.get(report_type, "Отчет")
    
    def get_report_data(self):
        """Возвращает тип отчета и выбранную дату"""
        report_type = self.get_report_type()
        selected_date = self.calendar.selectedDate().toString("dd.MM.yyyy")
        return report_type, selected_date

# Класс для работы с отчетами
class ReportGenerator:
    def __init__(self):
        self.control_file = 'control.xlsx'
        self.report_file = 'Отчёт.xlsx'
        self.summary_file = 'Сводный.xlsx'
        self.debug = True
        
        # Добавляем импорт traceback для отслеживания ошибок, если его еще нет
        try:
            import traceback
        except ImportError:
            pass  # traceback уже импортирован
            
    def log(self, message):
        """Выводит отладочные сообщения, если включен режим отладки"""
        if self.debug:
            print(f"[DEBUG] {message}")
            with open('report_generator.log', 'a', encoding='utf-8') as f:
                f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
    
    def generate_daily_report(self, date_str, summary_report_file=None):
        """Генерирует ежедневный отчет на основе даты и данных из сводного отчета"""
        if not os.path.exists(self.report_file):
            raise FileNotFoundError(f"Файл шаблона {self.report_file} не найден")
        
        # Проверяем наличие сводного отчета
        if summary_report_file is None or not os.path.exists(summary_report_file):
            # Если сводный отчет не указан, пытаемся найти его по дате
            summary_report_file = f'Сводный_{date_str.replace(".", "-")}.xlsx'
            if not os.path.exists(summary_report_file):
                # Если сводный отчет не найден, генерируем его
                summary_report_file = self.generate_summary_report(date_str)
        
        try:
            self.log(f"Формирование ежедневного отчета на основе сводного отчета: {summary_report_file}")
            
            # Загружаем данные из сводного отчета через pandas
            df_summary = pd.read_excel(summary_report_file)
            
            # Если данных нет, сообщаем об ошибке
            if df_summary.empty:
                raise ValueError(f"В сводном отчете {summary_report_file} нет данных")
            
            self.log(f"Колонки в сводном отчете: {df_summary.columns.tolist()}")
            
            # Создаем копию шаблона отчета для сохранения результатов
            report_output = f'Отчет_{date_str.replace(".", "-")}.xlsx'
            
            # Копируем оригинальный файл шаблона для сохранения форматирования
            shutil.copy(self.report_file, report_output)
            
            # Загружаем файл отчета
            wb = load_workbook(report_output)
            ws = wb.active
            
            # Пишем дату в файл (ищем ячейку с "Дата" для заполнения)
            for row in range(1, 10):
                for col in range(1, 5):
                    cell_value = str(ws.cell(row=row, column=col).value or "").lower()
                    if "дата" in cell_value:
                        # Проверяем, не является ли целевая ячейка объединенной
                        target_cell = ws.cell(row=row, column=col+1)
                        if hasattr(target_cell, 'coordinate'):
                            # Это не объединенная ячейка, можем напрямую изменять значение
                            target_cell.value = date_str
                        else:
                            # Это объединенная ячейка, найдем основную ячейку и изменим её
                            # Ищем все объединенные диапазоны
                            for merged_range in ws.merged_cells.ranges:
                                if target_cell.coordinate in merged_range:
                                    # Нашли диапазон, содержащий текущую ячейку
                                    # Получаем координату верхней левой ячейки диапазона
                                    main_cell_coord = merged_range.coord.split(':')[0]
                                    # Записываем значение в основную ячейку диапазона
                                    ws[main_cell_coord] = date_str
                                    break
                        break
            
            # Находим основные блоки данных
            # Обычно это строки 7 и 18 для заголовков двух таблиц
            header_row1 = None
            header_row2 = None
            
            for row in range(1, 25):
                cell_value = str(ws.cell(row=row, column=1).value or "").lower()
                if "наименование" in cell_value and "отливк" in cell_value:
                    if header_row1 is None:
                        header_row1 = row
                    elif header_row2 is None:
                        header_row2 = row
            
            if header_row1 is None:
                header_row1 = 7  # Значение по умолчанию
            if header_row2 is None:
                header_row2 = 18  # Значение по умолчанию
                
            self.log(f"Найдены строки заголовков: {header_row1} и {header_row2}")
            
            # Загружаем данные из сводного отчета
            summary_data = []
            try:
                for index, row in df_summary.iterrows():
                    # Пропускаем пустые строки или строки ИТОГО
                    if pd.isna(row[0]) or "ИТОГО" in str(row[0]).upper():
                        continue
                    
                    summary_data.append(row.to_dict())
            except Exception as e:
                self.log(f"Ошибка при чтении сводного отчета: {str(e)}")
                
                # Попробуем загрузить через openpyxl, если pandas не сработал
                wb_summary = load_workbook(summary_report_file)
                ws_summary = wb_summary.active
                
                # Ищем строку заголовков
                header_row_summary = 3  # По умолчанию для нового формата
                for row in range(1, 10):
                    cell_value = str(ws_summary.cell(row=row, column=1).value or "").lower()
                    if "наименование" in cell_value and "отливк" in cell_value:
                        header_row_summary = row
                        break
                
                # Получаем заголовки
                headers = []
                for col in range(1, ws_summary.max_column + 1):
                    headers.append(str(ws_summary.cell(row=header_row_summary, column=col).value or ""))
                
                # Собираем данные
                for row in range(header_row_summary + 1, ws_summary.max_row + 1):
                    if ws_summary.cell(row=row, column=1).value is None:
                        continue
                    
                    if "ИТОГО" in str(ws_summary.cell(row=row, column=1).value).upper():
                        continue
                    
                    row_data = {}
                    for col in range(1, len(headers) + 1):
                        if col <= ws_summary.max_column:
                            row_data[headers[col-1]] = ws_summary.cell(row=row, column=col).value
                    
                    if row_data:
                        summary_data.append(row_data)
                
                wb_summary.close()
            
            self.log(f"Прочитано {len(summary_data)} записей из сводного отчета")
            
            # Заполняем первую таблицу
            data_row1 = header_row1 + 1
            for i, data in enumerate(summary_data[:10]):  # Первые 10 записей
                # Определяем наименование отливки и номера плавок
                name = None
                for key in data:
                    if "наименование" in str(key).lower() and "отливк" in str(key).lower():
                        name = data[key]
                        break
                
                if not name:
                    continue
                
                # Ищем столбцы для заполнения в первой таблице
                for col in range(1, ws.max_column + 1):
                    header = str(ws.cell(row=header_row1, column=col).value or "").lower()
                    
                    if "наименование" in header and "отливк" in header:
                        ws.cell(row=data_row1 + i, column=col).value = name
                    elif "номер" in header and "плавк" in header:
                        # Ищем номера плавок в данных
                        for key in data:
                            if "номер" in str(key).lower() and "плавк" in str(key).lower():
                                ws.cell(row=data_row1 + i, column=col).value = data[key]
                                break
                    elif "отлито" in header and not any(x in header for x in ["брак", "сорт", "доработк"]):
                        # Ищем отлито в данных
                        for key in data:
                            if "отлито" in str(key).lower() and not any(x in str(key).lower() for x in ["брак", "сорт", "доработк"]):
                                ws.cell(row=data_row1 + i, column=col).value = data[key]
                                break
                    elif "принято" in header and not any(x in header for x in ["брак", "сорт", "доработк"]):
                        # Ищем принято в данных
                        for key in data:
                            if "принято" in str(key).lower() and not any(x in str(key).lower() for x in ["брак", "сорт", "доработк"]):
                                ws.cell(row=data_row1 + i, column=col).value = data[key]
                                break
                    elif "годн" in header:
                        # Ищем процент годности в данных
                        for key in data:
                            if "годн" in str(key).lower():
                                cell = ws.cell(row=data_row1 + i, column=col)
                                cell.value = data[key]
                                cell.number_format = '0.00"%"'
                                break
            
            # Заполняем вторую таблицу
            data_row2 = header_row2 + 1
            for i, data in enumerate(summary_data[10:20]):  # Следующие 10 записей
                # Определяем наименование отливки и номера плавок
                name = None
                for key in data:
                    if "наименование" in str(key).lower() and "отливк" in str(key).lower():
                        name = data[key]
                        break
                
                if not name:
                    continue
                
                # Ищем столбцы для заполнения во второй таблице
                for col in range(1, ws.max_column + 1):
                    header = str(ws.cell(row=header_row2, column=col).value or "").lower()
                    
                    if "наименование" in header and "отливк" in header:
                        ws.cell(row=data_row2 + i, column=col).value = name
                    elif "номер" in header and "плавк" in header:
                        # Ищем номера плавок в данных
                        for key in data:
                            if "номер" in str(key).lower() and "плавк" in str(key).lower():
                                ws.cell(row=data_row2 + i, column=col).value = data[key]
                                break
                    elif "отлито" in header and not any(x in header for x in ["брак", "сорт", "доработк"]):
                        # Ищем отлито в данных
                        for key in data:
                            if "отлито" in str(key).lower() and not any(x in str(key).lower() for x in ["брак", "сорт", "доработк"]):
                                ws.cell(row=data_row2 + i, column=col).value = data[key]
                                break
                    elif "принято" in header and not any(x in header for x in ["брак", "сорт", "доработк"]):
                        # Ищем принято в данных
                        for key in data:
                            if "принято" in str(key).lower() and not any(x in str(key).lower() for x in ["брак", "сорт", "доработк"]):
                                ws.cell(row=data_row2 + i, column=col).value = data[key]
                                break
                    elif "годн" in header:
                        # Ищем процент годности в данных
                        for key in data:
                            if "годн" in str(key).lower():
                                cell = ws.cell(row=data_row2 + i, column=col)
                                cell.value = data[key]
                                cell.number_format = '0.00"%"'
                                break
            
            # Устанавливаем форматирование для всех заполненных ячеек
            for row_start, row_end in [(data_row1, data_row1 + 10), (data_row2, data_row2 + 10)]:
                for row in range(row_start, row_end):
                    for col in range(1, ws.max_column + 1):
                        cell = ws.cell(row=row, column=col)
                        if cell.value is not None:
                            # Устанавливаем выравнивание в зависимости от типа данных
                            if isinstance(cell.value, str) and len(cell.value) > 10:
                                cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                            else:
                                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Сохраняем файл
            wb.save(report_output)
            wb.close()
            self.log(f"Ежедневный отчет успешно сохранен в файл: {report_output}")
            
            return report_output
        
        except Exception as e:
            print(f"ОШИБКА при формировании ежедневного отчета: {str(e)}")
            traceback_str = traceback.format_exc()
            print(traceback_str)
            self.log(f"ОШИБКА при формировании ежедневного отчета: {str(e)}")
            self.log(traceback_str)
            raise Exception(f"Ошибка при формировании ежедневного отчета: {str(e)}")
    
    def generate_summary_report(self, date_str):
        """Генерирует сводный отчет по наименованиям отливок"""
        if not os.path.exists(self.control_file):
            raise FileNotFoundError(f"Файл {self.control_file} не найден")
        
        try:
            import traceback
            self.log(f"Начало формирования сводного отчета за {date_str}")
            
            # Загружаем данные из control.xlsx
            self.log(f"Загрузка данных из {self.control_file}")
            df_control = pd.read_excel(self.control_file)
            
            # Заменяем NaN значения на 0 для всех числовых колонок
            numeric_cols = df_control.select_dtypes(include=['float64', 'int64']).columns
            df_control[numeric_cols] = df_control[numeric_cols].fillna(0)
            
            # Выводим колонки для отладки
            self.log(f"Колонки в control.xlsx: {df_control.columns.tolist()}")
            self.log(f"Всего записей в control.xlsx: {len(df_control)}")
            
            # Конвертируем даты
            self.log(f"Конвертация даты приемки")
            if 'Контроль_дата_приемки' in df_control.columns:
                df_control['Контроль_дата_приемки'] = pd.to_datetime(df_control['Контроль_дата_приемки'], format='%d.%m.%Y', errors='coerce')
            else:
                self.log(f"ВНИМАНИЕ: Колонка 'Контроль_дата_приемки' не найдена")
                df_control['Контроль_дата_приемки'] = pd.NaT
            
            # Если указана дата, фильтруем по ней
            if date_str:
                self.log(f"Фильтрация по дате {date_str}")
                selected_date = datetime.strptime(date_str, '%d.%m.%Y')
                df_filtered = df_control[df_control['Контроль_дата_приемки'].dt.date == selected_date.date()]
                
                self.log(f"После фильтрации осталось записей: {len(df_filtered)}")
                if df_filtered.empty:
                    raise ValueError(f"Нет данных за {date_str}")
            else:
                df_filtered = df_control
            
            # Создаем новый сводный отчет вместо использования шаблона
            suffix = f"_{date_str.replace('.', '-')}" if date_str else f"_{datetime.now().strftime('%d-%m-%Y')}"
            summary_output = f'Сводный{suffix}.xlsx'
            
            # Работаем с новым пустым файлом
            wb = Workbook()
            ws = wb.active
            ws.title = "Сводный отчет"
            
            # Заголовок отчета
            ws.merge_cells('A1:H1')
            ws['A1'] = f"СВОДНЫЙ ОТЧЕТ ПО ПЛАВКАМ ЗА {date_str}"
            ws['A1'].font = Font(size=16, bold=True)
            ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
            
            # Заголовки таблицы
            headers = [
                'Наименование отливки',
                'Отлито',
                'Принято',
                'Второй сорт',
                'Доработка',
                'Окончательный брак',
                'Процент годности',
                'Номера плавок'
            ]
            
            # Добавляем детальные заголовки для всех типов дефектов
            # Второй сорт
            brak_headers = []
            for col in df_filtered.columns:
                if col.startswith('Второй_сорт_'):
                    name = col.replace('Второй_сорт_', '')
                    headers.append(f'Второй сорт: {name}')
                    brak_headers.append(col)
            
            # Доработка
            for col in df_filtered.columns:
                if col.startswith('Доработка_'):
                    name = col.replace('Доработка_', '')
                    headers.append(f'Доработка: {name}')
                    brak_headers.append(col)
            
            # Окончательный брак
            for col in df_filtered.columns:
                if col.startswith('Окончательный_брак_'):
                    name = col.replace('Окончательный_брак_', '')
                    headers.append(f'Брак: {name}')
                    brak_headers.append(col)
            
            # Записываем заголовки
            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=3, column=col)
                cell.value = header
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            
            # Создаем словарь для группировки номеров плавок по наименованиям
            self.log(f"Начинаем группировку номеров плавок по наименованиям отливок")
            naimenovaniya_plavki = {}
            
            # Определяем, какие колонки у нас есть в данных
            has_naimenovanie = 'Наименование_отливки' in df_filtered.columns
            
            if not has_naimenovanie:
                self.log("ВНИМАНИЕ: Колонка 'Наименование_отливки' отсутствует")
                # Если нет Наименование_отливки, проверяем альтернативы
                if 'наименование_отливки' in df_filtered.columns:
                    df_filtered['Наименование_отливки'] = df_filtered['наименование_отливки']
                    has_naimenovanie = True
                    self.log("Использую колонку 'наименование_отливки'")
            
            if not has_naimenovanie:
                raise KeyError("Колонка с наименованием отливки не найдена в данных")
            
            # Группируем номера плавок по наименованиям отливок
            for idx, row in df_filtered.iterrows():
                name = row['Наименование_отливки']
                if pd.isna(name) or not name:
                    continue
                    
                plavka = row['Номер_плавки'] if 'Номер_плавки' in row and pd.notna(row['Номер_плавки']) else ""
                
                if name not in naimenovaniya_plavki:
                    naimenovaniya_plavki[name] = set()
                
                if plavka and not pd.isna(plavka):
                    plavka_str = str(plavka).strip()
                    if plavka_str:
                        naimenovaniya_plavki[name].add(plavka_str)
            
            # Группируем данные по наименованию отливки
            grouped_data = []
            
            for name, group in df_filtered.groupby('Наименование_отливки'):
                if pd.isna(name) or not name:
                    continue
                
                # Базовая информация
                group_info = {
                    'Наименование_отливки': name,
                    'Контроль_отлито': int(group['Контроль_отлито'].sum()) if 'Контроль_отлито' in group.columns else 0,
                    'Контроль_принято': int(group['Контроль_принято'].sum()) if 'Контроль_принято' in group.columns else 0,
                    'Номера_плавок': ", ".join(sorted(list(naimenovaniya_plavki.get(name, []))))
                }
                
                # Второй сорт
                second_sort_columns = [c for c in group.columns if c.startswith('Второй_сорт_')]
                group_info['Второй_сорт'] = int(group[second_sort_columns].sum().sum()) if second_sort_columns else 0
                
                # Детализация по второму сорту
                for col in second_sort_columns:
                    group_info[col] = int(group[col].sum())
                
                # Доработка
                rework_columns = [c for c in group.columns if c.startswith('Доработка_')]
                group_info['Доработка'] = int(group[rework_columns].sum().sum()) if rework_columns else 0
                
                # Детализация по доработке
                for col in rework_columns:
                    group_info[col] = int(group[col].sum())
                
                # Окончательный брак
                reject_columns = [c for c in group.columns if c.startswith('Окончательный_брак_')]
                group_info['Окончательный_брак'] = int(group[reject_columns].sum().sum()) if reject_columns else 0
                
                # Детализация по окончательному браку
                for col in reject_columns:
                    group_info[col] = int(group[col].sum())
                
                # Процент годности
                if group_info['Контроль_отлито'] > 0:
                    group_info['Процент_годности'] = round((group_info['Контроль_принято'] / group_info['Контроль_отлито']) * 100, 2)
                else:
                    group_info['Процент_годности'] = 0
                
                grouped_data.append(group_info)
            
            self.log(f"Сформировано {len(grouped_data)} групп для отчета")
            
            # Заполняем данные из сгруппированных записей
            for idx, row_data in enumerate(grouped_data):
                current_row = 4 + idx
                
                # Заполняем основные колонки
                ws.cell(row=current_row, column=1).value = row_data['Наименование_отливки']
                ws.cell(row=current_row, column=2).value = row_data['Контроль_отлито']
                ws.cell(row=current_row, column=3).value = row_data['Контроль_принято']
                ws.cell(row=current_row, column=4).value = row_data['Второй_сорт']
                ws.cell(row=current_row, column=5).value = row_data['Доработка']
                ws.cell(row=current_row, column=6).value = row_data['Окончательный_брак']
                ws.cell(row=current_row, column=7).value = row_data['Процент_годности']
                
                # Номера плавок с выравниванием влево и переносом слов
                cell = ws.cell(row=current_row, column=8)
                cell.value = row_data['Номера_плавок']
                cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                
                # Заполняем детальные данные по дефектам
                col_offset = 8  # После основных колонок
                
                # Второй сорт
                for i, col in enumerate([c for c in df_filtered.columns if c.startswith('Второй_сорт_')], 1):
                    if col in row_data:
                        ws.cell(row=current_row, column=col_offset + i).value = row_data[col]
                
                col_offset += len([c for c in df_filtered.columns if c.startswith('Второй_сорт_')])
                
                # Доработка
                for i, col in enumerate([c for c in df_filtered.columns if c.startswith('Доработка_')], 1):
                    if col in row_data:
                        ws.cell(row=current_row, column=col_offset + i).value = row_data[col]
                
                col_offset += len([c for c in df_filtered.columns if c.startswith('Доработка_')])
                
                # Окончательный брак
                for i, col in enumerate([c for c in df_filtered.columns if c.startswith('Окончательный_брак_')], 1):
                    if col in row_data:
                        ws.cell(row=current_row, column=col_offset + i).value = row_data[col]
            
            # Добавляем автофильтр
            last_column = len(headers)
            last_column_letter = ''
            if last_column <= 26:
                last_column_letter = chr(64 + last_column)
            else:
                # Для колонки > 26 используем правильное преобразование
                first_letter_idx = (last_column - 1) // 26
                second_letter_idx = (last_column - 1) % 26
                last_column_letter = chr(65 + first_letter_idx - 1) + chr(65 + second_letter_idx)
            
            ws.auto_filter.ref = f"A3:{last_column_letter}{3 + len(grouped_data)}"
            
            # Устанавливаем ширину столбцов
            ws.column_dimensions['A'].width = 30  # Наименование отливки
            ws.column_dimensions['B'].width = 10  # Отлито
            ws.column_dimensions['C'].width = 10  # Принято
            ws.column_dimensions['D'].width = 12  # Второй сорт
            ws.column_dimensions['E'].width = 12  # Доработка
            ws.column_dimensions['F'].width = 15  # Окончательный брак
            ws.column_dimensions['G'].width = 15  # Процент годности
            ws.column_dimensions['H'].width = 40  # Номера плавок
            
            # Устанавливаем ширину для детальных колонок
            for col in range(9, 9 + len(brak_headers)):
                # Правильное преобразование номера колонки в буквенное обозначение
                if col <= 26:
                    col_letter = chr(64 + col)
                else:
                    # Разбиваем на две буквы для колонок после Z
                    first_letter_idx = (col - 1) // 26
                    second_letter_idx = (col - 1) % 26
                    col_letter = chr(65 + first_letter_idx - 1) + chr(65 + second_letter_idx)
                
                ws.column_dimensions[col_letter].width = 15
            
            # Форматирование ячеек
            for row in range(4, 4 + len(grouped_data)):
                for col in range(1, 9 + len(brak_headers)):
                    cell = ws.cell(row=row, column=col)
                    if col == 1:  # Наименование отливки
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                    elif col == 8:  # Номера плавок
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        
                    # Форматируем процент годности
                    if col == 7:  # Процент годности
                        cell.number_format = '0.00"%"'
            
            # Добавляем суммы в конце таблицы
            sum_row = 4 + len(grouped_data)
            ws.cell(row=sum_row, column=1).value = "ИТОГО:"
            ws.cell(row=sum_row, column=1).font = Font(bold=True)
            
            # Суммируем числовые столбцы
            for col in range(2, 7):
                col_letter = chr(64 + col)
                ws.cell(row=sum_row, column=col).value = f"=SUM({col_letter}4:{col_letter}{sum_row - 1})"
                ws.cell(row=sum_row, column=col).font = Font(bold=True)
            
            # Средний процент годности
            ws.cell(row=sum_row, column=7).value = f"=C{sum_row}/B{sum_row}*100"
            ws.cell(row=sum_row, column=7).font = Font(bold=True)
            ws.cell(row=sum_row, column=7).number_format = '0.00"%"'
            
            # Суммируем детальные столбцы с дефектами
            for i in range(9, 9 + len(brak_headers)):
                # Правильное преобразование номера колонки в буквенное обозначение
                if i <= 26:
                    col_letter = chr(64 + i)
                else:
                    # Разбиваем на две буквы для колонок после Z
                    first_letter_idx = (i - 1) // 26
                    second_letter_idx = (i - 1) % 26
                    col_letter = chr(65 + first_letter_idx - 1) + chr(65 + second_letter_idx)
                
                ws.cell(row=sum_row, column=i).value = f"=SUM({col_letter}4:{col_letter}{sum_row - 1})"
                ws.cell(row=sum_row, column=i).font = Font(bold=True)
            
            # Применяем стили ко всей таблице
            for row in range(3, sum_row + 1):
                for col in range(1, 9 + len(brak_headers)):
                    cell = ws.cell(row=row, column=col)
                    thin_border = Border(left=Side(style='thin'), 
                                      right=Side(style='thin'), 
                                      top=Side(style='thin'), 
                                      bottom=Side(style='thin'))
                    cell.border = thin_border
            
            # Сохраняем готовый отчет
            wb.save(summary_output)
            self.log(f"Сводный отчет успешно сохранен в файл: {summary_output}")
            return summary_output
            
        except Exception as e:
            traceback_str = traceback.format_exc()
            self.log(f"ОШИБКА при формировании сводного отчета: {str(e)}")
            self.log(traceback_str)
            raise Exception(f"Ошибка при формировании сводного отчета: {str(e)}")
    
    def _find_column_index(self, worksheet, possible_names):
        """Находит индекс колонки по возможным именам"""
        last_column = worksheet.max_column
        
        # Проверяем в строках с 1 по 10 (обычно здесь находятся заголовки)
        for header_row in range(1, 10):
            for col in range(1, last_column + 1):
                cell_value = worksheet.cell(row=header_row, column=col).value
                if cell_value:
                    cell_value = str(cell_value).strip().lower()
                    for name in possible_names:
                        if name and str(name).strip().lower() in cell_value or cell_value in str(name).strip().lower():
                            return col
        
        return None

    def create_new_report_template(self):
        """Создает новый шаблон для отчета с полными данными о дефектах"""
        try:
            self.log("Создание нового шаблона для отчета с полной информацией о дефектах")
            
            # Загружаем данные из control.xlsx для получения структуры
            self.log(f"Загрузка структуры из {self.control_file}")
            df_control = pd.read_excel(self.control_file)
            
            # Создаем новый файл отчета
            output_file = 'Полный_сводный_отчет.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.title = "Полный сводный отчет"
            
            # Заголовок отчета
            ws.merge_cells('A1:H1')
            ws['A1'] = f"ПОЛНЫЙ СВОДНЫЙ ОТЧЕТ ПО ПЛАВКАМ"
            ws['A1'].font = Font(size=16, bold=True)
            ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
            
            # Информационный блок
            ws['A3'] = "Дата отчета:"
            ws['B3'] = datetime.now().strftime('%d.%m.%Y')
            ws['A3'].font = Font(bold=True)
            ws['B3'].alignment = Alignment(horizontal='center')
            
            ws['A4'] = "Диапазон дат:"
            ws['B4'] = "Все даты"
            ws['A4'].font = Font(bold=True)
            ws['B4'].alignment = Alignment(horizontal='center')
            
            # Создаем базовые заголовки таблицы
            headers = [
                'Наименование отливки',
                'Отлито, шт.',
                'Принято, шт.',
                'Второй сорт, шт.',
                'Доработка, шт.',
                'Окончательный брак, шт.',
                'Процент годности, %',
                'Номера плавок'
            ]
            
            # Добавляем заголовки для всех типов дефектов из control.xlsx
            # По категориям дефектов
            
            # Второй сорт
            second_sort_columns = [c for c in df_control.columns if c.startswith('Второй_сорт_')]
            for col in second_sort_columns:
                headers.append(f"ВС: {col.replace('Второй_сорт_', '')}")
            
            # Доработка
            rework_columns = [c for c in df_control.columns if c.startswith('Доработка_')]
            for col in rework_columns:
                headers.append(f"Д: {col.replace('Доработка_', '')}")
            
            # Окончательный брак
            reject_columns = [c for c in df_control.columns if c.startswith('Окончательный_брак_')]
            for col in reject_columns:
                headers.append(f"БР: {col.replace('Окончательный_брак_', '')}")
            
            # Записываем заголовки
            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=6, column=col)
                cell.value = header
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
                
            # Устанавливаем ширину столбцов
            ws.column_dimensions['A'].width = 30  # Наименование отливки
            ws.column_dimensions['B'].width = 10  # Отлито
            ws.column_dimensions['C'].width = 10  # Принято
            ws.column_dimensions['D'].width = 12  # Второй сорт
            ws.column_dimensions['E'].width = 12  # Доработка
            ws.column_dimensions['F'].width = 15  # Окончательный брак
            ws.column_dimensions['G'].width = 15  # Процент годности
            ws.column_dimensions['H'].width = 40  # Номера плавок
            
            # Устанавливаем ширину для детальных колонок
            for col in range(9, 9 + len(second_sort_columns) + len(rework_columns) + len(reject_columns)):
                # Правильное преобразование номера колонки в буквенное обозначение
                if col <= 26:
                    col_letter = chr(64 + col)
                else:
                    # Разбиваем на две буквы для колонок после Z
                    first_letter_idx = (col - 1) // 26
                    second_letter_idx = (col - 1) % 26
                    col_letter = chr(65 + first_letter_idx - 1) + chr(65 + second_letter_idx)
                
                ws.column_dimensions[col_letter].width = 15
            
            # Добавляем образец строки данных
            example_row = 7
            ws.cell(row=example_row, column=1).value = "Пример наименования отливки"
            ws.cell(row=example_row, column=2).value = 100
            ws.cell(row=example_row, column=3).value = 85
            ws.cell(row=example_row, column=4).value = 5
            ws.cell(row=example_row, column=5).value = 7
            ws.cell(row=example_row, column=6).value = 3
            ws.cell(row=example_row, column=7).value = 85
            ws.cell(row=example_row, column=7).number_format = '0.00"%"'
            ws.cell(row=example_row, column=8).value = "1001/25, 1002/25, 1003/25"
            
            # Пример для детальных столбцов дефектов
            defect_col = 9
            ws.cell(row=example_row, column=defect_col).value = 2
            ws.cell(row=example_row, column=defect_col+1).value = 3
            
            # Применяем стили к примеру строки
            for col in range(1, 9 + len(second_sort_columns) + len(rework_columns) + len(reject_columns)):
                cell = ws.cell(row=example_row, column=col)
                if col == 1:  # Наименование отливки
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                elif col == 8:  # Номера плавок
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                else:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Добавляем инструкции по использованию
            instruction_row = 9
            ws.merge_cells(f'A{instruction_row}:H{instruction_row}')
            ws.cell(row=instruction_row, column=1).value = "Инструкция: Это шаблон для полного сводного отчета со всеми деталями дефектов. Используйте функцию 'Сформировать отчет' для заполнения."
            ws.cell(row=instruction_row, column=1).font = Font(italic=True)
            ws.cell(row=instruction_row, column=1).alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
            
            # Применяем границы ко всем ячейкам таблицы заголовков и примера
            for row in range(6, 9):
                for col in range(1, 9 + len(second_sort_columns) + len(rework_columns) + len(reject_columns)):
                    cell = ws.cell(row=row, column=col)
                    thin_border = Border(left=Side(style='thin'), 
                                      right=Side(style='thin'), 
                                      top=Side(style='thin'), 
                                      bottom=Side(style='thin'))
                    cell.border = thin_border
            
            # Сохраняем файл
            wb.save(output_file)
            self.log(f"Шаблон полного отчета успешно создан: {output_file}")
            return output_file
            
        except Exception as e:
            traceback_str = traceback.format_exc()
            self.log(f"ОШИБКА при создании шаблона отчета: {str(e)}")
            self.log(traceback_str)
            raise Exception(f"Ошибка при создании шаблона отчета: {str(e)}")
            
    def generate_full_report(self, date_str=None):
        """Генерирует полный отчет со всеми деталями дефектов"""
        try:
            self.log(f"Формирование полного отчета" + (f" за {date_str}" if date_str else ""))
            
            # Загружаем данные из control.xlsx
            self.log(f"Загрузка данных из {self.control_file}")
            df_control = pd.read_excel(self.control_file)
            
            # Заменяем NaN значения на 0 для всех числовых колонок
            numeric_cols = df_control.select_dtypes(include=['float64', 'int64']).columns
            df_control[numeric_cols] = df_control[numeric_cols].fillna(0)
            
            # Конвертируем даты
            self.log(f"Конвертация даты приемки")
            if 'Контроль_дата_приемки' in df_control.columns:
                df_control['Контроль_дата_приемки'] = pd.to_datetime(df_control['Контроль_дата_приемки'], format='%d.%m.%Y', errors='coerce')
            else:
                self.log(f"ВНИМАНИЕ: Колонка 'Контроль_дата_приемки' не найдена")
                df_control['Контроль_дата_приемки'] = pd.NaT
            
            # Если указана дата, фильтруем по ней
            if date_str:
                self.log(f"Фильтрация по дате {date_str}")
                selected_date = datetime.strptime(date_str, '%d.%m.%Y')
                df_filtered = df_control[df_control['Контроль_дата_приемки'].dt.date == selected_date.date()]
                
                self.log(f"После фильтрации осталось записей: {len(df_filtered)}")
                if df_filtered.empty:
                    raise ValueError(f"Нет данных за {date_str}")
            else:
                df_filtered = df_control
                date_str = "Все даты"
            
            # Создаем новый файл отчета
            suffix = f"_{date_str.replace('.', '-').replace(' ', '_')}"
            output_file = f'Полный_сводный_отчет{suffix}.xlsx'
            
            wb = Workbook()
            ws = wb.active
            ws.title = "Полный сводный отчет"
            
            # Заголовок отчета
            ws.merge_cells('A1:H1')
            ws['A1'] = f"ПОЛНЫЙ СВОДНЫЙ ОТЧЕТ ПО ПЛАВКАМ"
            ws['A1'].font = Font(size=16, bold=True)
            ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
            
            # Информационный блок
            ws['A3'] = "Дата отчета:"
            ws['B3'] = datetime.now().strftime('%d.%m.%Y')
            ws['A3'].font = Font(bold=True)
            ws['B3'].alignment = Alignment(horizontal='center')
            
            ws['A4'] = "Диапазон дат:"
            ws['B4'] = date_str
            ws['A4'].font = Font(bold=True)
            ws['B4'].alignment = Alignment(horizontal='center')
            
            # Создаем заголовки таблицы
            headers = [
                'Наименование отливки',
                'Отлито, шт.',
                'Принято, шт.',
                'Второй сорт, шт.',
                'Доработка, шт.',
                'Окончательный брак, шт.',
                'Процент годности, %',
                'Номера плавок'
            ]
            
            # Доп. колонки дефектов
            defect_columns = []
            
            # Второй сорт
            second_sort_columns = [c for c in df_filtered.columns if c.startswith('Второй_сорт_')]
            for col in second_sort_columns:
                headers.append(f"ВС: {col.replace('Второй_сорт_', '')}")
                defect_columns.append(col)
            
            # Доработка
            rework_columns = [c for c in df_filtered.columns if c.startswith('Доработка_')]
            for col in rework_columns:
                headers.append(f"Д: {col.replace('Доработка_', '')}")
                defect_columns.append(col)
            
            # Окончательный брак
            reject_columns = [c for c in df_filtered.columns if c.startswith('Окончательный_брак_')]
            for col in reject_columns:
                headers.append(f"БР: {col.replace('Окончательный_брак_', '')}")
                defect_columns.append(col)
            
            # Записываем заголовки
            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=6, column=col)
                cell.value = header
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
            
            # Создаем словарь для группировки номеров плавок по наименованиям
            naimenovaniya_plavki = {}
            
            # Определяем, какие колонки у нас есть в данных
            has_naimenovanie = 'Наименование_отливки' in df_filtered.columns
            
            if not has_naimenovanie:
                self.log("ВНИМАНИЕ: Колонка 'Наименование_отливки' отсутствует")
                # Если нет Наименование_отливки, проверяем альтернативы
                if 'наименование_отливки' in df_filtered.columns:
                    df_filtered['Наименование_отливки'] = df_filtered['наименование_отливки']
                    has_naimenovanie = True
                    self.log("Использую колонку 'наименование_отливки'")
            
            if not has_naimenovanie:
                raise KeyError("Колонка с наименованием отливки не найдена в данных")
            
            # Группируем номера плавок по наименованиям отливок
            for idx, row in df_filtered.iterrows():
                name = row['Наименование_отливки']
                if pd.isna(name) or not name:
                    continue
                    
                plavka = row['Номер_плавки'] if 'Номер_плавки' in row and pd.notna(row['Номер_плавки']) else ""
                
                if name not in naimenovaniya_plavki:
                    naimenovaniya_plavki[name] = set()
                
                if plavka and not pd.isna(plavka):
                    plavka_str = str(plavka).strip()
                    if plavka_str:
                        naimenovaniya_plavki[name].add(plavka_str)
            
            # Группируем данные по наименованию отливки
            grouped_data = []
            
            for name, group in df_filtered.groupby('Наименование_отливки'):
                if pd.isna(name) or not name:
                    continue
                
                # Базовая информация
                group_info = {
                    'Наименование_отливки': name,
                    'Контроль_отлито': int(group['Контроль_отлито'].sum()) if 'Контроль_отлито' in group.columns else 0,
                    'Контроль_принято': int(group['Контроль_принято'].sum()) if 'Контроль_принято' in group.columns else 0,
                    'Номера_плавок': ", ".join(sorted(list(naimenovaniya_plavki.get(name, []))))
                }
                
                # Второй сорт
                group_info['Второй_сорт'] = int(group[second_sort_columns].sum().sum()) if second_sort_columns else 0
                
                # Детализация по второму сорту
                for col in second_sort_columns:
                    group_info[col] = int(group[col].sum())
                
                # Доработка
                group_info['Доработка'] = int(group[rework_columns].sum().sum()) if rework_columns else 0
                
                # Детализация по доработке
                for col in rework_columns:
                    group_info[col] = int(group[col].sum())
                
                # Окончательный брак
                group_info['Окончательный_брак'] = int(group[reject_columns].sum().sum()) if reject_columns else 0
                
                # Детализация по окончательному браку
                for col in reject_columns:
                    group_info[col] = int(group[col].sum())
                
                # Процент годности
                if group_info['Контроль_отлито'] > 0:
                    group_info['Процент_годности'] = round((group_info['Контроль_принято'] / group_info['Контроль_отлито']) * 100, 2)
                else:
                    group_info['Процент_годности'] = 0
                
                grouped_data.append(group_info)
            
            self.log(f"Сформировано {len(grouped_data)} групп для отчета")
            
            # Заполняем данные из сгруппированных записей
            for idx, row_data in enumerate(grouped_data):
                current_row = 7 + idx
                
                # Заполняем основные колонки
                ws.cell(row=current_row, column=1).value = row_data['Наименование_отливки']
                ws.cell(row=current_row, column=2).value = row_data['Контроль_отлито']
                ws.cell(row=current_row, column=3).value = row_data['Контроль_принято']
                ws.cell(row=current_row, column=4).value = row_data['Второй_сорт']
                ws.cell(row=current_row, column=5).value = row_data['Доработка']
                ws.cell(row=current_row, column=6).value = row_data['Окончательный_брак']
                ws.cell(row=current_row, column=7).value = row_data['Процент_годности']
                
                # Номера плавок с выравниванием влево и переносом слов
                cell = ws.cell(row=current_row, column=8)
                cell.value = row_data['Номера_плавок']
                cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                
                # Заполняем детальные данные по дефектам
                for i, col in enumerate(defect_columns, start=9):
                    if col in row_data:
                        ws.cell(row=current_row, column=i).value = row_data[col]
            
            # Добавляем автофильтр
            # Расчет правильной последней буквы колонки
            last_column = len(headers)
            last_column_letter = ''
            if last_column <= 26:
                last_column_letter = chr(64 + last_column)
            else:
                # Для колонки > 26 используем правильное преобразование
                first_letter_idx = (last_column - 1) // 26
                second_letter_idx = (last_column - 1) % 26
                last_column_letter = chr(65 + first_letter_idx - 1) + chr(65 + second_letter_idx)
                
            ws.auto_filter.ref = f"A6:{last_column_letter}{6 + len(grouped_data)}"
            
            # Устанавливаем ширину столбцов
            ws.column_dimensions['A'].width = 30  # Наименование отливки
            ws.column_dimensions['B'].width = 10  # Отлито
            ws.column_dimensions['C'].width = 10  # Принято
            ws.column_dimensions['D'].width = 12  # Второй сорт
            ws.column_dimensions['E'].width = 12  # Доработка
            ws.column_dimensions['F'].width = 15  # Окончательный брак
            ws.column_dimensions['G'].width = 15  # Процент годности
            ws.column_dimensions['H'].width = 40  # Номера плавок
            
            # Устанавливаем ширину для детальных колонок
            for col in range(9, 9 + len(defect_columns)):
                # Правильное преобразование номера колонки в буквенное обозначение
                if col <= 26:
                    col_letter = chr(64 + col)
                else:
                    # Разбиваем на две буквы для колонок после Z
                    first_letter_idx = (col - 1) // 26
                    second_letter_idx = (col - 1) % 26
                    col_letter = chr(65 + first_letter_idx - 1) + chr(65 + second_letter_idx)
                
                ws.column_dimensions[col_letter].width = 15
            
            # Форматирование ячеек
            for row in range(7, 7 + len(grouped_data)):
                for col in range(1, 9 + len(defect_columns)):
                    cell = ws.cell(row=row, column=col)
                    if col == 1:  # Наименование отливки
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                    elif col == 8:  # Номера плавок
                        cell.alignment = Alignment(horizontal='left', vertical='center', wrapText=True)
                    else:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        
                    # Форматируем процент годности
                    if col == 7:  # Процент годности
                        cell.number_format = '0.00"%"'
            
            # Добавляем суммы в конце таблицы
            sum_row = 7 + len(grouped_data)
            ws.cell(row=sum_row, column=1).value = "ИТОГО:"
            ws.cell(row=sum_row, column=1).font = Font(bold=True)
            
            # Суммируем числовые столбцы
            for col in range(2, 7):
                col_letter = chr(64 + col)
                ws.cell(row=sum_row, column=col).value = f"=SUM({col_letter}7:{col_letter}{sum_row - 1})"
                ws.cell(row=sum_row, column=col).font = Font(bold=True)
            
            # Средний процент годности
            ws.cell(row=sum_row, column=7).value = f"=C{sum_row}/B{sum_row}*100"
            ws.cell(row=sum_row, column=7).font = Font(bold=True)
            ws.cell(row=sum_row, column=7).number_format = '0.00"%"'
            
            # Суммируем детальные столбцы с дефектами
            for i in range(9, 9 + len(defect_columns)):
                # Правильное преобразование номера колонки в буквенное обозначение
                if i <= 26:
                    col_letter = chr(64 + i)
                else:
                    # Разбиваем на две буквы для колонок после Z
                    first_letter_idx = (i - 1) // 26
                    second_letter_idx = (i - 1) % 26
                    col_letter = chr(65 + first_letter_idx - 1) + chr(65 + second_letter_idx)
                
                ws.cell(row=sum_row, column=i).value = f"=SUM({col_letter}7:{col_letter}{sum_row - 1})"
                ws.cell(row=sum_row, column=i).font = Font(bold=True)
            
            # Применяем стили ко всей таблице
            for row in range(6, sum_row + 1):
                for col in range(1, 9 + len(defect_columns)):
                    cell = ws.cell(row=row, column=col)
                    thin_border = Border(left=Side(style='thin'), 
                                      right=Side(style='thin'), 
                                      top=Side(style='thin'), 
                                      bottom=Side(style='thin'))
                    cell.border = thin_border
            
            # Сохраняем готовый отчет
            wb.save(output_file)
            self.log(f"Полный сводный отчет успешно сохранен в файл: {output_file}")
            return output_file
            
        except Exception as e:
            traceback_str = traceback.format_exc()
            self.log(f"ОШИБКА при формировании полного отчета: {str(e)}")
            self.log(traceback_str)
            raise Exception(f"Ошибка при формировании полного отчета: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    form = ControlForm()
    form.show()
    sys.exit(app.exec())
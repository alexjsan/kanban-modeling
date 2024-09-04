import os
import sys
import ast
import ctypes
import random as rnd
import pandas as pd
import tkinter as tk
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates
from tkinter import filedialog
from datetime import datetime, timedelta
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt, QDateTime
from PyQt5.QtGui import QFont, QRegion, QIcon, QPixmap, QTextOption, QColor
from PyQt5.QtWidgets import QApplication, QMainWindow, QHeaderView, QVBoxLayout, QWidget, QTableWidgetItem, QAbstractItemView, QStyledItemDelegate, QDateEdit, QFrame

root = tk.Tk()
root.withdraw()

window_info = {'width': 1200, 'height': 600}
manager_window_info = {'width': 450, 'height': 450}
preparation_modeling_window_info = {'width': 700, 'height': 400}
results_window_info = {'width': 500, 'height': 300}
add_team_window_info = {'width': 450, 'height': 450}
add_worker_window_info = {'width': 900, 'height': 450}
button_size_table = {'width': 150, 'height': 30}
button_size_menu_left = {'width': 240, 'height': 50}
button_size_manager_start = {'width': 300, 'height': 100}
button_size_manager_create = {'width': 215, 'height': 46}
button_size_manager_editor = {'width': 215, 'height': 35}
button_size_add_team = {'width': 440, 'height': 60}
table_tasks_info = {'row count': 0}
global_file_name = ''
flag_schedule_tasks = False
flag_health_factor = False
flag_fatigue_factor = False
flag_interruptions_distractions = False

tasks = pd.DataFrame(columns = ['Над-задача', 'Перед-задача', 'Назва', 'Початок', 'Час виконання', 'Працівники'])
team_list = pd.DataFrame(columns = ['ПІБ', 'Спеціальність', 'Примітки'])
workers_list = pd.DataFrame(columns = ['ПІБ', 'Працює', 'Хворий', 'Робочих годин', 'Назва задачі', 'Дні хвороби'])
all_teams_dict = {}

# Головне вікно
class MainWindow(QMainWindow):

    # Ініціалізація
    def __init__(self):
        super(MainWindow, self).__init__()

        # Налаштування зовнішнього вигляду вікна
        self.setWindowTitle('KanbanM v0.1') # Назва вікна
        self.setGeometry(100, 100, window_info['width'], window_info['height']) # Розташування та розміри
        self.setFixedSize(window_info['width'], window_info['height']) # Фіксуємо розмір (користувач не може змінити розмір вікна)
        self.setStyleSheet("MainWindow { background-color: %s }" % QColor(255, 255, 255).name()) # Змінюємо колір заднього фону вікна
        icon = QPixmap(32, 32) # Обєкт іконки
        icon.fill() # Заповнюємо ікону білим колором
        self.setWindowIcon(QIcon(icon)) # Задаємо нову іконку вікна

        # Налаштування таблиці задач
        self.table = QtWidgets.QTableWidget(self) # Обєкт таблиці
        self.table.setColumnCount(len(tasks.columns)) # Кількість стовбців
        self.table.setRowCount(len(tasks.index)) # Кільксть рядків
        self.table.setHorizontalHeaderLabels(tasks.columns) # Назви стовбців
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch) # Вирівнюємо ширину стовбців відносно ширини таблиці
        self.table.setStyleSheet('QTableView { gridline-color: rgba(0,0,0,0); border: none}') # Налаштування сітки таблиці (елементи)
        self.table.horizontalHeader().setStyleSheet("QHeaderView { border: none; }") # Налаштування сітки таблиці (заголовки стовбців)
        self.table.verticalHeader().setStyleSheet("QHeaderView { border: none; }") # Налаштування сітки таблиці (заголовки рядків)
        self.table.setGeometry(250, 50, window_info['width'] - 250, window_info['height'] - 100) # Розташування та розміри (відносно вікна)
        self.table.setDragDropMode(QAbstractItemView.InternalMove) # Drag and drop режим (дозволяє перетаскувати елементи таблиці)
        self.table.setItemDelegateForColumn(3, DateTimeDelegate(self)) # Встановлюємо тип даних для колонки "Початок" (DateTimePicker)
        self.table.cellClicked.connect(self.cell_clicked) # Подія натискання на клітинку таблиці

        # Налаштування кнопки (додавання рядка в таблицю)
        self.button1 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button1.setText('Додати рядок') # Текст кнопки
        self.button1.setGeometry(276, 563, button_size_table['width'], button_size_table['height']) # Розташування та розміри
        self.button1.clicked.connect(self.add_row_to_table) # Додаємо до події натискання кнопки функцію додавання рядку

        # Налаштування кнопки (видалення рядка)
        self.button2 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button2.setText('Видалити рядок') # Текст кнопки
        self.button2.setGeometry(431, 563, button_size_table['width'], button_size_table['height']) # Розташування та розміри
        self.button2.clicked.connect(self.delete_row_from_table) # Додаємо до події натискання кнопки функцію видалення рядка

        # Налаштування кнопки (оновлення данних в tasks)
        self.button3 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button3.setText('Зберегти таблицю') # Текст кнопки
        self.button3.setGeometry(586, 563, button_size_table['width'], button_size_table['height']) # Розташування та розміри
        self.button3.clicked.connect(lambda: self.table_to_dataframe(self.table)) # Додаємо до події натискання кнопки функцію оновлення даних
        
        # Налаштування кнопки (відновлення даних)
        self.button4 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button4.setText('Відновити дані') # Текст кнопки
        self.button4.setGeometry(741, 563, button_size_table['width'], button_size_table['height']) # Розташування та розміри
        self.button4.clicked.connect(self.dataframe_to_table) # Додаємо до події натискання кнопки функцію відновлення даних
        
        # Налаштування кнопки (Відкрити таблицю)
        self.button5 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button5.setText('Відкрити список задач') # Текст кнопки
        self.button5.setGeometry(5, 545, button_size_menu_left['width'], button_size_menu_left['height']) # Розташування та розміри
        self.button5.clicked.connect(self.open_csv_as_table) # Додаємо до події натискання кнопки функцію відкривання таблиці

        # Налаштування кнопки (Зберегти таблицю)
        self.button6 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button6.setText('Зберегти список задач') # Текст кнопки
        self.button6.setGeometry(5, 490, button_size_menu_left['width'], button_size_menu_left['height']) # Розташування та розміри
        self.button6.clicked.connect(self.save_table_as_csv) # Додаємо до події натискання кнопки функцію збереження таблиці

        # Налаштування кнопки (Менеджер команд)
        self.button7 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button7.setText('Менеджер команд') # Текст кнопки
        self.button7.setGeometry(5, 435, button_size_menu_left['width'], button_size_menu_left['height']) # Розташування та розміри
        self.button7.clicked.connect(self.open_manager) # Додаємо до події натискання кнопки функцію відкривання менеджеру подій

        # Налаштування кнопки (Моделювання)
        self.button8 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button8.setText('Розпочати моделювання') # Текст кнопки
        self.button8.setGeometry(5, 325, button_size_menu_left['width'], button_size_menu_left['height']) # Розташування та розміри
        self.button8.clicked.connect(self.open_preset) # Додаємо до події натискання кнопки функцію відкривання меню попереднього налаштування

        # Налаштування кнопки (Додавання команди в проект)
        self.button9 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button9.setText('Додати команду') # Текст кнопки
        self.button9.setGeometry(5, 380, button_size_menu_left['width'], button_size_menu_left['height']) # Розташування та розміри
        self.button9.clicked.connect(self.open_addteam) # Додаємо до події натискання кнопки функцію відкривання меню додавання команди в проект

        # Налаштування кнопки (Допомога)
        # self.button0 = QtWidgets.QPushButton(self) # Обєкт кнопки
        # self.button0.setText('Допомога') # Текст кнопки
        # self.button0.setGeometry(1095, 563, button_size_table['width'] - 50, button_size_table['height']) # Розташування та розміри
        # self.button0.clicked.connect() # Додаємо до події натискання кнопки функцію виведення справки (допомоги)

        # Текст "Список задач"
        self.lable1 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable1.setText('Список задач') # Вміст тексту
        self.lable1.setGeometry(650, 7, 150, 30) # Розташування та розміри
        font = QFont() # Обєкт шрифта
        font.setPixelSize(18) # Розмір шрифта
        self.lable1.setFont(font) # Міняємо шрифт

    # Оновлення даних DataFram'y
    def table_to_dataframe(self, table):
        global tasks
        rows = table.rowCount()
        cols = table.columnCount()
        data = []
        columns = []
        for c in range(cols):
            columns.append(table.horizontalHeaderItem(c).text())
        for r in range(rows):
            row_data = []
            for c in range(cols):
                item = table.item(r, c)
                if item is not None:
                    row_data.append(item.text())
                else:
                    row_data.append('')
            data.append(row_data)
        tasks = pd.DataFrame(data, columns=columns)
    
    # Оновлення даних таблиці
    def dataframe_to_table(self):
        global tasks
        if tasks.shape[0] > 0:
            table_tasks_info['row count'] = tasks.shape[0]
            self.table.setRowCount(table_tasks_info['row count'])
            for row in range(tasks.shape[0]):
                for col in range(tasks.shape[1]):
                    table_item = QTableWidgetItem(str(tasks.iloc[row, col]))
                    self.table.setItem(row, col, table_item)

    # Додаємо рядок в таблицю
    def add_row_to_table(self):
        table_tasks_info['row count'] += 1
        self.table.setRowCount(table_tasks_info['row count'])

    # Видаляємо рядок з таблиці
    def delete_row_from_table(self):
        if table_tasks_info['row count'] > 0 and self.table.currentRow() > -1:
            table_tasks_info['row count'] -= 1
            self.table.removeRow(self.table.currentRow())
            self.table.setCurrentItem(self.table.item(-1, -1))

    # Збереження таблиці в форматі csv
    def save_table_as_csv(self):
        file_name = filedialog.asksaveasfilename(defaultextension=".csv")
        if file_name:
            tasks.to_csv(file_name, sep=';', decimal=',', encoding='windows-1251')

    # Відкриваємо csv як таблицю
    def open_csv_as_table(self):
        global tasks
        file_name = filedialog.askopenfilename(filetypes=[('CSV files', '*.csv')])
        if file_name:
            tasks = pd.read_csv(file_name, sep=';', decimal=',', encoding='windows-1251', na_filter=False)
            if tasks.columns[1:].equals(pd.Index(['Над-задача', 'Перед-задача', 'Назва', 'Початок', 'Час виконання', 'Працівники'])):
                tasks = tasks.drop(tasks.columns[0], axis=1)
                self.dataframe_to_table()
            else:
                QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "Зміст файлу не є списком задач", QtWidgets.QMessageBox.Ok, self).exec_()

    # Відкриваємо менеджер команд
    def open_manager(self):
        self.manager = ManagerWorkersWindow()
        self.manager.show()

    # Відкриваємо вікно попереднього налаштування
    def open_preset(self):
        self.preset = PreparationForModelingWindow()
        self.preset.show()

    # Відкриваємо вікно попереднього налаштування
    def open_addteam(self):
        self.addteam = AddTeamWindow()
        self.addteam.show()

    # Натискаемо на колонку з працівниками
    def cell_clicked(self, row, column):
        if column == 5:
            self.addworker = AddWorkerWindow(row, self)
            self.addworker.show()

# Вікно менеджеру команд (створення, видалення, редагування команд)
class ManagerWorkersWindow(QWidget):

    # Ініціалізація
    def __init__(self):
        super().__init__()

        # \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ Початок \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/

        # Налаштування зовнішнього вигляду вікна
        self.setWindowTitle('Менеджер команд') # Назва вікна
        self.setGeometry(1100, 400, manager_window_info['width'], manager_window_info['height']) # Розташування та розміри
        self.setFixedSize(manager_window_info['width'], manager_window_info['height']) # Фіксуємо розмір (користувач не може змінити розмір вікна)
        self.setWindowModality(Qt.ApplicationModal) # Поки вікно активне - взаємодія з іншими вікнами неможлива
        self.setStyleSheet("ManagerWorkersWindow { background-color: %s }" % QColor(255, 255, 255).name()) # Змінюємо колір заднього фону вікна
        icon = QPixmap(32, 32) # Обєкт іконки
        icon.fill() # Заповнюємо ікону білим колором
        self.setWindowIcon(QIcon(icon)) # Задаємо нову іконку вікна

        ## Налаштування кнопки (Нова команда)
        self.button1 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button1.setText('Створити команду') # Текст кнопки
        self.button1.setGeometry(75, 122, button_size_manager_start['width'], button_size_manager_start['height']) # Розташування та розміри
        self.button1.clicked.connect(self.new_list_teams) # Додаємо до події натискання кнопки функцію створення нової команди

        ## Налаштування кнопки (Налаштувати існуючу команду)
        self.button2 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button2.setText('Редагування існуючої команди') # Текст кнопки
        self.button2.setGeometry(75, 227, button_size_manager_start['width'], button_size_manager_start['height']) # Розташування та розміри
        self.button2.clicked.connect(self.open_list_teams) # Додаємо до події натискання кнопки функцію вибору існуючої команди

        # \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ Створення нової команди \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/

        # Налаштування текстового поля (Поле для виведення команди)
        self.text_field1 = QtWidgets.QTextEdit(self) # Обєкт текстового поля
        self.text_field1.setGeometry(0, 40, 450, 275) # Розташування та розміри
        self.text_field1.setWordWrapMode(QTextOption.NoWrap) # Без переносу слів
        self.text_field1.setReadOnly(True) # Тільки читати
        self.text_field1.setFrameStyle(QFrame.NoFrame) # Прибрати кордони
        self.text_field1.hide()

        # Налаштування текстового поля (Поле для введення ПІБ)
        self.text_field2 = QtWidgets.QLineEdit(self) # Обєкт текстового поля
        self.text_field2.setGeometry(185, 324, 255, 20) # Розташування та розміри
        self.text_field2.hide() # Прибираємо обєкт

        # Налаштування текстового поля (Поле для введення спеціальності)
        self.text_field3 = QtWidgets.QLineEdit(self) # Обєкт текстового поля
        self.text_field3.setGeometry(185, 349, 255, 20) # Розташування та розміри
        self.text_field3.hide() # Прибираємо обєкт

        # Налаштування текстового поля (Поле для введення приміток)
        self.text_field8 = QtWidgets.QLineEdit(self) # Обєкт текстового поля
        self.text_field8.setGeometry(185, 374, 255, 20) # Розташування та розміри
        self.text_field8.hide() # Прибираємо обєкт

        # Текст "Список працівників"
        self.lable1 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable1.setText('Список працівників') # Вміст тексту
        self.lable1.setGeometry(125, 5, 200, 30) # Розташування та розміри
        font = QFont() # Обєкт шрифта
        font.setPixelSize(18) # Розмір шрифта
        self.lable1.setFont(font) # Міняємо шрифт
        self.lable1.hide() # Прибираємо обєкт

        # Текст "ПІБ"
        self.lable2 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable2.setText('ПІБ:') # Вміст тексту
        self.lable2.setGeometry(10, 325, 150, 30) # Розташування та розміри
        self.lable2.adjustSize() # Розмір залежить від змісту
        self.lable2.hide() # Прибираємо обєкт

        # Текст "Спеціальність"
        self.lable3 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable3.setText('Спеціальність:') # Вміст тексту
        self.lable3.setGeometry(10, 350, 150, 30) # Розташування та розміри
        self.lable3.adjustSize() # Розмір залежить від змісту
        self.lable3.hide() # Прибираємо обєкт

        # Текст "Примітки"
        self.lable8 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable8.setText('Примітки:') # Вміст тексту
        self.lable8.setGeometry(10, 375, 150, 30) # Розташування та розміри
        self.lable8.adjustSize() # Розмір залежить від змісту
        self.lable8.hide() # Прибираємо обєкт

        # Налаштування кнопки (Додати працівника)
        self.button3 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button3.setText('Додати працівника') # Текст кнопки
        self.button3.setGeometry(7, 399, button_size_manager_create['width'], button_size_manager_create['height']) # Розташування та розміри
        self.button3.clicked.connect(self.add_team) # Додаємо до події натискання кнопки функцію додавання працівника
        self.button3.hide() # Прибираємо обєкт

        # Налаштування кнопки (Зберегти команду та вийти з менеджеру)
        self.button4 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button4.setText('Зберегти та вийти') # Текст кнопки
        self.button4.setGeometry(228, 399, button_size_manager_create['width'], button_size_manager_create['height']) # Розташування та розміри
        self.button4.clicked.connect(self.save_and_close) # Додаємо до події натискання кнопки функцію збереження та виходу
        self.button4.hide() # Прибираємо обєкт

        # \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ Редагування команди \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/

        # Налаштування текстового поля (Поле для виведення списку команд)
        self.text_field4 = QtWidgets.QTextEdit(self) # Обєкт текстового поля
        self.text_field4.setGeometry(0, 40, 450, 210) # Розташування та розміри
        self.text_field4.setWordWrapMode(QTextOption.NoWrap) # Без переносу слів
        self.text_field4.setReadOnly(True) # Тільки читати
        self.text_field4.setFrameStyle(QFrame.NoFrame) # Прибрати кордони
        self.text_field4.hide() # Прибираємо обєкт

        # Налаштування текстового поля (Поле для введення ПІБ на видалення)
        self.text_field5 = QtWidgets.QLineEdit(self) # Обєкт текстового поля
        self.text_field5.setGeometry(185, 261, 255, 20) # Розташування та розміри
        self.text_field5.hide() # Прибираємо обєкт

        # Налаштування текстового поля (Поле для введення ПІБ)
        self.text_field6 = QtWidgets.QLineEdit(self) # Обєкт текстового поля
        self.text_field6.setGeometry(185, 334, 255, 20) # Розташування та розміри
        self.text_field6.hide() # Прибираємо обєкт

        # Налаштування текстового поля (Поле для введення спеціальності)
        self.text_field7 = QtWidgets.QLineEdit(self) # Обєкт текстового поля
        self.text_field7.setGeometry(185, 359, 255, 20) # Розташування та розміри
        self.text_field7.hide() # Прибираємо обєкт

        # Налаштування текстового поля (Поле для введення приміток)
        self.text_field9 = QtWidgets.QLineEdit(self) # Обєкт текстового поля
        self.text_field9.setGeometry(185, 384, 255, 20) # Розташування та розміри
        self.text_field9.hide() # Прибираємо обєкт

        # Текст "Список працівників"
        self.lable4 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable4.setText('Список працівників') # Вміст тексту
        self.lable4.setGeometry(125, 5, 200, 30) # Розташування та розміри
        font = QFont() # Обєкт шрифта
        font.setPixelSize(18) # Розмір шрифта
        self.lable4.setFont(font) # Міняємо шрифт
        self.lable4.hide() # Прибираємо обєкт

        # Текст "ПІБ працівника на видалення"
        self.lable5 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable5.setText('ПІБ (видалення):') # Вміст тексту
        self.lable5.setGeometry(10, 262, 150, 30) # Розташування та розміри
        self.lable5.adjustSize() # Розмір залежить від змісту
        self.lable5.hide() # Прибираємо обєкт

        # Текст "ПІБ"
        self.lable6 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable6.setText('ПІБ:') # Вміст тексту
        self.lable6.setGeometry(10, 335, 150, 30) # Розташування та розміри
        self.lable6.adjustSize() # Розмір залежить від змісту
        self.lable6.hide() # Прибираємо обєкт

        # Текст "Спеціальність"
        self.lable7 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable7.setText('Спеціальність:') # Вміст тексту
        self.lable7.setGeometry(10, 360, 150, 30) # Розташування та розміри
        self.lable7.adjustSize() # Розмір залежить від змісту
        self.lable7.hide() # Прибираємо обєкт

        # Текст "Примітки"
        self.lable9 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable9.setText('Примітки:') # Вміст тексту
        self.lable9.setGeometry(10, 385, 150, 30) # Розташування та розміри
        self.lable9.adjustSize() # Розмір залежить від змісту
        self.lable9.hide() # Прибираємо обєкт

        # Налаштування кнопки (Видалення працівника)
        self.button5 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button5.setText('Видалення працівника') # Текст кнопки
        self.button5.setGeometry(7, 285, button_size_manager_editor['width'] * 2 + 6, button_size_manager_editor['height']) # Розташування та розміри
        self.button5.clicked.connect(self.delete_team) # Додаємо до події натискання кнопки функцію видалення працівника
        self.button5.hide() # Прибираємо обєкт

        # Налаштування кнопки (Додати працівника)
        self.button6 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button6.setText('Додати працівника') # Текст кнопки
        self.button6.setGeometry(7, 408, button_size_manager_editor['width'], button_size_manager_editor['height']) # Розташування та розміри
        self.button6.clicked.connect(self.add_team_editor) # Додаємо до події натискання кнопки функцію додавання працівника
        self.button6.hide() # Прибираємо обєкт

        # Налаштування кнопки (Зберегти список та вийти з менеджеру)
        self.button7 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button7.setText('Зберегти та вийти') # Текст кнопки
        self.button7.setGeometry(228, 408, button_size_manager_editor['width'], button_size_manager_editor['height']) # Розташування та розміри
        self.button7.clicked.connect(self.save_and_close) # Додаємо до події натискання кнопки функцію збереження та виходу
        self.button7.hide() # Прибираємо обєкт

    # Новий список
    def new_list_teams(self):
        global global_file_name
        global team_list
        file_name = filedialog.asksaveasfilename(defaultextension=".csv")
        if file_name:
            self.button1.hide()
            self.button2.hide()
            self.text_field1.show()
            self.text_field2.show()
            self.text_field3.show()
            self.text_field8.show()
            self.lable1.show()
            self.lable2.show()
            self.lable3.show()
            self.lable8.show()
            self.button3.show()
            self.button4.show()
            global_file_name = file_name
            team_list.drop(team_list.index, inplace=True)

    # Відкрити список для редагування
    def open_list_teams(self):
        global global_file_name
        global team_list
        file_name = filedialog.askopenfilename(filetypes=[('CSV files', '*.csv')])
        if file_name:
            self.button1.hide()
            self.button2.hide()
            self.text_field4.show()
            self.text_field5.show()
            self.text_field6.show()
            self.text_field7.show()
            self.text_field9.show()
            self.lable4.show()
            self.lable5.show()
            self.lable6.show()
            self.lable7.show()
            self.lable9.show()
            self.button5.show()
            self.button6.show()
            self.button7.show()
            team_list.drop(team_list.index, inplace=True)
            team_list = pd.read_csv(file_name, sep=';', decimal=',', encoding='windows-1251', na_filter=False)
            team_list = team_list.drop(team_list.columns[0], axis=1)
            for row in team_list.values:
                self.text_field4.append(f'ПІБ: {str(row[0])}; Спеціальність: {str(row[1])}; Примітки: {str(row[2])}')
            global_file_name = file_name
            return file_name

    # Додати працівника (для вікна створення)
    def add_team(self):
        if self.text_field2.text() == '':
            QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "ПІБ працівника відсутні", QtWidgets.QMessageBox.Ok, self).exec_()
        else:
            team_list.loc[len(team_list.index)] = [self.text_field2.text(), self.text_field3.text(), self.text_field8.text()]
            self.text_field1.append(f'ПІБ: {self.text_field2.text()}; Спеціальність: {self.text_field3.text()}; Примітки: {self.text_field8.text()}')

    # Додати працівника (для вікна редагування)
    def add_team_editor(self):
        if self.text_field6.text() == '':
            QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "ПІБ працівника відсутні", QtWidgets.QMessageBox.Ok, self).exec_()
        else:
            team_list.loc[len(team_list.index)] = [self.text_field6.text(), self.text_field7.text(), self.text_field9.text()]
            self.text_field4.append(f'ПІБ: {self.text_field6.text()}; Спеціальність: {self.text_field7.text()}; Примітки: {self.text_field9.text()}')
         
    # Видалити працівника
    def delete_team(self):
        if self.text_field5.text() == '':
            QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "ПІБ працівника відсутні", QtWidgets.QMessageBox.Ok, self).exec_()
        elif not any([row[0] == self.text_field5.text() for row in team_list.values]):
            QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", f"Працівник \"{self.text_field5.text()}\" не існує", QtWidgets.QMessageBox.Ok, self).exec_()
        else:
            team_list.drop(team_list.loc[team_list['ПІБ'] == self.text_field5.text()].iloc[-1, :].name, inplace=True)
            self.text_field4.clear()
            for row in team_list.values:
                self.text_field4.append(f'ПІБ: {self.text_field6.text()}; Спеціальність: {self.text_field7.text()}; Примітки: {self.text_field9.text()}')

    # Зберегти та вийти (для вікна створення та вікна редагування)
    def save_and_close(self):
        team_list.to_csv(global_file_name, sep=';', decimal=',', encoding='windows-1251')
        self.text_field1.clear()
        self.close()

# Вікно попереднього налаштування перед моделюванням
class PreparationForModelingWindow(QWidget):

    # Ініціалізація
    def __init__(self):
        super().__init__()

        self.clickable_tasklist = False

        global flag_schedule_tasks
        flag_schedule_tasks = False
        global flag_health_factor
        flag_health_factor = False
        global flag_fatigue_factor
        flag_fatigue_factor = False
        global flag_interruptions_distractions
        flag_interruptions_distractions = False

        # Налаштування зовнішнього вигляду вікна
        self.setWindowTitle('Попереднє налаштування') # Назва вікна
        self.setGeometry(925, 400, preparation_modeling_window_info['width'], preparation_modeling_window_info['height']) # Розташування та розміри
        self.setFixedSize(preparation_modeling_window_info['width'], preparation_modeling_window_info['height']) # Фіксуємо розмір (користувач не може змінити розмір вікна)
        self.setWindowModality(Qt.ApplicationModal) # Поки вікно активне - взаємодія з іншими вікнами неможлива
        self.setStyleSheet("PreparationForModelingWindow { background-color: %s }" % QColor(255, 255, 255).name()) # Змінюємо колір заднього фону вікна
        icon = QPixmap(32, 32) # Обєкт іконки
        icon.fill() # Заповнюємо ікону білим колором
        self.setWindowIcon(QIcon(icon)) # Задаємо нову іконку вікна

        # Налаштування текстового поля (Поле для введення шляху файлу з задачами)
        self.text_field1 = QtWidgets.QLineEdit(self) # Обєкт текстового поля
        self.text_field1.setGeometry(5, 25, 660, 20) # Розташування та розміри
        self.text_field1.setReadOnly(True) # Запис даних з клавіатури неможливий

        # Налаштування текстового поля (Поле для виведення списку задач)
        self.text_field3 = QtWidgets.QTextEdit(self) # Обєкт текстового поля
        self.text_field3.setGeometry(350, 75, 340, 315) # Розташування та розміри
        self.text_field3.setWordWrapMode(QTextOption.NoWrap) # Без переносу слів
        self.text_field3.setReadOnly(True) # Тільки читати

        # Текст "Виберіть список задач"
        self.lable1 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable1.setText('Виберіть список задач:') # Вміст тексту
        self.lable1.setGeometry(5, 5, 150, 30) # Розташування та розміри
        self.lable1.adjustSize() # Розмір залежить від змісту

        # Текст "Список задач"
        self.lable4 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable4.setText('Список задач:') # Вміст тексту
        self.lable4.setGeometry(350, 55, 150, 30) # Розташування та розміри
        self.lable4.adjustSize() # Розмір залежить від змісту

        # Текст "Додаткові налаштування"
        self.lable5 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable5.setText('Додаткові налаштування:') # Вміст тексту
        self.lable5.setGeometry(5, 255, 150, 30) # Розташування та розміри
        self.lable5.adjustSize() # Розмір залежить від змісту

        # Налаштування кнопки (Вибір шляху файлу з задачами)
        self.button1 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button1.setText('...') # Текст кнопки
        self.button1.setGeometry(670, 24, 22, 22) # Розташування та розміри
        self.button1.clicked.connect(self.set_filename_taskstable) # Додаємо до події натискання кнопки функцію вибору шляху файлу з задачами

        # Налаштування кнопки (Завершення підготовки до моделювання)
        self.button3 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button3.setText('Провести моделювання') # Текст кнопки
        self.button3.setGeometry(5, 350, 340, 40) # Розташування та розміри
        self.button3.clicked.connect(self.start_modeling) # Додаємо до події натискання кнопки функцію проведення моделювання
        self.button3.setEnabled(False) # Натискання кнопки неможливе

        # Налаштування флажку (Графіки)
        self.checkbox1 = QtWidgets.QCheckBox(self) # Обєкт флажку
        self.checkbox1.setText('Надати графіки') # Текст
        self.checkbox1.setGeometry(5, 55, 20, 20) # Розташування та розміри
        self.checkbox1.adjustSize() # Розмір залежить від змісту
        self.checkbox1.stateChanged.connect(self.on_checkbox1_changed) # Опрацьовуємо зміну положення флажка

        # Налаштування флажку (Стан здоров'я)
        self.checkbox2 = QtWidgets.QCheckBox(self) # Обєкт флажку
        self.checkbox2.setText('Врахуати фактор стану здоров\'я') # Текст
        self.checkbox2.setGeometry(5, 280, 20, 20) # Розташування та розміри
        self.checkbox2.adjustSize() # Розмір залежить від змісту
        self.checkbox2.stateChanged.connect(self.on_checkbox2_changed) # Опрацьовуємо зміну положення флажка

        # Налаштування флажку (Втомленність)
        self.checkbox3 = QtWidgets.QCheckBox(self) # Обєкт флажку
        self.checkbox3.setText('Врахуати фактор втомленності') # Текст
        self.checkbox3.setGeometry(5, 300, 20, 20) # Розташування та розміри
        self.checkbox3.adjustSize() # Розмір залежить від змісту
        self.checkbox3.stateChanged.connect(self.on_checkbox3_changed) # Опрацьовуємо зміну положення флажка

        # Налаштування флажку (Переривання та відволікаючі фактори)
        self.checkbox4 = QtWidgets.QCheckBox(self) # Обєкт флажку
        self.checkbox4.setText('Врахуати переривання та відволікаючі фактори') # Текст
        self.checkbox4.setGeometry(5, 320, 20, 20) # Розташування та розміри
        self.checkbox4.adjustSize() # Розмір залежить від змісту
        self.checkbox4.stateChanged.connect(self.on_checkbox4_changed) # Опрацьовуємо зміну положення флажка

    # Вибір шляху файлу з задачами
    def set_filename_taskstable(self):
        global tasks
        file_name = filedialog.askopenfilename(filetypes=[('CSV files', '*.csv')])
        if file_name:
            tasks = pd.read_csv(file_name, sep=';', decimal=',', encoding='windows-1251', na_filter=False)
            if tasks.columns[1:].equals(pd.Index(['Над-задача', 'Перед-задача', 'Назва', 'Початок', 'Час виконання', 'Працівники'])):
                self.clickable_tasklist = True
                self.text_field1.setText(file_name)
                self.text_field3.clear()
                for row in tasks.values:
                    self.text_field3.append(f'Назва: {str(row[3])}\tНад-задача: {str(row[1])}\tПеред-задача: {str(row[2])}\tПочаток: {str(row[4])}\tЧас виконання: {str(row[5])}\tПрацівники: {str(row[6])}\t')
                if self.clickable_tasklist:
                    self.button3.setEnabled(True)
            else:
                QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "Зміст файлу не є списком задач", QtWidgets.QMessageBox.Ok, self).exec_()

    # Зміна положення флажку (checkbox1)
    def on_checkbox1_changed(self, state):
        global flag_schedule_tasks
        if state == 0:
            flag_schedule_tasks = False
        else:
            flag_schedule_tasks = True

    # Зміна положення флажку (checkbox2)
    def on_checkbox2_changed(self, state):
        global flag_health_factor
        if state == 0:
            flag_health_factor = False
        else:
            flag_health_factor = True

    # Зміна положення флажку (checkbox3)
    def on_checkbox3_changed(self, state):
        global flag_fatigue_factor
        if state == 0:
            flag_fatigue_factor = False
        else:
            flag_fatigue_factor = True

    # Зміна положення флажку (checkbox4)
    def on_checkbox4_changed(self, state):
        global flag_interruptions_distractions
        if state == 0:
            flag_interruptions_distractions = False
        else:
            flag_interruptions_distractions = True

    # Розпочаток моделювання
    def start_modeling(self):
        global tasks
        if tasks.shape[0] == 0:
            QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "Список задач має містити хоча б одну задачу", QtWidgets.QMessageBox.Ok, self).exec_()
        else:
            self.close()
            self.results = ResultsWindow()
            self.results.show()

# Вікно виводу результатів моделювання
class ResultsWindow(QWidget):

    # Ініціалізація
    def __init__(self):
        super().__init__()

        self.first_initial_show = True # Флаг, що фіксує першу ініціалізацію вікна

        # Налаштування зовнішнього вигляду вікна
        self.setWindowTitle('Результат моделювання') # Назва вікна
        self.setGeometry(300, 200, results_window_info['width'], results_window_info['height']) # Розташування та розміри
        self.setFixedSize(results_window_info['width'], results_window_info['height']) # Фіксуємо розмір (користувач не може змінити розмір вікна)
        # self.setWindowModality(Qt.ApplicationModal) # Поки вікно активне - взаємодія з іншими вікнами неможлива
        self.setStyleSheet("ResultsWindow { background-color: %s }" % QColor(255, 255, 255).name()) # Змінюємо колір заднього фону вікна
        icon = QPixmap(32, 32) # Обєкт іконки
        icon.fill() # Заповнюємо ікону білим колором
        self.setWindowIcon(QIcon(icon)) # Задаємо нову іконку вікна

        # Налаштування текстового поля (Виведення результатів моделювання)
        self.text_field1 = QtWidgets.QTextEdit(self) # Обєкт текстового поля
        self.text_field1.setGeometry(0, 0, results_window_info['width'], results_window_info['height']) # Розташування та розміри
        self.text_field1.setWordWrapMode(QTextOption.NoWrap) # Без переносу слів
        self.text_field1.setReadOnly(True) # Тільки читати
        self.text_field1.setStyleSheet("border: none;") # Видаляємо границі

    # Подія, що виникає під час появлення вікна
    def showEvent(self, event):
        if self.first_initial_show:
            self.first_initial_show = False
            super().showEvent(event)
            self.modeling(self.data_check())

    # Підготовка данних для моделювання
    def data_check(self):
        global tasks
        global workers_list
        global global_num_workers
        workers_names = []
        do_tasks = pd.DataFrame(columns = ['Над-задача', 'Перед-задача', 'Назва', 'Початок', 'Час виконання', 'Працівники'])
        notend_flag = True
        while notend_flag:
            notend_flag = False
            for index, fir_row in tasks.iterrows():
                if fir_row[1] != '':
                    for sec_row in tasks.values:
                        if fir_row[1] == sec_row[3] and sec_row[2] != '' and fir_row[2] == '':
                            tasks.loc[index, 'Перед-задача'] = sec_row[2]
                            notend_flag = True
        for row in tasks.values:
            do_tasks = do_tasks.append(pd.Series(row[1:], index=do_tasks.columns), ignore_index=True)
        do_tasks.sort_values(do_tasks.columns[3], inplace=True, ignore_index=True)
        do_tasks['Workers number'] = [self.count_workers(i, do_tasks) for i in range(len(do_tasks))]
        do_tasks['Complete'] = [False for i in range(len(do_tasks))]
        do_tasks['In progress'] = [False for i in range(len(do_tasks))]
        do_tasks['Time left'] = [row[4] for index, row in do_tasks.iterrows()]
        for index, row in do_tasks.iterrows():
            if not pd.isna(row['Працівники']):
                task_workers_str = row['Працівники']
                task_workers = ast.literal_eval(task_workers_str) if task_workers_str else []
                for w in task_workers:
                    workers_names.append(str(w['name']))
        workers_names = list(set(workers_names))
        workers_list = pd.concat([workers_list, pd.DataFrame([[name, False, False, 0, "", 0] for name in workers_names], columns=['ПІБ', 'Працює', 'Хворий', 'Робочих годин', 'Назва задачі', 'Дні хвороби'])], ignore_index=True)
        return do_tasks

    # Підрахунок кількості працівників
    def count_workers(self, row, do_tasks):
        workersNum = 0
        if not pd.isna(do_tasks.at[row, 'Працівники']):
            task_workers_str = do_tasks.at[row, 'Працівники']
            task_workers = ast.literal_eval(task_workers_str) if task_workers_str else []
            workersNum = len(task_workers)
        return workersNum

    # Дешифрування комірки працівників
    def decrypt_workers(self, workers_str):
        if not pd.isna(workers_str):
            workers = ast.literal_eval(workers_str) if workers_str else []
        return workers

    # Mоделювання
    def modeling(self, do_tasks):
        global workers_list, flag_schedule_tasks, flag_health_factor, flag_fatigue_factor, flag_interruptions_distractions
        current_date = datetime.strptime(do_tasks.iloc[0, 3], '%d-%m-%Y').date() - timedelta(days=1)
        start = current_date
        up_tasks = set([val for val in do_tasks.iloc[:, 0] if val != ''])
        fatigue_coef = 0
        work_hours = 8
        work_days = 0
        tasks_gantt_chart_data = []
        dates_chart = []
        complete_chart = []
        uncomplete_chart = []
        pie2_chart = []
        while not do_tasks.iloc[:, 7].all():
            current_date = current_date + timedelta(days=1)
            if current_date.weekday() not in [5, 6]:
                for index1, row1 in do_tasks.iterrows():
                    if not row1[7] and not row1[8]:
                        if row1[1] == '' and row1[2] not in up_tasks:
                            for w in self.decrypt_workers(row1[5]):
                                for index2, row2 in workers_list.iterrows():
                                    if w['name'] == row2[0] and not row2[1] and not row2[2]:
                                        workers_list.at[index2, 'Працює'] = True
                                        workers_list.at[index2, 'Назва задачі'] = row1[2]
                                        do_tasks.at[index1, 'In progress'] = True
                        elif row1[1] != '' and row1[2] not in up_tasks:
                            for index3, row3 in do_tasks.iterrows():
                                if row3[2] == row1[1] and row3[7]:
                                    for w in self.decrypt_workers(row1[5]):
                                        for index2, row2 in workers_list.iterrows():
                                            if w['name'] == row2[0] and not row2[1] and not row2[2]:
                                                workers_list.at[index2, 'Працює'] = True
                                                workers_list.at[index2, 'Назва задачі'] = row1[2]
                                                do_tasks.at[index1, 'In progress'] = True
                for index1, row1 in do_tasks.iterrows():
                    if row1[7] and not row1[8]:
                        for w in self.decrypt_workers(row1[5]):
                            for index2, row2 in workers_list.iterrows():
                                if w['name'] == row2[0] and not row2[1] and not row2[2]:
                                    workers_list.at[index2, 'Працює'] = True
                                    workers_list.at[index2, 'Назва задачі'] = row1[2]
                if flag_fatigue_factor:
                    fatigue_coef += rnd.random() * 0.15
                if flag_health_factor:
                    for index1, row1 in workers_list.iterrows():
                        if rnd.randint(0, 100) in range(3):
                            workers_list.at[index1, 'Хворий'] = True
                    for index1, row1 in workers_list.iterrows():
                        if row1[2] == True:
                            workers_list.at[index1, 'Дні хвороби'] += 1
                if flag_interruptions_distractions:
                    for index1, row1 in do_tasks.iterrows():
                        if row1[8] and rnd.randint(0, 100) in range(20):
                            do_tasks.at[index1, 'Time left'] += 1
                for index1, row1 in do_tasks.iterrows():
                    for w in self.decrypt_workers(row1[5]):
                        for index2, row2 in workers_list.iterrows():
                            if row1[8] and w['name'] == row2[0] and not row2[2] and row2[4] == row1[2]:
                                work_hours_with_fatigue_coef = round(work_hours - work_hours * fatigue_coef)
                                if not any(task['name'] == row1[2] for task in tasks_gantt_chart_data):
                                    tasks_gantt_chart_data.append({'name': row1[2], 'start': current_date, 'end': current_date})
                                if row1[9] <= work_hours_with_fatigue_coef:
                                    do_tasks.at[index1, 'Time left'] = 0
                                    do_tasks.at[index1, 'Complete'] = True
                                    do_tasks.at[index1, 'In progress'] = False
                                    workers_list.at[index2, 'Робочих годин'] += row1[9]
                                    row1[9] = 0
                                    for task in tasks_gantt_chart_data:
                                        if task['name'] == row1[2]:
                                            task['end'] = current_date + timedelta(days=1)
                                            break
                                    for w_all in self.decrypt_workers(row1[5]):
                                        for index3, row3 in workers_list.iterrows():
                                            if w_all['name'] == row3[0] and row3[4] == row1[2]:
                                                workers_list.at[index3, 'Працює'] = False
                                                workers_list.at[index3, 'Назва задачі'] = ""
                                else:
                                    row1[9] -= work_hours_with_fatigue_coef
                                    do_tasks.at[index1, 'Time left'] -= work_hours_with_fatigue_coef
                                    workers_list.at[index2, 'Робочих годин'] += work_hours_with_fatigue_coef
                continue_loop = True
                coplete_flag = False
                while continue_loop:
                    continue_loop = False
                    for index1, row1 in do_tasks.iterrows():
                        coplete_flag = True
                        for index2, row2 in do_tasks.iterrows():
                            if row1[2] in up_tasks and row2[0] == row1[2] and not row2[7]:
                                coplete_flag = False
                        if row1[2] in up_tasks and coplete_flag:
                            do_tasks.at[index1, 'Complete'] = True
                if flag_health_factor:
                    for index1, row1 in workers_list.iterrows():
                        if row1[2] == True and rnd.randint(0, 100) in range(80):
                            workers_list.at[index1, 'Хворий'] = False
                work_days += 1
            else:
                fatigue_coef = 0
                if flag_health_factor:
                    for index1, row1 in workers_list.iterrows():
                        if row1[2] == True and rnd.randint(0, 100) in range(80):
                            workers_list.at[index1, 'Хворий'] = False
            dates_chart.append(current_date.strftime("%m/%d"))
            complete_chart.append(sum([1 if row[7] == True else 0 for index, row in do_tasks.iterrows()]))
            uncomplete_chart.append(sum([1 if row[7] == False else 0 for index, row in do_tasks.iterrows()]))
        self.text_field1.append('Початок: ' + str(start))
        self.text_field1.append('Прогнозований кінець: ' + str(current_date))
        self.text_field1.append('Кількість робочих днів: ' + str(work_days))
        self.text_field1.append('')
        self.text_field1.append('Кількість витрачених годин на проект: ' + str(sum([row1[3] for index1, row1 in workers_list.iterrows()])) + ', з них:')
        for index1, row1 in workers_list.iterrows():
            self.text_field1.append(f'{str(row1[0])} - {str(row1[3])} годин')
        if flag_health_factor:
            self.text_field1.append('')
            self.text_field1.append('Журнал пропусків:')
            for index1, row1 in workers_list.iterrows():
                self.text_field1.append(f'{str(row1[0])}, пропуски по стану здоров\'я: {str(row1[5])}')
        if flag_schedule_tasks:
            fig, axis = plt.subplots(2, 2, figsize=(15, 9))
            fig.canvas.set_window_title('Результат моделювання (графіки)')
            axis[0, 0].set_title('Графік виконання задач')
            axis[0, 0].set_xlabel('Дата')
            axis[0, 0].set_ylabel('Задача')
            axis[0, 0].xaxis.set_major_formatter(matplotlib.dates.DateFormatter('%d-%m'))
            axis[0, 0].xaxis_date()
            axis[0, 0].set_yticks(range(len(tasks_gantt_chart_data)))
            axis[0, 0].set_yticklabels([task['name'] for task in tasks_gantt_chart_data])
            for i, task in enumerate(tasks_gantt_chart_data):
                s = matplotlib.dates.date2num(task['start'])
                e = matplotlib.dates.date2num(task['end'])
                axis[0, 0].barh(i, e - s, left=s)
            axis[0, 0].grid(True)
            axis[1, 0].bar(np.arange(len(dates_chart)), complete_chart, 0.25, label='Виконані')
            axis[1, 0].bar(np.arange(len(dates_chart)) + 0.25, uncomplete_chart, 0.25, label='Не виконані')
            axis[1, 0].set_xlabel('Дата')
            axis[1, 0].set_ylabel('Кількість задач')
            axis[1, 0].set_title('Графік виконаних/не виконаних задач')
            axis[1, 0].set_xticks(np.arange(len(dates_chart)))
            axis[1, 0].set_xticklabels(dates_chart, rotation=45)
            axis[1, 0].legend()
            axis[1, 0].grid(True)
            axis[0, 1].set_title('Кількість витрачених годин на проект')
            axis[0, 1].pie([row[3] for index, row in workers_list.iterrows()], labels = [f'{str(row[0])} - {str(row[3])} годин' for index, row in workers_list.iterrows()])
            axis[0, 1].pie([1], colors=['white'], radius=0.7)
            for index1, row1 in do_tasks.iterrows():
                for w in self.decrypt_workers(row1[5]):
                    if w['name'] not in [row['name'] for row in pie2_chart]:
                        pie2_chart.append({'name': w['name'], 'tasks': 1})
                    else:
                        for row2 in pie2_chart:
                            if row2['name'] == w['name']:
                                row2['tasks'] += 1
            axis[1, 1].set_title('Розподіл задач по працівникам')
            axis[1, 1].pie([row['tasks'] for row in pie2_chart], labels = [f"{str(row['name'])} - {str(row['tasks'])} зад." for row in pie2_chart])
            axis[1, 1].pie([1], colors=['white'], radius=0.7)
            plt.show()

# Вікно додавання команди в проект
class AddTeamWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Налаштування зовнішнього вигляду вікна
        self.setWindowTitle('Додати команду до проекту') # Назва вікна
        self.setGeometry(1100, 400, add_team_window_info['width'], add_team_window_info['height']) # Розташування та розміри
        self.setFixedSize(add_team_window_info['width'], add_team_window_info['height']) # Фіксуємо розмір (користувач не може змінити розмір вікна)
        self.setWindowModality(Qt.ApplicationModal) # Поки вікно активне - взаємодія з іншими вікнами неможлива
        self.setStyleSheet("AddTeamWindow { background-color: %s }" % QColor(255, 255, 255).name()) # Змінюємо колір заднього фону вікна
        icon = QPixmap(32, 32) # Обєкт іконки
        icon.fill() # Заповнюємо ікону білим колором
        self.setWindowIcon(QIcon(icon)) # Задаємо нову іконку вікна

        # Налаштування текстового поля (Виведення результатів моделювання)
        self.text_field1 = QtWidgets.QTextEdit(self) # Обєкт текстового поля
        self.text_field1.setGeometry(0, 40, add_team_window_info['width'], 270) # Розташування та розміри
        self.text_field1.setWordWrapMode(QTextOption.NoWrap) # Без переносу слів
        self.text_field1.setReadOnly(True) # Тільки читати
        self.text_field1.setStyleSheet("border: none;") # Видаляємо границі

        # Текст "Список команд"
        self.lable1 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable1.setText('Список команд') # Вміст тексту
        self.lable1.setGeometry(150, 5, 150, 30) # Розташування та розміри
        font = QFont() # Обєкт шрифта
        font.setPixelSize(18) # Розмір шрифта
        self.lable1.setFont(font) # Міняємо шрифт

        # Налаштування кнопки (Додати команду)
        self.button1 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button1.setText('Додати команду') # Текст кнопки
        self.button1.setGeometry(5, 320, button_size_add_team['width'], button_size_add_team['height']) # Розташування та розміри
        self.button1.clicked.connect(self.add_team) # Додаємо до події натискання кнопки функцію додавання команди

        # Налаштування кнопки (Видалити усі команди)
        self.button2 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button2.setText('Видалити усі команди') # Текст кнопки
        self.button2.setGeometry(5, 385, button_size_add_team['width'], button_size_add_team['height']) # Розташування та розміри
        self.button2.clicked.connect(self.delete_all) # Додаємо до події натискання кнопки функцію видалення усіх команд

    # Подія, що виникає під час появлення вікна
    def showEvent(self, event):
        super().showEvent(event)
        for t in all_teams_dict:
            self.text_field1.append(f'Команда \'{t}\':')
            for w in all_teams_dict[t]:
                self.text_field1.append(f'    ПІБ: {w["name"]}; Спеціальність: {w["specialty"]}; Примітки: {w["note"]}')

    # Додавання команди
    def add_team(self):
        global all_teams_dict
        file_name = filedialog.askopenfilename(filetypes=[('CSV files', '*.csv')])
        if file_name:
            team_list = pd.read_csv(file_name, sep=';', decimal=',', encoding='windows-1251', na_filter=False)
            if team_list.columns[1:].equals(pd.Index(['ПІБ', 'Спеціальність', 'Примітки'])):
                team = []
                for row in team_list.values:
                    worker = {'name': row[1], 'specialty': row[2], 'note': row[3]}
                    team.append(worker)
                all_teams_dict[os.path.splitext(os.path.basename(file_name))[0]] = team
                self.text_field1.clear()
                for t in all_teams_dict:
                    self.text_field1.append(f'Команда \'{t}\':')
                    for w in all_teams_dict[t]:
                        self.text_field1.append(f'    ПІБ: {w["name"]}; Спеціальність: {w["specialty"]}; Примітки: {w["note"]}')
            else:
                QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "Зміст файлу не є списком команд", QtWidgets.QMessageBox.Ok, self).exec_()

    # Видалення усіх команд
    def delete_all(self):
        returnMsgBoxValue = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Question, "Підтвердження", "Ви впевнені, що хочете видалити УСІ команди?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.Cancel, self).exec_()
        if returnMsgBoxValue == QtWidgets.QMessageBox.Yes:
            all_teams_dict.clear()
            self.text_field1.clear()

# Вікно додавання працівників до задачі
class AddWorkerWindow(QWidget):
    def __init__(self, row, main_window):
        super().__init__()
        self.main_window = main_window # Посилання на головне вікно
        self.row = row  # Посилання на рядок

        # Налаштування зовнішнього вигляду вікна
        self.setWindowTitle('Додати працівника') # Назва вікна
        self.setGeometry(950, 400, add_worker_window_info['width'], add_worker_window_info['height']) # Розташування та розміри
        self.setFixedSize(add_worker_window_info['width'], add_worker_window_info['height']) # Фіксуємо розмір (користувач не може змінити розмір вікна)
        self.setWindowModality(Qt.ApplicationModal) # Поки вікно активне - взаємодія з іншими вікнами неможлива
        self.setStyleSheet("AddWorkerWindow { background-color: %s }" % QColor(255, 255, 255).name()) # Змінюємо колір заднього фону вікна
        icon = QPixmap(32, 32) # Обєкт іконки
        icon.fill() # Заповнюємо ікону білим колором
        self.setWindowIcon(QIcon(icon)) # Задаємо нову іконку вікна

        # Налаштування текстового поля (Поле для виведення усіх працівників)
        self.text_field1 = QtWidgets.QTextEdit(self) # Обєкт текстового поля
        self.text_field1.setGeometry(450, 40, 450, 450) # Розташування та розміри
        self.text_field1.setWordWrapMode(QTextOption.NoWrap) # Без переносу слів
        self.text_field1.setReadOnly(True) # Тільки читати
        self.text_field1.setFrameStyle(QFrame.NoFrame) # Прибрати кордони

        # Налаштування текстового поля (Поле для виведення працівників завдання)
        self.text_field2 = QtWidgets.QTextEdit(self) # Обєкт текстового поля
        self.text_field2.setGeometry(0, 40, 450, 245) # Розташування та розміри
        self.text_field2.setWordWrapMode(QTextOption.NoWrap) # Без переносу слів
        self.text_field2.setReadOnly(True) # Тільки читати
        self.text_field2.setFrameStyle(QFrame.NoFrame) # Прибрати кордони

        # Налаштування текстового поля (Поле для введення ПІБ працівника)
        self.text_field3 = QtWidgets.QLineEdit(self) # Обєкт текстового поля
        self.text_field3.setGeometry(5, 315, 440, 20) # Розташування та розміри

        # Текст "Список працівників"
        self.lable1 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable1.setText('Список працівників') # Вміст тексту
        self.lable1.setGeometry(600, 5, 200, 30) # Розташування та розміри
        font1 = QFont() # Обєкт шрифта
        font1.setPixelSize(18) # Розмір шрифта
        self.lable1.setFont(font1) # Міняємо шрифт

        # Текст "Додані працівники"
        self.lable2 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable2.setText('Додані працівники') # Вміст тексту
        self.lable2.setGeometry(150, 5, 200, 30) # Розташування та розміри
        font2 = QFont() # Обєкт шрифта
        font2.setPixelSize(18) # Розмір шрифта
        self.lable2.setFont(font2) # Міняємо шрифт

        # Текст "Введіть ПІБ працівника"
        self.lable3 = QtWidgets.QLabel(self) # Обєкт тексту
        self.lable3.setText('Введіть ПІБ працівника:') # Вміст тексту
        self.lable3.setGeometry(5, 295, 150, 30) # Розташування та розміри
        self.lable3.adjustSize() # Розмір залежить від змісту

        # Налаштування кнопки (Додати працівника)
        self.button1 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button1.setText('Додати працівника') # Текст кнопки
        self.button1.setGeometry(5, 340, 215, 50) # Розташування та розміри
        self.button1.clicked.connect(self.add_worker_to_task) # Додаємо до події натискання кнопки функцію додавання працівника

        # Налаштування кнопки (Видалити працівника)
        self.button2 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button2.setText('Видалити працівника') # Текст кнопки
        self.button2.setGeometry(230, 340, 215, 50) # Розташування та розміри
        self.button2.clicked.connect(self.delete_worker_from_task) # Додаємо до події натискання кнопки функцію видалення працівника

        # Налаштування кнопки (Готово)
        self.button3 = QtWidgets.QPushButton(self) # Обєкт кнопки
        self.button3.setText('Готово') # Текст кнопки
        self.button3.setGeometry(5, 395, 440, 50) # Розташування та розміри
        self.button3.clicked.connect(self.close_window) # Додаємо до події натискання кнопки функцію виходу

    # Подія, що виникає під час появлення вікна
    def showEvent(self, event):
        super().showEvent(event)
        for t in all_teams_dict:
            self.text_field1.append(f'Команда \'{t}\':')
            for w in all_teams_dict[t]:
                self.text_field1.append(f'    ПІБ: {w["name"]}; Спеціальність: {w["specialty"]}; Примітки: {w["note"]}')
        item = self.main_window.table.item(self.row, 5)
        self.text_field2.clear()
        if item is not None and item.text():
            task_workers_str = item.text()
            task_workers = ast.literal_eval(task_workers_str) if task_workers_str else []
        else:
            task_workers = []
        for w in task_workers:
            self.text_field2.append(f'ПІБ: {w["name"]}; Спеціальність: {w["specialty"]}; Примітки: {w["note"]}')
     
    # Додавання працівника до задачі
    def add_worker_to_task(self):
        item = self.main_window.table.item(self.row, 5)
        self.text_field2.clear()
        if item is not None and item.text():
            task_workers_str = item.text()
            task_workers = ast.literal_eval(task_workers_str) if task_workers_str else []
        else:
            task_workers = []
        if self.text_field3.text() == "":
            QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "Введіть ПІБ працівника", QtWidgets.QMessageBox.Ok, self).exec_()
        else:
            worker_found = False
            for t in all_teams_dict:
                for w in all_teams_dict[t]:
                    if str(w["name"]) == self.text_field3.text():
                        task_workers.append(w)
                        worker_found = True
                        break
                if worker_found:
                    break
            if worker_found:
                task_workers_str = str(task_workers)
                if item is None:
                    item = QtWidgets.QTableWidgetItem()
                    self.main_window.table.setItem(self.row, 5, item)
                item.setText(task_workers_str)
            else:
                QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "Працівника не знайдено", QtWidgets.QMessageBox.Ok, self).exec_()
        for w in task_workers:
            self.text_field2.append(f'ПІБ: {w["name"]}; Спеціальність: {w["specialty"]}; Примітки: {w["note"]}')

    # Видалити працівника з задачі
    def delete_worker_from_task(self):
        item = self.main_window.table.item(self.row, 5)
        self.text_field2.clear()
        if item is not None and item.text():
            task_workers_str = item.text()
            task_workers = ast.literal_eval(task_workers_str) if task_workers_str else []
        else:
            task_workers = []
        if self.text_field3.text() == "":
            QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "Введіть ПІБ працівника", QtWidgets.QMessageBox.Ok, self).exec_()
        else:
            worker_to_delete = next((w for w in task_workers if str(w['name']) == self.text_field3.text()), None)
            if worker_to_delete:
                task_workers.remove(worker_to_delete)
                task_workers_str = str(task_workers)
                if item is None:
                    item = QtWidgets.QTableWidgetItem()
                    self.main_window.table.setItem(self.row, 5, item)
                item.setText(task_workers_str)
            else:
                QtWidgets.QMessageBox(QtWidgets.QMessageBox.Information, "Помилка", "Працівника не знайдено", QtWidgets.QMessageBox.Ok, self).exec_()
        for w in task_workers:
            self.text_field2.append(f'ПІБ: {w["name"]}; Спеціальність: {w["specialty"]}; Примітки: {w["note"]}')

    # Bийти
    def close_window(self):
        self.close()

# Делегат елемента вибору дати
class DateTimeDelegate(QtWidgets.QStyledItemDelegate):

    # Було натисно на комірку
    def createEditor(self, parent, option, index):
        editor = QtWidgets.QDateTimeEdit(parent)
        editor.setCalendarPopup(True)
        editor.setDisplayFormat("dd-MM-yyyy")
        editor.setDateTime(QDateTime.currentDateTime())
        return editor

    # Редактор заповнюємо даними з моделі
    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.DisplayRole)
        editor.setDateTime(QDateTime.fromString(value, "dd-MM-yyyy"))

    # Відредаговані дані записуємо назад у модель
    def setModelData(self, editor, model, index):
        value = editor.dateTime().toString("dd-MM-yyyy")
        model.setData(index, value, Qt.DisplayRole)

def application():
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    ctypes.windll.kernel32.CloseHandle(ctypes.windll.kernel32.GetConsoleWindow())

    app = QApplication(sys.argv)
    window = MainWindow()

    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    application()
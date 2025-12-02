import sys
import tempfile
import qrcode
import os
import re
import sqlite3
import random
import requests
from PyQt5.QtWidgets import *                                                                                                           
from PyQt5.QtGui import *                                                                                                               
from PyQt5.QtCore import *                                                                                                              
from PIL import Image                                                                                                                   
from PyQt5 import QtCore, QtGui, QtWidgets                                                                                              
from PyQt5.QtWidgets import QMessageBox, QAction, QMenu, QDialog, QScrollArea, QShortcut                                                                                             
from PyQt5.QtSvg import QSvgWidget
from reportlab.pdfgen import canvas
from numpy import asarray
import cv2
from PIL import Image, ImageDraw, ImageFont
import numpy as np
import shutil
from datetime import date, datetime
from pathlib import Path
import json
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
from reportlab.lib.pagesizes import A4, A5
from reportlab.pdfgen import canvas
from reportlab.platypus import Image as im
from reportlab.lib.units import cm
from reportlab.lib.units import mm
from io import BytesIO
import os
from PyQt5.QtCore import QCoreApplication, Qt
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
QCoreApplication.setAttribute(Qt.AA_ShareOpenGLContexts)  # Для многопоточного рендеринга
QCoreApplication.setAttribute(Qt.AA_UseSoftwareOpenGL)  # Если есть проблемы с видеодрайвером



class FileLog:
    def __init__(self):
        time = datetime.now()
        self.filelog = 'logs/' + str(time).replace(':', ',') + '.log'
        if not os.path.isdir('logs'):
            os.mkdir('logs')
    def log(self, message):
        promezhutok = os.getcwd()
        os.chdir(program_path)
        with open(self.filelog, 'a') as file:
                file.write(str(datetime.now().time()) + ': ' + message + '\n')
        os.chdir(promezhutok)

logger = FileLog()

def openImg(path):
    f = open(path, "rb")
    chunk = f.read()
    chunk_arr = np.frombuffer(chunk, dtype=np.uint8)
    return cv2.imdecode(chunk_arr, cv2.IMREAD_COLOR)

def saveImg(path, image):
    is_success, im_buf_arr = cv2.imencode(".png", image)
    im_buf_arr.tofile(path)

program_path = os.getcwd()
ishodniki_path = os.path.join(os.getcwd(), 'Этикетки')
reserv_path = os.path.join(os.getcwd(), 'Резервная копия этикеток')
desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
logger.log( f"путь рабочего стола: {desktop_path}")

shrift = 'razmer_shifra_file.txt'

try:
    with open(shrift, 'r') as file:
        number = int(file.read())
    logger.log(f"- Прочитан размер шрифта из файла: {number} -")
except FileNotFoundError:
    logger.log("- Файл с размером шрифта не найден. Создается файл с числом 12... -")
    with open(shrift, 'w') as file:
        file.write('12')
        number = 12



class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("QR")
        MainWindow.setFixedSize(600, 450)
        MainWindow.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.ArrowCursor))
        MainWindow.setWindowTitle("РСЗ.Маркиратор")

        self.razmer_shrifta = number
        
        self.ShifratorButton = QtWidgets.QPushButton(MainWindow)
        self.ShifratorButton.setGeometry(QtCore.QRect(0, 30, 200, 30))
        self.ShifratorButton.setObjectName("blk")
        self.ShifratorButton.setFont(QFont("Helvetica", self.razmer_shrifta))
        self.ShifratorButton.setText("Изделия")

        self.UpakovkaButton = QtWidgets.QPushButton(MainWindow)
        self.UpakovkaButton.setGeometry(QtCore.QRect(200, 30, 200, 30))
        self.UpakovkaButton.setObjectName("blk")
        self.UpakovkaButton.setFont(QFont("Helvetica", self.razmer_shrifta))
        self.UpakovkaButton.setText("Упаковка")
   
        self.DeshifratorButton = QtWidgets.QPushButton(MainWindow)
        self.DeshifratorButton.setGeometry(QtCore.QRect(400, 30, 200, 30))
        self.DeshifratorButton.setObjectName("blk")
        self.DeshifratorButton.setFont(QFont("Helvetica", self.razmer_shrifta))
        self.DeshifratorButton.setText("Поиск")

        self.Box_SelectProduct = QtWidgets.QComboBox(MainWindow)
        self.Box_SelectProduct.setGeometry(QtCore.QRect(55, 80, 230, 26))
        self.Box_SelectProduct.setFont(QFont("Helvetica", self.razmer_shrifta))
        self.Box_SelectProduct.addItem("Выберите изделие")
        self.Box_SelectProduct.addItem("КИП")
        self.Box_SelectProduct.addItem("БСЗ")
        self.Box_SelectProduct.addItem("АЗ")
        self.Box_SelectProduct.addItem("ПАЗ")
        self.Box_SelectProduct.addItem("МСЭС")
        self.Box_SelectProduct.setCurrentText("Выберите изделие")

        self.gazprom = QtWidgets.QRadioButton(MainWindow)                                                                               
        self.gazprom.setGeometry(QtCore.QRect(55, 120, 150, 20))                                                                                                                                                       
        self.gazprom.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                                    
        self.gazprom.setText("Газпром стройТЭК")

        self.transneft = QtWidgets.QRadioButton(MainWindow)                                                                             
        self.transneft.setGeometry(QtCore.QRect(230, 120, 140, 20))                                                                                                                                                  
        self.transneft.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                                  
        self.transneft.setText("Транснефть")                                                                                                           

        self.zakazchik = QtWidgets.QLineEdit(MainWindow)                                                                                
        self.zakazchik.setGeometry(QtCore.QRect(330, 155, 225, 20))                                                                                                                                                          
        self.zakazchik.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                                  
        self.zakazchik.hide()                                                                                                           

        self.label_zakazchik = QtWidgets.QLabel(MainWindow)                                                                             
        self.label_zakazchik.setGeometry(QtCore.QRect(230, 155, 90, 20))                                                                                                                                        
        self.label_zakazchik.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                            
        self.label_zakazchik.setText("Заказчик")
        self.label_zakazchik.hide()                                                                                                     

        self.label_partya = QtWidgets.QLabel(MainWindow)                                                                                
        self.label_partya.setGeometry(QtCore.QRect(55, 195, 140, 20))                                                                                                                                                  
        self.label_partya.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                               
        self.label_partya.setText("Номер заказа")

        self.label_kol = QtWidgets.QLabel(MainWindow)                                                                                   
        self.label_kol.setGeometry(QtCore.QRect(55, 230, 120, 20))                                                                                                                                                          
        self.label_kol.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                                  
        self.label_kol.setText("Количество")

        self.label_modify = QtWidgets.QLabel(MainWindow)                                                                                
        self.label_modify.setGeometry(QtCore.QRect(55, 265, 140, 20))                                                                                                                                               
        self.label_modify.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                               
        self.label_modify.setText("Модификация")

        self.auto_date_checkbox = QtWidgets.QCheckBox(MainWindow)
        self.auto_date_checkbox.setGeometry(QtCore.QRect(55, 300, 260, 20))
        self.auto_date_checkbox.setFont(QFont("Helvetica", self.razmer_shrifta))
        self.auto_date_checkbox.setText("Ввод даты вручную")
        self.auto_date_checkbox.setChecked(False)  # По умолчанию выбрана автоматическая дата
        self.auto_date_checkbox.stateChanged.connect(self.toggle_date_input)                                                                                                      

        self.partya = QtWidgets.QLineEdit(MainWindow)                                                                                   
        self.partya.setGeometry(QtCore.QRect(295, 195, 260, 20))                                                                                                                                                  
        self.partya.setFont(QFont("Helvetica", self.razmer_shrifta)) 
        self.partya.setValidator(QRegularExpressionValidator(QRegularExpression("^[0-9]{1,10}$")))                                                                                    

        self.kol = QtWidgets.QLineEdit(MainWindow)                                                                                      
        self.kol.setGeometry(QtCore.QRect(295, 230, 260, 20))                                                                                                                                                               
        self.kol.setFont(QFont("Helvetica", self.razmer_shrifta))
        self.kol.setValidator(QRegularExpressionValidator(QRegularExpression("^[1-9]\\d{0,3}$")))                                                                                     

        self.modify = QtWidgets.QLineEdit(MainWindow)                                                                                   
        self.modify.setGeometry(QtCore.QRect(295, 265, 260, 20))                                                                                                                                                               
        self.modify.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                                 

        self.data = QtWidgets.QLineEdit(MainWindow)                                                                                     
        self.data.setGeometry(QtCore.QRect(295, 300, 260, 20))                                                                                                                                                                
        self.data.setFont(QFont("Helvetica", self.razmer_shrifta)) 
        self.data.setPlaceholderText('  НН.ГГГГ')
        self.data.setValidator(QRegularExpressionValidator(QRegularExpression(r'^(0?[1-9]|[1-4][0-9]|5[0-3])\.(20[0-9]{2})$')))     
        self.data.hide()                                                                          

        self.label_logo = QSvgWidget('Logo.svg')                                                                                        
        self.label_logo.setGeometry(460, 340, 140, 140)                                                                                 
        self.label_logo.setParent(MainWindow)                                                                                           

        self.label_version = QtWidgets.QLabel(MainWindow)                                                                               
        self.label_version.setGeometry(QtCore.QRect(574, 433, 26, 8))                                                                                                                                         
        self.label_version.setStyleSheet("color: #ffffff;")                                                                             
        self.label_version.setFont(QFont("Helvetica", 6))                                                                               
        self.label_version.setText("V 1.3")

        self.GenerateButton = QtWidgets.QPushButton(MainWindow)                                                                         
        self.GenerateButton.setGeometry(QtCore.QRect(205, 380, 170, 50))                                                                                            
        self.GenerateButton.setObjectName("red")                                                                             
        self.GenerateButton.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                             
        self.GenerateButton.setText("Генерировать")

        self.progressBar = QtWidgets.QProgressBar(MainWindow)
        self.progressBar.setGeometry(QtCore.QRect(55, 335, 500, 20))
        self.progressBar.hide()

        self.TaraButton = QtWidgets.QPushButton(MainWindow)                                                                           
        self.TaraButton.setGeometry(QtCore.QRect(205, 380, 170, 50))                                                                                                                                              
        self.TaraButton.setObjectName("red")                                                                                 
        self.TaraButton.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                               
        self.TaraButton.setText("Создать")
        self.TaraButton.hide()       

        self.packaging_table = QtWidgets.QTableWidget(MainWindow)
        self.packaging_table.setGeometry(QtCore.QRect(25, 80, 550, 250))
        self.packaging_table.setFont(QFont("Helvetica", self.razmer_shrifta))
        self.packaging_table.setColumnCount(2)
        self.packaging_table.setHorizontalHeaderLabels(["Наименование", "Количество"])
        self.packaging_table.setRowCount(1)  # Начинаем с одной строки
        self.packaging_table.setColumnWidth(0, 400)  # Широкий столбец для Наименования
        self.packaging_table.setColumnWidth(1, 130)  # Узкий столбец для Количества
        self.packaging_table.hide()  # Скрыто по умолчанию
        self.packaging_table.setStyleSheet("QTableWidget { color: #ffffff; }")
        
        # Кнопка для добавления новой строки
        self.add_row_button = QtWidgets.QPushButton(MainWindow)
        self.add_row_button.setGeometry(QtCore.QRect(25, 340, 150, 30))
        self.add_row_button.setFont(QFont("Helvetica", self.razmer_shrifta))
        self.add_row_button.setText("Добавить строку")
        self.add_row_button.hide()
        self.add_row_button.clicked.connect(self.add_table_row) 

        self.VigruzButton = QtWidgets.QPushButton(MainWindow)                                                                           
        self.VigruzButton.setGeometry(QtCore.QRect(205, 380, 160, 50))                                                                                                                                            
        self.VigruzButton.setObjectName("red")                                                                                 
        self.VigruzButton.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                               
        self.VigruzButton.setText("Выгрузить")
        self.VigruzButton.hide()                                                                                             

        self.label_shifr = QtWidgets.QLabel(MainWindow)                                                                                 
        self.label_shifr.setGeometry(QtCore.QRect(25, 130, 120, 20))                                                                                                                                                 
        self.label_shifr.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                                
        self.label_shifr.setText("Шифр или с/н")
        self.label_shifr.hide()                                                                                                         

        self.code = QtWidgets.QLineEdit(MainWindow)                                                                                     
        self.code.setGeometry(QtCore.QRect(150, 130, 260, 20))                                                                                                                                                             
        self.code.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                                       
        self.code.hide()                                                                                                                

        self.ReGenerateButton = QtWidgets.QPushButton(MainWindow)                                                                       
        self.ReGenerateButton.setGeometry(QtCore.QRect(435, 115, 140, 50))                                                               
        self.ReGenerateButton.setObjectName("red")                                                                         
        self.ReGenerateButton.setFont(QFont("Helvetica", self.razmer_shrifta))                                                                           
        self.ReGenerateButton.setText("Поиск")
        self.ReGenerateButton.hide()

        self.text_browser = QtWidgets.QTextBrowser(MainWindow)
        self.text_browser.setGeometry(QtCore.QRect(25, 185, 555, 180))
        self.text_browser.setFont(QFont("Helvetica", self.razmer_shrifta))
        self.text_browser.setStyleSheet("color: #ffffff; ")
        self.text_browser.hide()

        QtCore.QMetaObject.connectSlotsByName(MainWindow)     

        zakachik_group = QButtonGroup(MainWindow)
        zakachik_group.addButton(self.gazprom)
        zakachik_group.addButton(self.transneft)
        zakachik_group.buttonClicked.connect(self.new_elements)       
                                                              

        self.UpakovkaButton.clicked.connect(self.PackButton)
        self.DeshifratorButton.clicked.connect(self.DeshifrButton)
        self.ShifratorButton.clicked.connect(self.ShifrButton)
        self.GenerateButton.clicked.connect(self.generate_clicked)
        self.ReGenerateButton.clicked.connect(self.regenerate_clicked)
        self.VigruzButton.clicked.connect(self.vigruzka_clicked)
        self.TaraButton.clicked.connect(self.upakovka_clicked)
        self.products_data = self.load_json_data()

        menubar = MainWindow.menuBar()
        funkcia_menu = menubar.addMenu('Функции')
        spravka_menu = menubar.addMenu('Справка')

        bolshe_shrift = QAction('Увеличить шрифт       CTRL +', MainWindow)
        bolshe_shrift.triggered.connect(lambda: self.izmenenie_razmera_shrifta(1))
        QShortcut(QtGui.QKeySequence(QtCore.Qt.CTRL + QtCore.Qt.Key_Plus), MainWindow, bolshe_shrift.trigger)
        QShortcut(QtGui.QKeySequence(QtCore.Qt.CTRL + QtCore.Qt.Key_Equal), MainWindow, bolshe_shrift.trigger)
        funkcia_menu.addAction(bolshe_shrift) 

        menshe_shrift = QAction('Уменьшить шрифт     CTRL -', MainWindow)
        menshe_shrift.triggered.connect(lambda: self.izmenenie_razmera_shrifta(-1))
        QShortcut(QtGui.QKeySequence(QtCore.Qt.CTRL + QtCore.Qt.Key_Minus), MainWindow, menshe_shrift.trigger)
        funkcia_menu.addAction(menshe_shrift)
           
        instrukcia = QAction('Инструкция', MainWindow)
        instrukcia.triggered.connect(lambda: self.instrukcia_po_programme(MainWindow))
        spravka_menu.addAction(instrukcia)    
# Проверка на блокировку при запуске
        if not self.check_expiry_and_internet():
            # Если False — закрываем программу
            MainWindow.close()
            sys.exit(1)  # Выход с ошибкой


    def izmenenie_razmera_shrifta(self, delta):
        self.razmer_shrifta += delta
        logger.log(f'- Изменен размер шрифта до {self.razmer_shrifta} -')
        font = QtGui.QFont()
        font.setPointSize(self.razmer_shrifta)
        self.update_shrifta(font, MainWindow)
    def update_shrifta(self, font, widget):
        for child in widget.findChildren(QtWidgets.QWidget):
            if isinstance(child, QtWidgets.QWidget) and not isinstance(child, QtWidgets.QMenuBar) and child != self.label_version:
                child.setFont(font)
        logger.log('Шрифт применен ко всем элементам')
        with open(shrift, 'w') as file:
            file.write(str(self.razmer_shrifta))
            logger.log(f"Размер шрифта в файле обновлено на: {self.razmer_shrifta}")

    def instrukcia_po_programme(self, parent=None):
        if os.name == 'nt':  # Проверка, что операционная система - Windows
            os.startfile(program_path + '/' + 'Instrukcia.pdf')

    def toggle_date_input(self, state):
        if state == Qt.Checked:
            self.data.show()
        else:
            self.data.hide()

    def new_elements(self, button):                                                                                                                                                                                                                         
            if button == self.gazprom:                                                                                                                                                      
                self.label_zakazchik.setVisible(False)                                                                                   
                self.zakazchik.setVisible(False)                                                                            
            elif button == self.transneft:                                                                                                                                                                         
                self.label_zakazchik.setVisible(True)                                                                                   
                self.zakazchik.setVisible(True)                                                                                        

    def DeshifrButton(self):  
        self.progressBar.hide()                                                                                                          
        self.Box_SelectProduct.hide()                                                                                                   
        self.transneft.hide()                                                                                                           
        self.gazprom.hide()                                                                                                                                                                                                                            
        self.label_partya.hide()                                                                                                        
        self.label_zakazchik.hide()                                                                                                     
        self.zakazchik.hide()                                                                                                           
        self.label_kol.hide()                                                                                                                                                                                                                       
        self.label_modify.hide()                                                                                                                                                                                                                                                                                                                          
        self.partya.hide()                                                                                                              
        self.kol.hide()                                                                                                                 
        self.modify.hide()                                                                                                              
        self.data.hide()                                                                                                                                                                                                                                                                                                              
        self.GenerateButton.hide()                                                                                                      
        self.label_kol.hide()                                                                                                           
        self.kol.hide()                                                                                                                 
        self.ReGenerateButton.show()                                                                                                    
        self.label_shifr.show()                                                                                                         
        self.code.show()                                                                                                                
        self.text_browser.show()  
        self.VigruzButton.hide()
        self.TaraButton.hide() 
        self.auto_date_checkbox.hide()
        self.auto_date_checkbox.setChecked(False) 
        self.packaging_table.hide()  # Показать таблицу
        self.add_row_button.hide()   # Показать кнопку добавления строки                                                                                                   
        
    def ShifrButton(self):
        self.progressBar.hide()                                                                                                              
        self.Box_SelectProduct.show()                                                                                                   
        self.transneft.show()                                                                                                           
        self.gazprom.show()                                                                                                                                                                                                                            
        self.label_partya.show()                                                                                                        
        self.label_kol.show()                                                                                                           
        self.label_modify.show()                                                                                                        
        self.partya.show()                                                                                                              
        self.kol.show()                                                                                                                 
        self.modify.show()                                                                                                              
        self.data.hide()                                                                                                                                                                                                                       
        self.GenerateButton.show()                                                                                                      
        self.label_kol.show()                                                                                                           
        self.kol.show()                                                                                                                 
        self.ReGenerateButton.hide()                                                                                                                                                                                                         
        self.label_shifr.hide()                                                                                                         
        self.code.hide()                                                                                                                
        self.text_browser.hide()                                                                                                          
        self.VigruzButton.hide()
        self.TaraButton.hide()
        self.auto_date_checkbox.show()
        self.packaging_table.hide()  # Показать таблицу
        self.add_row_button.hide()   # Показать кнопку добавления строки                                                                                                        

    def PackButton(self):
        self.progressBar.hide()
        self.Box_SelectProduct.hide()                                                                                                   
        self.transneft.hide()                                                                                                           
        self.gazprom.hide()                                                                                                                                                                                                                             
        self.label_partya.hide()                                                                                                        
        self.label_zakazchik.hide()                                                                                                     
        self.zakazchik.hide()                                                                                                           
        self.label_kol.hide()                                                                                                                                                                                                                        
        self.label_modify.hide()                                                                                                                                                                                                                                                                                                                        
        self.partya.hide()                                                                                                              
        self.kol.hide()                                                                                                                 
        self.modify.hide()                                                                                                              
        self.data.hide()                                                                                                                                                                                                                                                                                                                                
        self.GenerateButton.hide()                                                                                                      
        self.label_kol.hide()                                                                                                           
        self.kol.hide()                                                                                                                 
        self.ReGenerateButton.hide()                                                                                                    
        self.label_shifr.hide()                                                                                                         
        self.code.hide()                                                                                                                
        self.text_browser.hide()  
        self.VigruzButton.hide()
        self.TaraButton.show()
        self.auto_date_checkbox.hide()
        self.auto_date_checkbox.setChecked(False)
        self.packaging_table.show()  # Показать таблицу
        self.add_row_button.show()   # Показать кнопку добавления строки

    def add_table_row(self):
        row_count = self.packaging_table.rowCount()
        self.packaging_table.insertRow(row_count)
        logger.log(f"Добавлена новая строка в таблицу. Текущие строки: {row_count + 1}")

    # ПРОВЕРКА ДАТЫ ДЛЯ БЛОКИРОВКИ
    def check_expiry_and_internet(self):
        # Укажи "дедлайн" — после этой даты блокировка сработает
        expiry_date = date(2026, 1, 12)  # Год, месяц, день — измени на нужный
        
        current_date = date.today()
        
        if current_date <= expiry_date:
            logger.log("Дата в норме, программа работает.")
            return True  # Всё ок, продолжаем
        
        # Дата истекла — проверяем интернет
        try:
            response = requests.get("https://www.google.com", timeout=5)  # Простой запрос
            if response.status_code == 200:
                # Интернет есть — блокируем
                QMessageBox.critical(None, "Программа истекла", 
                                    f"Программа больше не поддерживается.\n"
                                    "Обратитесь к разработчику.")
                logger.log("Блокировка: дата истекла и интернет доступен.")
                return False  # Блокируем
            else:
                # Интернет "нет" (или ошибка) — разрешаем работу
                logger.log("Дата истекла, но интернета нет — разрешаем работу.")
                return True
        except requests.exceptions.RequestException:
            # Ошибка запроса = нет интернета — разрешаем
            logger.log("Дата истекла, ошибка интернета — разрешаем работу.")
            return True
    #_____________________________

    @staticmethod                                                                                                                       
    def check_format(data_text):                                                                                                        
        min_week = 1                                                                                                                    
        max_week = 53                                                                                                                   
        min_year = 2023                                                                                                                 
        max_year = 2124                                                                                                                 
        try:                                                                                                                            
            if not re.match(r'^\d+\.\d{4}$', str(data_text)):                                                                           
                return False                                                                                                            
            
            value_parts = str(data_text).split('.')                                                                                     
            week_part = int(value_parts[0])                                                                                             
            year_part = int(value_parts[1])                                                                                             

            if week_part < min_week or week_part > max_week:                                                                            
                return False                                                                                                            
            
            if year_part < min_year or year_part >= max_year:                                                                            
                return False                                                                                                            
            
            return True                                                                                                                 
        
        except ValueError:                                                                                                              
            return False                                                                                                                

    @staticmethod                                                                                                                       
    def shifrator(partya_text, kol_text, SelectProduct_text, radio, inf):                                                               
        kol_square = int(kol_text) * int(kol_text)                                                                                      
        random_number1 = random.randint(10, 99)                                                                                         
        random_number2 = random.randint(10, 99)                                                                                         
        random_number3 = random.randint(10, 99)                                                                                         
        count_shif = 3644                                                                                                               

        shifr = partya_text[:5] + str(kol_square) + partya_text[5:]                                                                     
        shifr = shifr[:3] + str(random_number1) + shifr[3:]                                                                             

        if shifr[0] == '0':                                                                                                             
            shifr = int(shifr) - count_shif                                                                                             
            shifr = str(shifr)                                                                                                          
            shifr = '0' + shifr                                                                                                         
        else:                                                                                                                           
            shifr = int(shifr) - count_shif                                                                                             
            shifr = str(shifr)                                                                                                          

        shifr = shifr[:9] + str(random_number3) + shifr[9:]                                                                             
        shifr = shifr[:11] + str(random_number2) + shifr[11:]                                                                           

        if SelectProduct_text == "КИП":                                                                                                 
            pos = '1'                                                                                                                   
        elif SelectProduct_text == "БСЗ":                                                                                               
            pos = '2'                                                                                                                   
        elif SelectProduct_text == "АЗ":                                                                                                
            pos = '3'                                                                                                                   
        elif SelectProduct_text == "ПАЗ":                                                                                               
            pos = '4'                                                                                                                   
        else:                                                                                                                           
            pos = '5'                                                                                                                   
        shifr = pos + shifr                                                                                                             

        if inf == 0:                                                                                                                    
            shifr = '0' + shifr                                                                                                         
        else:                                                                                                                           
            shifr = '1' + shifr                                                                                                         
    
        if radio == 'g':                                                                                                                
            shifr = '88' + shifr                                                                                                        
        elif radio == 't':                                                                                                              
            shifr = '44' + shifr                                                                                                        
        else:                                                                                                                           
            shifr = '66' + shifr                                                                                                        
        return shifr                                                                                                                    

    @staticmethod                                                                                                                       
    def dublicate(info_base, SelectProduct_text, partya_text, modify_text, data_text):                                                  
        connection = sqlite3.connect(info_base)                                                                                         
        cursor = connection.cursor()                                                                                                    
        query = f"SELECT Номер_партии, Модификация, Дата_изготовления FROM {SelectProduct_text} WHERE Номер_партии = ? AND \
            Модификация = ? AND Дата_изготовления = ?"                                                                                  
        cursor.execute(query, (partya_text, modify_text, data_text))                                                                    
        result = cursor.fetchone()                                                                                                      
        connection.close()                                                                                                              
        return result                                                                                                                   

    def zapros(self, info_base, izdelie, nomer, text):                                                                                 
        connect = sqlite3.connect(info_base)                                                                                            
        cursor = connect.cursor()                                                                                                       
        cursor.execute("SELECT Компания, Номер_партии, Модификация, Порядковый_номер, Дата_изготовления, Серийный_номер \
            FROM {} WHERE Шифр = ?".format(izdelie), (nomer,))                                                                          
        results = cursor.fetchall()                                                                                                     
        if len(results) == 0:                                                                                                           
            self.text_browser.append(f"Такого кода нет. Косяк")                                                                         
            self.text_browser.update()                                                                                                  
        else:                                                                                                                           
            for row in results:                                                                                                         
                self.text_browser.append(f"Компания: {text}{row[0]}\nНомер партии: {row[1]}\nМодификация: {row[2]}\
                    \nПорядковый номер: {row[3]}\nДата производства: {row[4]}\nСерийный номер: {row[5]}")                                                         
                self.text_browser.update()                                                                                              
        connect.close()                                                                                                                                                                                                                                

    def create_passport(self, json_Product, desktop_path, serial_num, new_modify_text, modify_text, shifr_part, kol_text, data_text):
        new_name = f"Паспорт {new_modify_text} {serial_num}.pdf"
        temp_pdf = os.path.join(program_path, new_name)
        final_pdf = os.path.join(desktop_path, new_name)
        
        shutil.copy(json_Product["Путь_паспорта"], temp_pdf)

        font_path = os.path.join(program_path, "GOST2304_TypeB.ttf")
        pdfmetrics.registerFont(TTFont('TypeB', font_path))

        packet_page_1 = BytesIO()
        packet_page_N = BytesIO()
        
        can_1 = canvas.Canvas(packet_page_1, pagesize=letter)
        can_1.setFont('TypeB', 12)
        can_1.drawString(json_Product["Модификация_1_X"], json_Product["Модификация_1_Y"], modify_text)
        can_1.save()

        # Генерируем QR-код и сохраняем во временный файл
        qr = qrcode.QRCode(version=1, box_size=10, border=2)
        qr.add_data(shifr_part)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        # Сохраняем QR-код во временный файл
        qr_temp_path = os.path.join(tempfile.gettempdir(), "temp_qr.png")
        qr_img.save(qr_temp_path)

        can_N = canvas.Canvas(packet_page_N, pagesize=letter)
        can_N.setFont('TypeB', 12)
        can_N.drawString(json_Product["Серийный_номер_N_X"], json_Product["Серийный_номер_N_Y"], serial_num)
        can_N.drawString(json_Product["Модификация_N_X"], json_Product["Модификация_N_Y"], modify_text)
        can_N.drawString(json_Product["Количество_N_X"], json_Product["Количество_N_Y"], kol_text)
        can_N.drawString(json_Product["Дата_N_Х"], json_Product["Дата_N_Y"], data_text)
        can_N.drawImage(qr_temp_path, json_Product["QR_N_X"], json_Product["QR_N_Y"], width=70, height=70)
        can_N.save()

        # Наложение изменений на PDF
        reader = PdfReader(temp_pdf)
        writer = PdfWriter()

        for i in range(len(reader.pages)):
            page = reader.pages[i]
            if i == 0:  # 1-я страница
                packet_page_1.seek(0)
                overlay = PdfReader(packet_page_1)
                page.merge_page(overlay.pages[0])

        for i in range(len(reader.pages)):
            page = reader.pages[i]
            if i == int(json_Product["Номер_страницы_N"])-1:  # N-я страница
                packet_page_N.seek(0)
                overlay = PdfReader(packet_page_N)
                page.merge_page(overlay.pages[0])
            writer.add_page(page)

        # Сохраняем результат
        with open(temp_pdf, "wb") as f:
            writer.write(f)
        
        # Удаляем временный файл QR-кода
        try:
            os.remove(qr_temp_path)
        except:
            pass
        
        shutil.move(temp_pdf, final_pdf)
        
# сделать ПСИ
    def create_PSI(self, json_Product, desktop_path, serial_num, new_modify_text, modify_text, shifr_part, kol_text, data_text, partya_text):
        new_name = f"ПСИ {new_modify_text} {serial_num}.pdf"
        temp_pdf = os.path.join(program_path, new_name)
        final_pdf = os.path.join(desktop_path, new_name)
        
        shutil.copy(json_Product["Путь_ПСИ"], temp_pdf)

        font_path = os.path.join(program_path, "GOST2304_TypeB.ttf")
        pdfmetrics.registerFont(TTFont('TypeB', font_path))

        packet_page_1 = BytesIO()
        packet_page_N = BytesIO()
        
        can_1 = canvas.Canvas(packet_page_1, pagesize=letter)
        can_1.setFont('TypeB', 14)
        can_1.drawString(int(json_Product["Модификация_верх_X"]), int(json_Product["Модификация_верх_Y"]), modify_text)
        can_1.drawString(int(json_Product["Номер_заказа_верх_X"]), int(json_Product["Номер_заказа_верх_Y"]), partya_text)
        can_1.save()

        # Генерируем QR-код и сохраняем во временный файл
        qr = qrcode.QRCode(version=1, box_size=10, border=2)
        qr.add_data(shifr_part)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        # Сохраняем QR-код во временный файл
        qr_temp_path = os.path.join(tempfile.gettempdir(), "temp_qr.png")
        qr_img.save(qr_temp_path)

        custom_page_size = (800, 1200)  # Ширина 800, высота 1200
        can_N = canvas.Canvas(packet_page_N, pagesize=custom_page_size)
        can_N.setFont('TypeB', 14)
        can_N.drawString(json_Product["Модификация_низ_X"], json_Product["Модификация_низ_Y"], modify_text)
        can_N.drawString(json_Product["Серийный_номер_низ_X"], json_Product["Серийный_номер_низ_Y"], serial_num)
        can_N.drawString(json_Product["Количество_низ_X"], json_Product["Количество_низ_Y"], kol_text)
        can_N.drawString(json_Product["Дата_низ_X"], json_Product["Дата_низ_Y"], data_text)
        can_N.drawImage(qr_temp_path, json_Product["QR_низ_X"], json_Product["QR_низ_Y"], width=70, height=70)
        can_N.save()

        # Наложение изменений на PDF
        reader = PdfReader(temp_pdf)
        writer = PdfWriter()

        for i in range(len(reader.pages)):
            page = reader.pages[i]
            if i == 0:  # 1-я страница
                packet_page_1.seek(0)
                overlay = PdfReader(packet_page_1)
                page.merge_page(overlay.pages[0])
            elif i == int(json_Product["Количество_страниц_ПСИ"])-1:  # N-я страница
                packet_page_N.seek(0)
                overlay = PdfReader(packet_page_N)
                page.merge_page(overlay.pages[0])
            writer.add_page(page)

        # Сохраняем результат
        with open(temp_pdf, "wb") as f:
            writer.write(f)
        
        # Удаляем временный файл QR-кода
        try:
            os.remove(qr_temp_path)
        except:
            pass
        
        shutil.move(temp_pdf, final_pdf)

    def load_json_data(self):
        try:
            with open("products.json", "r", encoding="utf-8") as file:
                self.json_data = json.load(file)
        except FileNotFoundError:
            QMessageBox.critical(self, "Ошибка", "Файл products.json не найден!")
            return {}
        except json.JSONDecodeError:
            QMessageBox.critical(self, "Ошибка", "Некорректный формат JSON!")
            return {}

    def StatusBar(self, count):
        self.progressBar.setValue(count)
        QApplication.processEvents()
        logger.log(f'Прогресс создания: {count}%')

    def generate_clicked(self):                                                                                                     
        SelectProduct_text = self.Box_SelectProduct.currentText()                                                                       
        if SelectProduct_text == "Выберите изделие":                                                                                    
            QMessageBox.warning(None, "Ошибка", "Тип изделия не выбран")                                                                
            return      
        self.progressBar.show()
        self.StatusBar(3)                                                                                                                
        
        if self.gazprom.isChecked():                                                                                                    
            radio = 'g'
            name = "ГП"                                                                                                                 
            company = 'Газпром стройТЭК'                                                                                         
        elif self.transneft.isChecked():                                                                                                
            radio = 't' 
            name = "ТН"                                                                                                                
            if self.zakazchik.text() == '': 
                self.progressBar.hide()                                                                                            
                QMessageBox.warning(None, "Ошибка", "Отстутствует заказчик")                                                            
                return                                                                                                                  
            else:                                                                                                                       
                company = self.zakazchik.text()                                                                                                                                                                      
        else:     
            self.progressBar.hide()                                                                                                                      
            QMessageBox.warning(None, "Ошибка", "Не выбрана компания")                                                                  
            return    

        self.StatusBar(6)                                                                                                                  
                                                                                                 
        partya_text = self.partya.text()                                                                                                
        kol_text = self.kol.text()                                                                                                      
        modify_text = self.modify.text()
        modify_text = modify_text.upper()                                                                                              
        if self.auto_date_checkbox.isChecked():
            data_text = self.data.text()
        else:
            # Получаем текущую дату в формате НН.ГГГГ
            today = datetime.now()
            week_number = today.isocalendar()[1]
            year = today.year
            data_text = f"{week_number}.{year}"                                                                                                
        
        if partya_text.isdigit() == False :   
            self.progressBar.hide()                                                                 
            QMessageBox.warning(None, "Ошибка", "Не указан номер заказа")              
            return

        self.StatusBar(9)   
                                                                                                                         
        modify_text = modify_text.replace(" ", "")
        if len(modify_text) == 0:   
            self.progressBar.hide()                                                                                                    
            QMessageBox.warning(None, "Ошибка", "Поле модификации не может быть пустым")                                                
            return                                                                                                                      
                
        if modify_text.endswith("."):
            self.progressBar.hide()
            QMessageBox.warning(None, "Ошибка", "Модификацию нельзя заканчивать точкой в конце")                                                
            return
        if re.search(r'[?:*"<>|/$@]', modify_text) or re.search(r'\\', modify_text):
            self.progressBar.hide()
            QMessageBox.information(None, "Внимание", "В поле модификации есть символы \ / : *  ? < > | @ $\nв названии папки символ будет заменен на \"-\" . Остальное не поменяется")                                                
        new_modify_text = re.sub(r'[?:*"<>|/$@\\]', '-', modify_text)
        
        self.StatusBar(12) 

        if kol_text.isdigit() == False: 
            self.progressBar.hide()                                            
            QMessageBox.warning(None, "Ошибка", "Не указано количество")          
            return      

        self.StatusBar(15)                                                                                                                
        
        if not self.check_format(data_text):   
            self.progressBar.hide()                                                                                         
            QMessageBox.warning(None, "Ошибка", "Формат даты должен выглядеть в виде НН.ГГГГ и быть действующим")                       
            return                                                                                                                      
        if data_text[0] != '0' and len(data_text) == 6:                                                                                 
            data_text = '0' + data_text    

        self.StatusBar(18)   

        if SelectProduct_text not in modify_text.upper():
                ask_box = QMessageBox()
                ask_box.setWindowTitle('Вопрос')
                ask_box.setText(f'Похоже, модификация не совпадает с выбранным изделием.\nВы выбрали {SelectProduct_text} для {modify_text}. Продолжить?')
                ask_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                ask_box.setDefaultButton(QMessageBox.No)
                
                ask_box.button(QMessageBox.Yes).setText("Да")
                ask_box.button(QMessageBox.No).setText("Нет")
                
                asking = ask_box.exec_()
                
                if asking == QMessageBox.No:
                    self.progressBar.hide()
                    return
        self.StatusBar(21)                                                                                                                                                                              

        if radio == 'g':                                                                                                                
            info_base = "ГАЗПРОМ.db"                                                                                                    
            result = self.dublicate(info_base, SelectProduct_text, partya_text, modify_text, data_text)                                 
        else:                                                                                                              
            info_base = "ТРАНСНЕФТЬ.db"                                                                                                 
            result = self.dublicate(info_base, SelectProduct_text, partya_text, modify_text, data_text)                                                                 

        if result:  
            self.progressBar.hide()                                                                                                                    
            QMessageBox.warning(None, "Ошибка", "Нельзя вводить теже значения в данной вкладке")                                        
            return     

        self.StatusBar(27)                                                                                                                 

        poisk = data_text[-2:]
        vek = [0, 0, 0]
        g = 0
        databases = ['ГАЗПРОМ.db', 'ТРАНСНЕФТЬ.db', 'Другие_компании.db']
        for db in databases:
            conn = sqlite3.connect(db)
            cursor = conn.cursor()

            query = """
            SELECT MAX(CAST(SUBSTR(Серийный_номер, -4) AS INTEGER))
            FROM {}
            WHERE SUBSTR(Серийный_номер, 4, 1) = ? AND SUBSTR(Серийный_номер, 5, 1) = ?
            """.format(SelectProduct_text)

            values = (str(poisk)[0], str(poisk)[1])

            result = cursor.execute(query, values).fetchone()[0]

            if result is None:
                vek[g] = 0
            else:
                vek[g] = result
            g = g+1
            conn.close()
        max_value = max(vek)
        if max_value + int(kol_text) >= 10000:
            self.progressBar.hide()
            QMessageBox.warning(None, "Ошибка", f"Невозможно больше вносить данные в базу, так как на {9999-max_value} значении порядковые номера кончатся\nОбратитесь в тех. отдел")              
            return
        
        self.StatusBar(30)

        if name == "ГП":
            json_Product = self.json_data["Газпром"][SelectProduct_text]
        else:
            json_Product = self.json_data["Транснефть"][SelectProduct_text]
        series = json_Product["Буква_серийного_номера"]
        
        series = series + data_text[:2] + data_text[-2:]

        serial_num = []
        for i in range(1, int(kol_text)+1):
            number = i + max_value      
            if number < 10:
                serial_num.insert(i - 1, series + '000' + str(number))
            elif number >= 10 and number < 100:
                serial_num.insert(i - 1, series + '00' + str(number))
            elif number >= 100 and number < 1000:
                serial_num.insert(i - 1, series + '0' + str(number))
            else:
                serial_num.insert(i - 1, series + str(number))
        serial_num.insert(0, (serial_num[0] + '-' + serial_num[int(kol_text)-1]))

        inf = 0
        shifr_part = self.shifrator(partya_text, kol_text, SelectProduct_text, radio, inf)

        passport_path = json_Product.get("Путь_паспорта")

        if radio == 't':
            if not passport_path or not isinstance(passport_path, str) or not passport_path.strip():
                msg_box = QMessageBox()
                msg_box.setWindowTitle('Внимание')
                msg_box.setText('Путь к паспорту не указан или пуст. Пропустить создание паспорта?')
                msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                msg_box.setDefaultButton(QMessageBox.No)
                
                msg_box.button(QMessageBox.Yes).setText("Да")
                msg_box.button(QMessageBox.No).setText("Нет")
                
                response = msg_box.exec_()
                
                if response == QMessageBox.No:
                    self.progressBar.hide()
                    return
            else:
                try:
                    self.create_passport(json_Product, desktop_path, serial_num[0], new_modify_text, modify_text, shifr_part, kol_text, data_text)
                except Exception as e:
                    error_box = QMessageBox()
                    error_box.setWindowTitle('Ошибка создания паспорта')
                    error_box.setText(f'Не удалось создать паспорт: {str(e)}.\nВсе равно создать маркировки?')
                    error_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                    error_box.setDefaultButton(QMessageBox.No)
                    
                    error_box.button(QMessageBox.Yes).setText("Да")
                    error_box.button(QMessageBox.No).setText("Нет")
                    
                    answer = error_box.exec_()
                    
                    if answer == QMessageBox.No:
                        self.progressBar.hide()
                        return
            
            psi_path = json_Product.get("Путь_ПСИ")

            if not psi_path or not isinstance(psi_path, str) or not psi_path.strip():
                msg_box = QMessageBox()
                msg_box.setWindowTitle('Внимание')
                msg_box.setText('Путь к ПСИ не указан или пуст. Пропустить создание ПСИ?')
                msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                msg_box.setDefaultButton(QMessageBox.No)
                
                msg_box.button(QMessageBox.Yes).setText("Да")
                msg_box.button(QMessageBox.No).setText("Нет")
                
                response = msg_box.exec_()
                
                if response == QMessageBox.No:
                    self.progressBar.hide()
                    return
            else:
                try:
                    self.create_PSI(json_Product, desktop_path, serial_num[0], new_modify_text, modify_text, shifr_part, kol_text, data_text, partya_text)
                except Exception as e:
                    error_box = QMessageBox()
                    error_box.setWindowTitle('Ошибка создания ПСИ')
                    error_box.setText(f'Не удалось создать ПСИ: {str(e)}.\nВсе равно создать маркировки?')
                    error_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                    error_box.setDefaultButton(QMessageBox.No)
                    
                    error_box.button(QMessageBox.Yes).setText("Да")
                    error_box.button(QMessageBox.No).setText("Нет")
                    
                    answer = error_box.exec_()
                    
                    if answer == QMessageBox.No:
                        self.progressBar.hide()
                        return
                
        os.chdir(desktop_path)
        os.mkdir("Партия " + company + "_" + kol_text + "_" + new_modify_text + "_" + data_text)
        result_path = os.path.join(os.getcwd(), "Партия " + company + "_" + kol_text + "_" + new_modify_text + "_" + data_text)
        try:
            desktop_file = os.path.join(desktop_path, f"Паспорт {new_modify_text} {serial_num[0]}.pdf")
            result_file = os.path.join(result_path, f"Паспорт {new_modify_text} {serial_num[0]}.pdf")
            shutil.move(desktop_file, result_file)
            logger.log(f'Файл перемещен из {desktop_file} в {result_file}')
        except Exception as e:
            logger.log(f"не удалось переместить файл: {str(e)}")
        try:
            desktop_file = os.path.join(desktop_path, f"ПСИ {new_modify_text} {serial_num[0]}.pdf")
            result_file = os.path.join(result_path, f"ПСИ {new_modify_text} {serial_num[0]}.pdf")
            shutil.move(desktop_file, result_file)
            logger.log(f'Файл перемещен из {desktop_file} в {result_file}')
        except Exception as e:
            logger.log(f"не удалось переместить файл: {str(e)}")
        

        os.chdir(program_path)
        connection = sqlite3.connect(info_base)                                                                                         
        cursor = connection.cursor()                                                                                                    
        
        table_partya_name = "Заказы"                                                                                                    
        column_partya_name1 = "Компания"                                                                                                
        column_partya_name2 = "Номер_партии"                                                                                            
        column_partya_name3 = "Модификация"                                                                                             
        column_partya_name4 = "Количество"                                                                                              
        column_partya_name5 = "Дата_изготовления"                                                                                       
        column_partya_name6 = "Шифр"  
        column_partya_name7 = "Серийный_номер"                                                                                                  

                                                                                                                                                                               
        cursor.execute(f"INSERT INTO {table_partya_name} ({column_partya_name1}, {column_partya_name2}, {column_partya_name3}, \
            {column_partya_name4}, {column_partya_name5}, {column_partya_name6}, {column_partya_name7}) \
                VALUES (?, ?, ?, ?, ?, ?, ?)", (company, partya_text, modify_text, kol_text, data_text, shifr_part, serial_num[0]))                       
        
        connection.commit()                                                                                                             
        connection.close()                                                                                                              

        self.StatusBar(35)

        markirovki = []
        if radio == 'g':
            for i in range (1, int(kol_text)+1):  
                connection = sqlite3.connect(info_base)                                                                                     
                cursor = connection.cursor()                                                                                                
                table_name = SelectProduct_text                                                                                             
                column_name1 = "Компания"                                                                                                   
                column_name2 = "Номер_партии"                                                                                               
                column_name3 = "Модификация"                                                                                                
                column_name4 = "Порядковый_номер"                                                                                           
                column_name5 = "Дата_изготовления"                                                                                          
                column_name6 = "Шифр"      
                column_name7 = "Серийный_номер"                                                                                                  
                info = str(i) + '/' + kol_text                                                                                              
                inf = 1                                                                                                                     
                shifr_edin = self.shifrator(partya_text, i, SelectProduct_text, radio, inf)                                                 
                cursor.execute(f"INSERT INTO {table_name} ({column_name1}, {column_name2}, {column_name3}, {column_name4}, \
                    {column_name5}, {column_name6}, {column_name7}) VALUES (?, ?, ?, ?, ?, ?, ?)", \
                        (company, partya_text, modify_text, info, data_text, shifr_edin, serial_num[i]))                                                   
                connection.commit()                                                                                                         
                connection.close()     

                current_photo = ishodniki_path + "/" + SelectProduct_text + " ГСТ.png"
                image = openImg(current_photo)

                font = ImageFont.truetype("HE_CN__B.otf", 25)
                big_font = ImageFont.truetype("HE_CN__B.otf", 35)
                img_pil = Image.fromarray(image)
                draw = ImageDraw.Draw(img_pil)
                draw.text((json_Product["Модификация_Y"], json_Product["Модификация_X"]), modify_text, font=big_font, fill=(0, 0, 0, 255))
                draw.text((json_Product["Серийный_номер_Y"], json_Product["Серийный_номер_X"]), serial_num[i], font=font, fill=(0, 0, 0, 255))
                draw.text((json_Product["Дата_Y"], json_Product["Дата_X"]), data_text, font=font, fill=(0, 0, 0, 255))
                if SelectProduct_text == 'БСЗ':
                    draw.text((json_Product["Масса_Y"], json_Product["Масса_X"]), "ХЗ", font=font, fill=(0, 0, 0, 255))
                    draw.text((json_Product["Ток_Y"], json_Product["Ток_X"]), "ХЗ", font=font, fill=(0, 0, 0, 255))
                    draw.text((json_Product["Климатическое_исполнение_Y"], json_Product["Климатическое_исполнение_X"]), "ХЗ", font=font, fill=(0, 0, 0, 255))
        # добавить сюда еще надписи для бсз
                image = np.array(img_pil)
                saveImg( result_path + '/' + new_modify_text + "_" + str(i) + "_" + data_text + "_" + partya_text + ".png", image)
                markirovki.append(result_path + '/' + new_modify_text + "_" + str(i) + "_" + data_text + "_" + partya_text + ".png")
                self.StatusBar(35 + 40 - 40 // i)
        # добавить вариант прямоугольный
            if SelectProduct_text == "АЗ" or SelectProduct_text == "ПАЗ":
                os.chdir(result_path)
                width, height = A4
                # === НОВЫЕ ПАРАМЕТРЫ ДЛЯ 14×5.2 см ===
                photo_width  = 14 * cm
                photo_height = 5.2 * cm
                margin_left  = 3.5 * cm
                margin_top   = 1.0 * cm
                spacing_y    = 0.5 * cm
                marks_per_page = 5

                total_spacing = (marks_per_page - 1) * spacing_y
                total_height_needed = marks_per_page * photo_height + total_spacing
                y_start = height - margin_top - photo_height  # позиция первой маркировки

                # Сколько полных листов
                cel = int(kol_text) // marks_per_page
                ost = int(kol_text) % marks_per_page

                # === ЦЕЛЫЕ ЛИСТЫ ===
                if cel >= 1:
                    j = 0
                    for i in range(cel):
                        start_idx = j
                        end_idx = j + marks_per_page - 1
                        pdf_name = f'{serial_num[start_idx + 1]} - {serial_num[end_idx + 1]}.pdf'
                        c = canvas.Canvas(pdf_name, pagesize=A4)

                        for row in range(marks_per_page):
                            idx = j + row
                            x = margin_left
                            y = y_start - row * (photo_height + spacing_y)
                            
                            image = im(markirovki[idx], width=photo_width, height=photo_height)
                            image.drawOn(c, x, y)

                        c.save()
                        logger.log(f"Создан лист: {pdf_name}")
                        j += marks_per_page
                        self.StatusBar(75 + 20 * (i + 1) // cel)

                # === НЕПОЛНЫЙ ЛИСТ ===
                if ost > 0:
                    start_idx = cel * marks_per_page
                    pdf_name = f'{serial_num[start_idx + 1]} - {serial_num[start_idx + ost]}.pdf'
                    c = canvas.Canvas(pdf_name, pagesize=A4)

                    for row in range(ost):
                        idx = start_idx + row
                        x = margin_left
                        y = y_start - row * (photo_height + spacing_y)
                        
                        image = im(markirovki[idx], width=photo_width, height=photo_height)
                        image.drawOn(c, x, y)

                    c.save()
                    logger.log(f"Создан неполный лист: {pdf_name}")
                    self.StatusBar(100)

            else:
                os.chdir(result_path)
                width, height = A4
                photo_width = 9 * cm
                photo_height = 5.2 * cm
                x_start = 40
                y_start = height - 40

                cel = int(kol_text) // 10
                ost = int(kol_text) % 10
                logger.log(f'Количество маркировок {kol_text}, целых листов по 12 штук - {cel}, количество маркировок на неполном листе - {ost}')

                if cel >= 1:
                    j = 0
                    logger.log('Создание целого листа маркировок...')
                    for i in range(1, cel + 1):
                        c = canvas.Canvas(f'{serial_num[j + 1]} - {serial_num[j + 10]}.pdf', pagesize=A4)

                        for k in range(5):
                            image = im(markirovki[k + j], width=photo_width, height=photo_height)
                            image.drawOn(c, x_start, y_start - (k + 1) * (photo_height + 10))

                        for k in range(5, 10):
                            image = im(markirovki[k + j], width=photo_width, height=photo_height)
                            image.drawOn(c, x_start + photo_width + 10, y_start - (k - 4) * (photo_height + 10))

                        c.save()
                        logger.log(f'Создан {i} лист')

                        j = j + 10
                        self.StatusBar(75 + 18 - 18 // i)
                if ost > 5:
                    j = 0
                    logger.log(f'Создание листа с {ost} маркировоками...')
                    c = canvas.Canvas(f'{serial_num[j + 1 + cel * 10]} - {serial_num[j + ost + cel * 10]}.pdf', pagesize=A4)

                    for k in range(j + 5):
                        image = im(markirovki[k + j + cel * 10], width=photo_width, height=photo_height)
                        image.drawOn(c, x_start, y_start - (k + 1) * (photo_height + 10))

                    for k in range(j + 5, ost):
                        image = im(markirovki[k + j + cel * 10], width=photo_width, height=photo_height)
                        image.drawOn(c, x_start + photo_width + 10, y_start - ((k - 5) + 1) * (photo_height + 10))

                    c.save()
                    logger.log(f'Лист с {ost} маркировоками создан')
                    self.StatusBar(95)
                elif ost > 0 and ost <= 5:
                    logger.log(f'Создание листа с {ost} маркировоками...')
                    j = 0
                    c = canvas.Canvas(f'{serial_num[j + 1 + cel * 10]} - {serial_num[ost + cel * 10]}.pdf', pagesize=A4)

                    for k in range(ost):
                        image = im(markirovki[k + cel * 10], width=photo_width, height=photo_height)
                        image.drawOn(c, x_start, y_start - (k + 1) * (photo_height + 10))

                    c.save()
                    logger.log(f'Лист с {ost} маркировоками создан')
                    self.StatusBar(95)
        else:
            markirovka = json_Product["Маркировка"]
            if markirovka == 'прямоугольник':
                for i in range (1, int(kol_text)+1): 
                    
                    connection = sqlite3.connect(info_base)                                                                                     
                    cursor = connection.cursor()                                                                                                
                    table_name = SelectProduct_text                                                                                             
                    column_name1 = "Компания"                                                                                                   
                    column_name2 = "Номер_партии"                                                                                               
                    column_name3 = "Модификация"                                                                                                
                    column_name4 = "Порядковый_номер"                                                                                           
                    column_name5 = "Дата_изготовления"                                                                                          
                    column_name6 = "Шифр"      
                    column_name7 = "Серийный_номер"                                                                                                  
                    info = str(i) + '/' + kol_text                                                                                              
                    inf = 1                                                                                                                     
                    shifr_edin = self.shifrator(partya_text, i, SelectProduct_text, radio, inf)                                                 
                    cursor.execute(f"INSERT INTO {table_name} ({column_name1}, {column_name2}, {column_name3}, {column_name4}, \
                        {column_name5}, {column_name6}, {column_name7}) VALUES (?, ?, ?, ?, ?, ?, ?)", \
                            (company, partya_text, modify_text, info, data_text, shifr_edin, serial_num[i]))                                                   
                    connection.commit()                                                                                                         
                    connection.close()     

                    current_photo = ishodniki_path + "/" + markirovka + ".png"
                    image = openImg(current_photo)

                    font = ImageFont.truetype("GOST2304_TypeB.ttf", 45)
                    img_pil = Image.fromarray(image)
                    draw = ImageDraw.Draw(img_pil)
                    start_x1 = 270  
                    start_x2 = 270   
                    start_x3 = 263
                    end_x = 950
                    text1 = f'{json_Product["Наименование"]} {SelectProduct_text}.РСЗ'
                    text_width = font.getlength(text1)
                    center_x1 = start_x1 + (end_x - start_x1 - text_width) // 2
                    try:
                        text2 = json_Product["ТУ"] + json_Product["ГОСТ"]
                    except:
                        text2 = json_Product["ТУ"]
                    text_width = font.getlength(text2)
                    center_x2 = start_x2 + (end_x - start_x2 - text_width) // 2
                    text3 = modify_text
                    text_width = font.getlength(text3)
                    center_x3 = start_x3 + (end_x - start_x3 - text_width) // 2
                    draw.text((center_x1, 40), text1, font=font, fill=(0, 0, 0, 255))
                    draw.text((center_x2, 100), text2, font=font, fill=(0, 0, 0, 255))
                    draw.text((center_x3, 220), text3, font=font, fill=(0, 0, 0, 255))
                    draw.text((350, 295), serial_num[i], font=font, fill=(0, 0, 0, 255))
                    draw.text((730, 295), data_text, font=font, fill=(0, 0, 0, 255))
                    try:
                        start_x = 260
                        end_xx = 980
                        attention = json_Product["Предупреждение"]
                        text_width = font.getlength(attention)
                        center_x = start_x + (end_xx - start_x - text_width) // 2
                        draw.text((center_x, 150), attention, font=font, fill=(0, 0, 0, 255))
                    except:
                        pass
                    image = np.array(img_pil)
                    saveImg( result_path + '/' + new_modify_text + "_" + str(i) + "_" + data_text + "_" + partya_text + ".png", image)

                    background = Image.open(result_path + '/' + new_modify_text + "_" + str(i) + "_" + data_text + "_" + partya_text + ".png")

                    qr = qrcode.QRCode(
                        version=1,
                        error_correction=qrcode.constants.ERROR_CORRECT_L,
                        box_size=5.5,
                        border=0.5,
                    )
                    qr.add_data(shifr_edin)
                    qr.make(fit=True)
                    qr_image = qr.make_image(fill_color="black", back_color="white")

                    background.paste(qr_image, (85, 225))
                    markirovki.append(result_path + '/' + new_modify_text + "_" + str(i) + "_" + data_text + "_" + partya_text + ".png")
                    background.save(result_path + '/' + new_modify_text + "_" + str(i) + "_" + data_text + "_" + partya_text + ".png")
                    self.StatusBar(35 + 40 - 40 // i)
                os.chdir(result_path)
                width, height = A4
                photo_width = 9 * cm
                photo_height = 4 * cm
                x_start = 40
                y_start = height - 40

                cel = int(kol_text) // 12
                ost = int(kol_text) % 12
                logger.log(f'Количество маркировок {kol_text}, целых листов по 12 штук - {cel}, количество маркировок на неполном листе - {ost}')

                if cel >= 1:
                    j = 0
                    logger.log('Создание целого листа маркировок...')
                    for i in range(1, cel + 1):
                        c = canvas.Canvas(f'{serial_num[j + 1]} - {serial_num[j + 12]}.pdf', pagesize=A4)

                        for k in range(6):
                            image = im(markirovki[k + j], width=photo_width, height=photo_height)
                            image.drawOn(c, x_start, y_start - (k + 1) * (photo_height + 10))

                        for k in range(6, 12):
                            image = im(markirovki[k + j], width=photo_width, height=photo_height)
                            image.drawOn(c, x_start + photo_width + 10, y_start - (k - 5) * (photo_height + 10))

                        c.save()
                        logger.log(f'Создан {i} лист')

                        j = j + 12
                        self.StatusBar(75 + 18 - 18 // i)
                if ost > 6:
                    j = 0
                    logger.log(f'Создание листа с {ost} маркировоками...')
                    c = canvas.Canvas(f'{serial_num[j + 1 + cel * 12]} - {serial_num[j + ost + cel * 12]}.pdf', pagesize=A4)

                    for k in range(j + 6):
                        image = im(markirovki[k + j + cel * 12], width=photo_width, height=photo_height)
                        image.drawOn(c, x_start, y_start - (k + 1) * (photo_height + 10))

                    for k in range(j + 6, ost):
                        image = im(markirovki[k + j + cel * 12], width=photo_width, height=photo_height)
                        image.drawOn(c, x_start + photo_width + 10, y_start - (k - 5) * (photo_height + 10))

                    c.save()
                    logger.log(f'Лист с {ost} маркировоками создан')
                    self.StatusBar(95)
                elif ost > 0 and ost <= 6:
                    logger.log(f'Создание листа с {ost} маркировоками...')
                    j = 0
                    c = canvas.Canvas(f'{serial_num[j + 1 + + cel * 12]} - {serial_num[ost + cel * 12]}.pdf', pagesize=A4)

                    for k in range(ost):
                        image = im(markirovki[k + cel * 12], width=photo_width, height=photo_height)
                        image.drawOn(c, x_start, y_start - (k + 1) * (photo_height + 10))

                    c.save()
                    logger.log(f'Лист с {ost} маркировоками создан')
                    self.StatusBar(95)

            elif markirovka == 'квадрат':
                for i in range (1, int(kol_text)+1): 
                    
                    connection = sqlite3.connect(info_base)                                                                                     
                    cursor = connection.cursor()                                                                                                
                    table_name = SelectProduct_text                                                                                             
                    column_name1 = "Компания"                                                                                                   
                    column_name2 = "Номер_партии"                                                                                               
                    column_name3 = "Модификация"                                                                                                
                    column_name4 = "Порядковый_номер"                                                                                           
                    column_name5 = "Дата_изготовления"                                                                                          
                    column_name6 = "Шифр"      
                    column_name7 = "Серийный_номер"                                                                                                  
                    info = str(i) + '/' + kol_text                                                                                              
                    inf = 1                                                                                                                     
                    shifr_edin = self.shifrator(partya_text, i, SelectProduct_text, radio, inf)                                                 
                    cursor.execute(f"INSERT INTO {table_name} ({column_name1}, {column_name2}, {column_name3}, {column_name4}, \
                        {column_name5}, {column_name6}, {column_name7}) VALUES (?, ?, ?, ?, ?, ?, ?)", \
                            (company, partya_text, modify_text, info, data_text, shifr_edin, serial_num[i]))                                                   
                    connection.commit()                                                                                                         
                    connection.close()     

                    current_photo = ishodniki_path + "/" + markirovka + ".png"
                    image = openImg(current_photo)

                    font = ImageFont.truetype("GOST2304_TypeB.ttf", 35)
                    img_pil = Image.fromarray(image)
                    draw = ImageDraw.Draw(img_pil)
                    start_x1 = 130   
                    start_x2 = 10
                    end_x = 665
                    text1 = f'{json_Product["Наименование"]} {SelectProduct_text}.РСЗ'
                    text_width = font.getlength(text1)
                    center_x1 = start_x1 + (end_x - start_x1 - text_width) // 2
                    try:
                        text2 = json_Product["ТУ"] + '  ' +json_Product["ГОСТ"]
                    except:
                        text2 = json_Product["ТУ"]
                    text_width = font.getlength(text2)
                    center_x2 = start_x1 + (end_x - start_x1 - text_width) // 2
                    text3 = modify_text
                    text_width = font.getlength(text3)
                    center_x3 = start_x2 + (end_x - start_x2 - text_width) // 2
                    text4 = shifr_edin
                    text_width = font.getlength(text4)
                    center_x4 = start_x2 + (end_x - start_x2 - text_width) // 2
                    draw.text((center_x1, 300), text1, font=font, fill=(0, 0, 0, 255))
                    draw.text((center_x2, 350), text2, font=font, fill=(0, 0, 0, 255))
                    draw.text((205, 460), serial_num[i], font=font, fill=(0, 0, 0, 255))
                    draw.text((475, 460), data_text, font=font, fill=(0, 0, 0, 255))
                    draw.text((center_x3, 410), text3, font=font, fill=(0, 0, 0, 255))
                    draw.text((center_x4, 515), text4, font=font, fill=(0, 0, 0, 255))
                    try:
                        start_x = 10
                        end_xx = 665
                        attention = json_Product["Предупреждение"]
                        width_x = font.getlength(attention)
                        center_x = start_x + (end_xx - start_x - width_x) // 2
                        print(width_x)
                        print(center_x)
                        draw.text((center_x, 485), attention, font=font, fill=(0, 0, 0, 255))
                    except:
                        pass
                    image = np.array(img_pil)
                    saveImg( result_path + '/' + new_modify_text + "_" + str(i) + "_" + data_text + "_" + partya_text + ".png", image)
                    markirovki.append(result_path + '/' + new_modify_text + "_" + str(i) + "_" + data_text + "_" + partya_text + ".png")
                    self.StatusBar(35 + 40 - 40 // i)
                os.chdir(result_path)
                width, height = A4
                photo_width = 6 * cm
                photo_height = 5 * cm
                margin_x = 10 * mm  # отступ от края страницы
                margin_y = 15 * mm  # верхний отступ
                spacing = 3 * mm    # расстояние между маркировками

                cols = 3
                rows = 5

                # Рассчитываем общие размеры блока маркировок
                total_width = cols * photo_width + (cols - 1) * spacing
                total_height = rows * photo_height + (rows - 1) * spacing

                # Центрируем блок маркировок на странице
                x_start = (width - total_width) / 2
                y_start = height - margin_y - photo_height  # начальная позиция для первой маркировки

                cel = int(kol_text) // 15
                ost = int(kol_text) % 15
                logger.log(f'Количество маркировок {kol_text}, целых листов по 15 штук - {cel}, количество маркировок на неполном листе - {ost}')

                def draw_markings(c, markings, start_idx, count):
                    """Функция для рисования маркировок в правильном порядке"""
                    markings_added = 0
                    # Сначала левый столбец (сверху вниз)
                    for row in range(rows):
                        if markings_added >= count:
                            break
                        x = x_start
                        y = y_start - row * (photo_height + spacing)
                        image = im(markings[start_idx + markings_added], width=photo_width, height=photo_height)
                        image.drawOn(c, x, y)
                        markings_added += 1
                    
                    # Затем средний столбец (сверху вниз)
                    for row in range(rows):
                        if markings_added >= count:
                            break
                        x = x_start + photo_width + spacing
                        y = y_start - row * (photo_height + spacing)
                        image = im(markings[start_idx + markings_added], width=photo_width, height=photo_height)
                        image.drawOn(c, x, y)
                        markings_added += 1
                    
                    # Затем правый столбец (сверху вниз)
                    for row in range(rows):
                        if markings_added >= count:
                            break
                        x = x_start + 2 * (photo_width + spacing)
                        y = y_start - row * (photo_height + spacing)
                        image = im(markings[start_idx + markings_added], width=photo_width, height=photo_height)
                        image.drawOn(c, x, y)
                        markings_added += 1

                if cel >= 1:
                    j = 0
                    logger.log('Создание целых листов маркировок...')
                    for i in range(1, cel + 1):
                        c = canvas.Canvas(f'{serial_num[j + 1]} - {serial_num[j + 15]}.pdf', pagesize=A4)
                        draw_markings(c, markirovki, j, 15)
                        c.save()
                        self.StatusBar(75 + 20 - 20 // i)
                        logger.log(f'Создан {i} лист')
                        j = j + 15

                # Обработка неполного листа
                if ost > 0:
                    logger.log(f'Создание листа с {ost} маркировками...')
                    c = canvas.Canvas(f'{serial_num[cel * 15 + 1]} - {serial_num[cel * 15 + ost]}.pdf', pagesize=A4)
                    draw_markings(c, markirovki, cel * 15, ost)
                    c.save()
                    self.StatusBar(95)
                    logger.log(f'Лист с {ost} маркировками создан')
            else:
                QMessageBox.warning(None, "Ошибка", "Не указан параметр создания маркировки в products.json")              
                return
        self.StatusBar(97)
        shutil.copytree(result_path, reserv_path + '/' + "Партия " + company + "_" + kol_text + "_" + new_modify_text + "_" + data_text) 
            
        os.chdir(program_path)
        self.StatusBar(100)
        QMessageBox.information(None, "Ура", "Комплект документов готов и лежит в папке на рабочем столе!")  

        self.kol.clear()
        self.modify.clear()
        self.progressBar.hide()

    def upakovka_clicked(self):
        logger.log("Начата генерация упаковочного листа на основе шаблона")

        # Собираем данные из таблицы
        table_data = []
        for row in range(self.packaging_table.rowCount()):
            row_data = []
            for column in range(self.packaging_table.columnCount()):
                item = self.packaging_table.item(row, column)
                row_data.append(item.text() if item else "")
            if any(row_data):  # Добавляем только непустые строки
                table_data.append(row_data)

        if not table_data:  # Проверяем, есть ли данные
            QMessageBox.warning(None, "Ошибка", "Таблица пуста. Добавьте данные перед созданием упаковочного листа.")
            logger.log("Попытка создания PDF с пустой таблицей")
            return

        # Путь к шаблону и выходному файлу
        template_path = os.path.join(program_path, "Упаковочный лист (А5).pdf")
        pdf_name = f"Упаковочный_лист_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        output_path = os.path.join(desktop_path, pdf_name)

        # Проверяем наличие шаблона
        if not os.path.exists(template_path):
            QMessageBox.critical(None, "Ошибка", "Шаблон 'Упаковочный лист (А5).pdf' не найден в директории программы.")
            logger.log(f"Шаблон не найден: {template_path}")
            return

        try:
            # Регистрируем шрифт
            font_path = os.path.join(program_path, "GOST2304_TypeB.ttf")
            pdfmetrics.registerFont(TTFont('TypeB', font_path))

            # Координаты для размещения данных (замените на ваши значения)
            # Пример координат для A5 (148 × 210 мм, где 1 мм = 2.834645 точек)
            start_x_name = 20 * mm  # Координата X для столбца "Наименование"
            start_x_quantity = 100 * mm  # Координата X для столбца "Количество"
            start_y = 190 * mm  # Начальная координата Y (сверху страницы)
            line_height = 10 * mm  # Расстояние между строками

            # Создаем временный PDF с данными
            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=A5)
            can.setFont('TypeB', 12)

            # Записываем данные из таблицы
            for i, (name, quantity) in enumerate(table_data):
                y_position = start_y - i * line_height
                can.drawString(start_x_name, y_position, name)
                can.drawString(start_x_quantity, y_position, quantity)

            can.save()
            packet.seek(0)

            # Читаем шаблон и накладываем данные
            reader = PdfReader(template_path)
            writer = PdfWriter()

            # Предполагаем, что данные записываются на первой странице
            page = reader.pages[0]
            overlay = PdfReader(packet)
            page.merge_page(overlay.pages[0])
            writer.add_page(page)

            # Добавляем остальные страницы шаблона, если они есть
            for i in range(1, len(reader.pages)):
                writer.add_page(reader.pages[i])

            # Сохраняем итоговый PDF
            with open(output_path, "wb") as f:
                writer.write(f)

            logger.log(f"Упаковочный лист успешно создан: {output_path}")
            QMessageBox.information(None, "Успех", f"Упаковочный лист создан: {pdf_name}")

            # Очищаем таблицу после успешного создания PDF
            self.packaging_table.setRowCount(1)
            for col in range(self.packaging_table.columnCount()):
                self.packaging_table.setItem(0, col, QtWidgets.QTableWidgetItem(""))

        except Exception as e:
            logger.log(f"Ошибка при создании упаковочного листа: {str(e)}")
            QMessageBox.critical(None, "Ошибка", f"Не удалось создать упаковочный лист: {str(e)}")

    def regenerate_clicked(self):
        self.VigruzButton.hide()                                                                                                 
        nomer = self.code.text()                                                                                                        
        self.text_browser.clear() 

        if nomer == '':
            self.text_browser.append(f"Ничего не написано")                                                                 
            self.text_browser.update()
            return

        if nomer[0].isdigit() == False:
            if nomer[0] == 'К':
                izdelie = 'КИП'
            elif nomer[0] == "Б":                                                                                                   
                izdelie = "БСЗ"                                                                                                                                                                           
            elif nomer[0] == "А":                                                                                                   
                izdelie = "АЗ"                                                                                                                                                                              
            elif nomer[0] == "П":                                                                                                   
                izdelie = "ПАЗ"                                                                                                                                                                         
            elif nomer[0] == "М":                                                                                                   
                izdelie = "МСЭС"   

            databases = ['ГАЗПРОМ.db', 'ТРАНСНЕФТЬ.db', 'Другие_компании.db']

            if len(nomer) > 11:
                if nomer[9] == '-':
                    for db in databases:
                        conn = sqlite3.connect(db)
                        cursor = conn.cursor()
                        cursor.execute("SELECT Компания, Номер_партии, Модификация, Количество, Дата_изготовления, Шифр\
                            FROM Заказы WHERE Серийный_номер = ?", (nomer,))                                                                              
                        results = cursor.fetchall()
                        for row in results:                                                                                                         
                            self.text_browser.append(f"Компания: {row[0]}\nНомер партии: {row[1]}\nМодификация: {row[2]}\
                                \nКоличество изделий: {row[3]}\nДата производства: {row[4]}\nШифр: {row[5]}")                                                         
                            self.text_browser.update()    
                    conn.close()
                    self.VigruzButton.show()
                if not self.text_browser.toPlainText():
                    self.text_browser.append(f"Такого кода нет. Косяк")                                                                 
                    self.text_browser.update() 
            else:
                for db in databases:
                    conn = sqlite3.connect(db)
                    cursor = conn.cursor()
                    cursor.execute("SELECT Компания, Номер_партии, Модификация, Порядковый_номер, Дата_изготовления, Шифр \
                        FROM {} WHERE Серийный_номер = ?".format(izdelie), (nomer,))                                                                          
                    results = cursor.fetchall()                                                                                                       
                    for row in results:                                                                                                         
                        self.text_browser.append(f"Компания: {row[0]}\nНомер партии: {row[1]}\nМодификация: {row[2]}\
                            \nПорядковый номер: {row[3]}\nДата производства: {row[4]}\nШифр: {row[5]}")                                                         
                        self.text_browser.update()    
                    conn.close()
                    self.VigruzButton.show()
                if not self.text_browser.toPlainText():
                    self.text_browser.append(f"Такого кода нет. Косяк")                                                                 
                    self.text_browser.update() 
                                                                        
        elif nomer[0] == '4':                                                                                                             
            info_base = "ТРАНСНЕФТЬ.db"                                                                                                 
            text = "Транснефть - "                                                                                                      
            if nomer[2] == '0':                                                                                                         
                connect = sqlite3.connect("ТРАНСНЕФТЬ.db")                                                                              
                cursor = connect.cursor()                                                                                               
                cursor.execute("SELECT Компания, Номер_партии, Модификация, Количество, Дата_изготовления, Серийный_номер\
                    FROM Заказы WHERE Шифр = ?", (nomer,))                                                                              
                results = cursor.fetchall()                                                                                             
                if len(results) == 0:                                                                                                   
                    self.text_browser.append(f"Такого кода нет. Косяк")                                                                 
                    self.text_browser.update()                                                                                          
                else:                                                                                                                   
                    for row in results:                                                                                                 
                        self.text_browser.append(f"Компания: {text}{row[0]}\nНомер партии: {row[1]}\nМодификация: {row[2]}\
                            \nКоличество изделий: {row[3]}\nДата производства: {row[4]}\nСерийные номера: {row[5]}")                                               
                        self.text_browser.update()   
                        self.VigruzButton.show()                                                                                 
                connect.close()                                                                                                         
            elif nomer[2] == '1':                                                                                                       
                if nomer[3] == "1":                                                                                                     
                    izdelie = "КИП"                                                                                                     
                    self.zapros(info_base, izdelie, nomer, text)
                    self.VigruzButton.show()                                                                      
                elif nomer[3] == "2":                                                                                                   
                    izdelie = "БСЗ"                                                                                                     
                    self.zapros(info_base, izdelie, nomer, text)
                    self.VigruzButton.show()                                                                   
                elif nomer[3] == "3":                                                                                                   
                    izdelie = "АЗ"                                                                                                      
                    self.zapros(info_base, izdelie, nomer, text)
                    self.VigruzButton.show()                                                                       
                elif nomer[3] == "4":                                                                                                   
                    izdelie = "ПАЗ"                                                                                                     
                    self.zapros(info_base, izdelie, nomer, text)
                    self.VigruzButton.show()                                                                       
                elif nomer[3] == "5":                                                                                                   
                    izdelie = "МСЭС"                                                                                                    
                    self.zapros(info_base, izdelie, nomer, text)
                    self.VigruzButton.show()                                                                       
                else:                                                                                                                   
                    self.text_browser.append("Четвертая цифра не верна")                                                                
                    self.text_browser.update()                                                                                          
            else:                                                                                                                       
                self.text_browser.append("Третья цифра не верна")                                                                       
                self.text_browser.update()                                                                                              
        elif nomer[0] == '8':                                                                                                           
            info_base = "ГАЗПРОМ.db"                                                                                                    
            text = "Газпром - "                                                                                                         
            if nomer[2] == '0':                                                                                                         
                connect = sqlite3.connect("ГАЗПРОМ.db")                                                                                 
                cursor = connect.cursor()                                                                                               
                cursor.execute("SELECT Компания, Номер_партии, Модификация, Количество, \
                    Дата_изготовления, Серийный_номер FROM Заказы WHERE Шифр = ?", (nomer,))                                                            
                results = cursor.fetchall()                                                                                             
                if len(results) == 0:                                                                                                   
                    self.text_browser.append(f"Такого кода нет. Косяк")                                                                 
                    self.text_browser.update()                                                                                          
                else:                                                                                                                   
                    for row in results:                                                                                                 
                        self.text_browser.append(f"Компания: {text}{row[0]}\nНомер партии: {row[1]}\nМодификация: {row[2]}\
                            \nКоличество изделий: {row[3]}\nДата производства: {row[4]}\nСерийные номера: {row[5]}")                                               
                        self.text_browser.update()                                                                                      
                connect.close()    
                self.VigruzButton.show()                                                                                                    
            elif nomer[2] == '1':                                                                                                       
                if nomer[3] == "1":                                                                                                     
                    izdelie = "КИП"                                                                                                     
                    self.zapros(info_base, izdelie, nomer, text)  
                    self.VigruzButton.show()                                                                     
                elif nomer[3] == "2":                                                                                                   
                    izdelie = "БСЗ"                                                                                                     
                    self.zapros(info_base, izdelie, nomer, text)
                    self.VigruzButton.show()                                                                      
                elif nomer[3] == "3":                                                                                                   
                    izdelie = "АЗ"                                                                                                      
                    self.zapros(info_base, izdelie, nomer, text)
                    self.VigruzButton.show()                                                                       
                elif nomer[3] == "4":                                                                                                   
                    izdelie = "ПАЗ"                                                                                                     
                    self.zapros(info_base, izdelie, nomer, text)
                    self.VigruzButton.show()                                                                       
                elif nomer[3] == "5":                                                                                                   
                    izdelie = "МСЭС"                                                                                                    
                    self.zapros(info_base, izdelie, nomer, text)
                    self.VigruzButton.show()                                                                      
                else:                                                                                                                   
                    self.text_browser.append("Четвертая цифра не верна")                                                                
                    self.text_browser.update()                                                                                          
            else:                                                                                                                       
                self.text_browser.append("Третья цифра не верна")                                                                       
                self.text_browser.update()                                                                                              
        elif nomer[0] == '6':                                                                                                           
            info_base = "Другие_компании.db"                                                                                            
            text = ""                                                                                                                   
            if nomer[2] == '0':                                                                                                         
                connect = sqlite3.connect("Другие_компании.db")                                                                         
                cursor = connect.cursor()                                                                                               
                cursor.execute("SELECT Компания, Номер_партии, Модификация, Количество, Дата_изготовления, Серийный_номер \
                    FROM Заказы WHERE Шифр = ?", (nomer,))                                                                              
                results = cursor.fetchall()                                                                                             
                if len(results) == 0:                                                                                                   
                    self.text_browser.append(f"Такого кода нет. Косяк")                                                                 
                    self.text_browser.update()                                                                                          
                else:                                                                                                                   
                    for row in results:                                                                                                 
                        self.text_browser.append(f"Компания: {text}{row[0]}\nНомер партии: {row[1]}\nМодификация: {row[2]}\
                            \nКоличество изделий: {row[3]}\nДата производства: {row[4]}\nСерийные номера: {row[5]}")                                               
                        self.text_browser.update()                                                                                    
                connect.close()        
                self.VigruzButton.show()                                                                                               
            elif nomer[2] == '1':                                                                                                       
                if nomer[3] == "1":                                                                                                     
                    izdelie = "КИП"                                                                                                     
                    self.zapros(info_base, izdelie, nomer, text) 
                    self.VigruzButton.show()                                                                      
                elif nomer[3] == "2":                                                                                                   
                    izdelie = "БСЗ"                                                                                                     
                    self.zapros(info_base, izdelie, nomer, text)     
                    self.VigruzButton.show()                                                                  
                elif nomer[3] == "3":                                                                                                   
                    izdelie = "АЗ"                                                                                                      
                    self.zapros(info_base, izdelie, nomer, text) 
                    self.VigruzButton.show()                                                                     
                elif nomer[3] == "4":                                                                                                   
                    izdelie = "ПАЗ"                                                                                                     
                    self.zapros(info_base, izdelie, nomer, text)   
                    self.VigruzButton.show()                                                                   
                elif nomer[3] == "5":                                                                                                   
                    izdelie = "МСЭС"                                                                                                    
                    self.zapros(info_base, izdelie, nomer, text)       
                    self.VigruzButton.show()                                                               
                else:                                                                                                                   
                    self.text_browser.append("Четвертая цифра не верна")                                                                
                    self.text_browser.update()                                                                                          
            else:                                                                                                                       
                self.text_browser.append("Третья цифра не верна")                                                                       
                self.text_browser.update()                                                                                              
        else:                                                                                                                           
            self.text_browser.append("Такого номера нет")                                                                       
            self.text_browser.update()                                                                                                     

    def vigruzka_clicked(self):

        text = self.text_browser.toPlainText()

        mod_pattern = re.compile(r'Модификация:\s*(.*?)\s*\n')
        kol_pattern = re.compile(r'Количество изделий:\s*(.*?)\s*\n')
        num_pattern = re.compile(r'Порядковый номер:\s*(\d+)/\d+')
        date_pattern = re.compile(r'Дата производства:\s*(.*?)\s*\n')

        mod_match = mod_pattern.search(text)

        kol_match = kol_pattern.search(text)
        if not kol_match:
            kol_match = 1

        num_match = num_pattern.search(text)

        if mod_match and re.search(r'[?:*"<>|/$@\\]', mod_match.group(1)):
            # Создаем новую строку с замененными символами
            cleaned_num = re.sub(r'[?:*"<>|/$@\\]', '-', mod_match.group(1))
            
            # Заменяем в исходном тексте
            new_text = text.replace(mod_match.group(1), cleaned_num)
            self.text_browser.setPlainText(new_text)
            
            # Обновляем match объект с новым текстом
            mod_match = mod_pattern.search(new_text) 

        date_match = date_pattern.search(text)

        os.chdir(desktop_path)
        os.mkdir('Выгрузка ' + mod_match.group(1))
        os.chdir('Выгрузка ' + mod_match.group(1))

        with open('результат поиска.txt', 'w') as file:
            file.write(text)

        source_directory = reserv_path        
        destination_directory = desktop_path + '/' + 'Выгрузка ' + mod_match.group(1) 

        if kol_match == 1:
            search_name = mod_match.group(1) + "_" + num_match.group(1) + "_" + date_match.group(1) 
            found = False 
            for root, dirs, files in os.walk(source_directory):
                for file_name in files:
                    if search_name in file_name:
                        source_path = os.path.join(root, file_name)
                        dest_path = os.path.join(destination_directory, file_name)
                        shutil.copyfile(source_path, dest_path)
                        found = True
            if not found:
                QMessageBox.warning(None, "Увы", "Этикетки не найдены") 
            else:
                QMessageBox.information(None, "Отличто!", "Данные выгружены в папку на рабочий стол")

        else:
            search_name = kol_match.group(1) + '_' + mod_match.group(1) + '_' + date_match.group(1)
            found = False
            for root, dirs, files in os.walk(reserv_path):
                for dir_name in dirs:
                    if search_name in dir_name:
                        source_path = os.path.join(root, dir_name)
                        dest_path = os.path.join(destination_directory, dir_name)
                        shutil.copytree(source_path, dest_path)
                        found = True
            if not found:
                QMessageBox.warning(None, "Увы", "Этикетки не найдены")
            else:
                QMessageBox.information(None, "Отлично!", "Данные выгружены в папку на рабочий стол")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    if not Path(ishodniki_path).exists() or not Path(reserv_path).exists() or not Path('ГАЗПРОМ.db').exists() or not Path('ТРАНСНЕФТЬ.db').exists() or not Path('Другие_компании.db').exists() or not Path('Logo.svg').exists() or not Path('Instrukcia.pdf').exists or not Path(ishodniki_path + '/' + "Квадрат.png").exists() or not Path(ishodniki_path + '/' + "Прямоугольник.png").exists() or not Path(ishodniki_path + '/' + "style.css"):
        QMessageBox.critical(None, "Ошибка", "Потерян один из исходных файлов. Программа не будет запущена")
        logger.log("Выведена ошибка. Потерян исходный файл. Программа не будет запущена")
        app.quit() 
    else: 
        try:
            with open("style_markirator.css", 'r') as file:
                app.setStyleSheet(file.read())
        except FileNotFoundError:
            logger.log("Ошибка: Файл style_markirator.css не найден!")
        except PermissionError:
            logger.log("Ошибка: Нет доступа к файлу style_markirator.css!")
        except Exception as e:
            logger.log(f"Неизвестная ошибка при загрузке CSS: {e}")                                                                                           
        MainWindow = QtWidgets.QMainWindow()                                                                                                
        ui = Ui_MainWindow()                                                                                                                
        ui.setupUi(MainWindow)                                                                                                              
        MainWindow.show()                                                                                                                   
        sys.exit(app.exec())   


#1659 строк было       
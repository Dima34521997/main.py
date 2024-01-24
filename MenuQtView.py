from PySide6 import QtWidgets
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QPushButton,
                               QLabel,
                               QGridLayout,
                               QLineEdit,
                               QTextEdit,
                               QFileDialog)


# region Наши импорты
import Editors as ed

from Commands import ListOfElements
from Commands import Specification
from Commands import PurchasedItems

from DataModel import *
import DataModel as dm
# endregion

data = 1


class MainWindow(QtWidgets.QWidget):
    dm: DataModel

    def __init__(self, dm: DataModel):
        super().__init__()
        self.setWindowTitle('Генератор документов')
        self.dm = dm

        # region Создание кнопок и привязка к ним функций-обработчиков
        self.button_P = QPushButton('Перечень элементов')
        self.button_P.setStyleSheet('QPushButton {background-color: #fad2d5; '
                                    'border-color: gray; border style: outset;'
                                    'border-width: 2px; border-radius: 7px;'
                                    'border-color: beige; font: normal 15px;'
                                    'min-width: 14em; padding: 13px;}')

        self.button_P.setMaximumWidth(400)
        self.button_P.setMinimumWidth(300)
        #self.button_P.setStyleSheet('QPushButton {border-color: white}')
        self.button_P.clicked.connect(self.button_P_clicked)

        self.button_SP = QPushButton('Спецификация')
        self.button_SP.setStyleSheet('QPushButton {background-color: #fad2d9; '
                                     'border-color: gray; border style: outset;'
                                     'border-width: 2px; border-radius: apx;'
                                     'border-color: beige; font: normal 15px;'
                                     'min-width: 14em; padding: 13px;}')

        self.button_SP.setMaximumWidth(400)
        self.button_SP.setMinimumWidth(300)
        self.button_SP.clicked.connect(self.button_SP_clicked)

        self.button_VP = QPushButton('Покупная ведомость')
        self.button_VP.setStyleSheet('QPushButton {background-color: #fad2df; '
                                     'border-color: gray; border style: outset;'
                                     'border-width: 2px; border-radius: 7px;'
                                     'border-color: beige; font: normal 15px;'
                                     'min-width: 14em; padding: 13px;}')
        self.button_VP.setMaximumWidth(200)
        self.button_VP.setMinimumWidth(300)
        self.button_VP.clicked.connect(self.button_VP_clicked)

        self.button_Path = QPushButton('Путь к шаблонам')
        self.button_Path.clicked.connect(self.path_to_templates)

        self.button_SelectBOM = QPushButton('Выбрать BOM')
        self.button_SelectBOM.setStyleSheet('QPushButton {background-color: #9fede9; '
                                            'border-color: #d2faee; border style: solid;'
                                            'border-width: 6px; border-radius: 7px;'
                                            'border-color: beige; font: normal 14px;'
                                            'min-width: 14em; padding: 15px;}')
        self.button_SelectBOM.clicked.connect(self.select_BOM)

        self.button_Clear = QPushButton('Очистить')
        self.button_Clear.setStyleSheet('QPushButton {background-color: #9fede9; '
                                        'border-color: gray; border style: solid;'
                                        'border-width: 2px; border-radius: 7px;'
                                        'border-color: beige; font: normal 14px;'
                                        'min-width: 14em; padding: 15px;}')
        self.button_Clear.clicked.connect(self.button_Clear_clicked)
        # endregion

        # region Создание лейблов
        self.label_Developed = QLabel('Разработал')

        self.label_ProjectName = QLabel('Название проекта')
        self.label_PathToTemplates = QLabel('Путь к шаблонам')
        self.label_Verify = QLabel('Проверил')
        self.label_Control = QLabel('Нормоконтроль')
        self.label_Approve = QLabel('Утвердил')
        # endregion

        # region Создание текстбоксов
        self.textbox_Developed = QLineEdit('разработал...', alignment=Qt.AlignCenter)
        self.textbox_Developed.setStyleSheet('QLineEdit {color: gray;}')

        self.textbox_ProjectName = QLineEdit("название проекта...", alignment=Qt.AlignCenter)
        self.textbox_ProjectName.setStyleSheet('QLineEdit {color: gray;}')

        self.textbox_PathToTemplates = QLineEdit(" путь к шаблонам...", alignment=Qt.AlignCenter)
        self.textbox_PathToTemplates.setStyleSheet('QLineEdit {color: gray;}')

        self.textbox_Verify = QLineEdit("проверил...", alignment=Qt.AlignCenter)
        self.textbox_Verify.setStyleSheet('QLineEdit {color: gray;}')

        self.textbox_Control = QLineEdit("нормоконтроль...", alignment=Qt.AlignCenter)
        self.textbox_Control.setStyleSheet('QLineEdit {color: gray;}')

        self.textbox_Approve = QLineEdit("утвердил...", alignment=Qt.AlignCenter)
        self.textbox_Approve.setStyleSheet('QLineEdit {color: gray;}')

        self.textbox_AddedBOMs = QTextEdit(self)
        self.textbox_AddedBOMs.setMinimumWidth(400)

        # endregion

        self.UI()

    def UI(self):
        # region Макет окна сеточного типа
        grid = QGridLayout()
        self.setLayout(grid)
        # endregion

        # region Добавление элементов в сетку окна

        # Кнопки
        grid.addWidget(self.button_P, 2, 1, alignment=Qt.AlignCenter)
        grid.addWidget(self.button_SP, 3, 1, alignment=Qt.AlignCenter)
        grid.addWidget(self.button_VP, 4, 1, alignment=Qt.AlignCenter)
        grid.addWidget(self.button_Path, 7, 0)
        grid.addWidget(self.button_SelectBOM, 0, 0)
        grid.addWidget(self.button_Clear, 6, 0)

        # Лейблы
        grid.addWidget(self.label_ProjectName, 8, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.label_Developed, 9, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.label_Verify, 10, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.label_Control, 11, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.label_Approve, 12, 0, alignment=Qt.AlignCenter)

        # Текстбоксы
        grid.addWidget(self.textbox_PathToTemplates, 7, 1)
        grid.addWidget(self.textbox_ProjectName, 8, 1)
        grid.addWidget(self.textbox_Developed, 9, 1)
        grid.addWidget(self.textbox_Verify, 10, 1)
        grid.addWidget(self.textbox_Control, 11, 1)
        grid.addWidget(self.textbox_Approve, 12, 1)
        grid.addWidget(self.textbox_AddedBOMs, 1, 0, 5, 1)

        # endregion

    # region Обработка кнопок

    def parse_labels(self):
        '''Забирает из текстбоксов ввведенные значения'''
        self.dm.data_for_first_page['ProjectName'] = self.textbox_ProjectName.text()
        self.dm.data_for_first_page['Dev'] = self.textbox_Developed.text()
        self.dm.data_for_first_page['Verify'] = self.textbox_Verify.text()
        self.dm.data_for_first_page['NControl'] = self.textbox_Control.text()
        self.dm.data_for_first_page['Approve'] = self.textbox_Approve.text()

    def button_P_clicked(self):
        """Создать перечень элементов"""

        if self.dm.files:
            self.parse_labels()
            ListOfElements.Execute(self.dm)
        else:
            print('Ничего не выбрано!')

    def button_SP_clicked(self):
        """Создать cпецификацию"""

        if self.dm.files:
            self.parse_labels()
            Specification.Execute(self.dm)
        else:
            print('Ничего не выбрано!')
    def button_VP_clicked(self):
        print('Нажата')

    def path_to_templates(self):
        self.textbox_PathToTemplates.clear()
        selected_path = ed.path_for_win(QFileDialog.getExistingDirectory(self))
        self.textbox_PathToTemplates.insert(selected_path)
        self.dm.TemplatesPath = selected_path


    def select_BOM(self):
        """Открывает диалоговое окно с файловой системой для выбора BOM"""

        selected_files: list = QFileDialog.getOpenFileNames(self, filter='*.xls')[0]

        for file in selected_files:
            if file and ed.path_for_win(file) not in self.dm.files:
                self.dm.files.append(ed.path_for_win(file))
                self.textbox_AddedBOMs.append(f'{self.dm.counter}) {ed.only_name(file)}')
                self.dm.counter += 1
            else:
                print(f'BOM {ed.only_name(file)} уже добавлен!')
        print(f'selected_files {selected_files}')
        print(f'files_to_open {self.dm.files}')

    def button_Clear_clicked(self):
        self.textbox_AddedBOMs.clear()
        self.dm.counter = 1
        self.dm.files.clear()

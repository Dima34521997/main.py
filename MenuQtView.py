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
        self.button_P.clicked.connect(self.button_P_clicked)

        self.button_SP = QPushButton('Спецификация')
        self.button_SP.clicked.connect(self.button_SP_clicked)

        self.button_VP = QPushButton('Покупная ведомость')
        self.button_VP.clicked.connect(self.button_VP_clicked)

        self.button_Path = QPushButton('Путь к шаблонам')
        self.button_Path.clicked.connect(self.path_to_templates)

        self.button_SelectBOM = QPushButton('Выбрать BOM')
        self.button_SelectBOM.clicked.connect(self.select_BOM)

        self.button_Clear = QPushButton('Очистить')
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
        self.textbox_Developed = QLineEdit('разработал...')
        self.textbox_ProjectName = QLineEdit("название проекта...")
        self.textbox_PathToTemplates = QLineEdit(" путь к шаблонам...")
        self.textbox_Verify = QLineEdit("проверил...")
        self.textbox_Control = QLineEdit("нормоконтроль...", alignment=Qt.AlignCenter)
        self.textbox_Approve = QLineEdit("утвердил...")
        self.textbox_AddedBOMs = QTextEdit(self)
        # endregion

        self.UI()

    def UI(self):
        # region Макет окна сеточного типа
        grid = QGridLayout()
        self.setLayout(grid)
        # endregion

        # region Добавление элементов в сетку окна

        # Кнопки
        grid.addWidget(self.button_P, 2, 1)
        grid.addWidget(self.button_SP, 3, 1)
        grid.addWidget(self.button_VP, 4, 1)
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

    def data_change(self):
        '''
        Текст из соответствующих текстбоксов вставляет в значения
        словаря data
        '''
        self.dm.data['Project_Name'] = 'TEST'
        self.dm.data['Templates_Path'] = self.textbox_PathToTemplates.copy()

    # region Обработка кнопок
    def button_P_clicked(self):
        """Создать перечень элементов"""

        if self.dm.files:
            ListOfElements.Execute(self.dm)
        else:
            pass

    def button_SP_clicked(self):
        print('Нажата')

    def button_VP_clicked(self):
        print('Нажата')

    def path_to_templates(self):
        self.textbox_PathToTemplates.clear()
        selected_path = ed.path_for_win(QFileDialog.getExistingDirectory(self))
        self.textbox_PathToTemplates.insert(selected_path)
        # self.data_change()
        self.dm.data['Templates_Path'] = selected_path
        print(f'selected_path {selected_path}')
        print(self.dm.data)

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

    # @staticmethod
    # def onModified(event):
    #     data['Templates_Path'] = 'C:\\Users\\EDN\\PycharmProjects\\mainDocsUI\\Шаблоны\\\\Шаблон ПЭ.docx'
    #     data['Project_Name'] = 'ХРЮЧЕВО'
    #
    #     data['Razrab'] = 'МСЬЕ ИНДУСИО'
    #     data['Proveril'] = 'KEBAB'
    #     data['N_control'] = 'ПОЛИНА'
    #     data['Utverdil'] = 'КТО-НИБУДЬ'
    #
    #     with open("Profile.json", "w") as f:
    #         json.dump(data, f, ensure_ascii=False, indent=4)

    # endregion

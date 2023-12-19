import locale
import sys

from MenuQtView import *

locale.setlocale(locale.LC_ALL, ('ru_RU', 'UTF-8'))

if __name__ == "__main__":
    dm = DataModel()

    app = QtWidgets.QApplication([])

    mainWindow = MainWindow(dm)
    mainWindow.resize(650, 400)
    mainWindow.show()

    sys.exit(app.exec())

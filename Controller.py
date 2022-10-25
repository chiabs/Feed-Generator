from ExcelReader import Reader
from GUI import MainWindow
from PyQt5.QtWidgets import QApplication
import sys

class Controller():
    def __init__(self):
        app = QApplication(sys.argv)
        app.setStyle('Fusion')
        Reader.create_file()
        window = MainWindow()
        window.show()
        app.exec()







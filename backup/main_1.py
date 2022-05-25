from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QApplication, QMainWindow, QComboBox

import sys

class Window(QMainWindow):
    def __init__(self):
        super(Window, self).__init__()

        self.setWindowTitle('Первая программа для расписания')
        self.setGeometry(300, 250, 400, 300)

        self.new_text = QtWidgets.QLabel(self)

        self.main_text = QtWidgets.QLabel(self)
        self.main_text.setText('Это лейбл')
        self.main_text.move(100, 75)
        self.main_text.adjustSize()

        self.btn = QtWidgets.QPushButton(self)
        self.btn.setText('Кноппппка')
        self.btn.move(70, 125)
        self.btn.setFixedWidth(200)
        self.btn.clicked.connect(self.add_label)

    def add_label(self):
        self.new_text.setText("ОГО!!!")
        self.new_text.move(200, 200)
        self.new_text.adjustSize()

def application():
    app = QApplication(sys.argv)
    window = Window()

    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    application()
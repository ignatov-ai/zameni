import sys

from PyQt6.QtWidgets import (
    QApplication,
    QLabel,
    QMainWindow,
    QPushButton,
    QTabWidget,
    QWidget, QLineEdit, QComboBox, QDateEdit, QHBoxLayout, QVBoxLayout,
)

from PyQt6.QtCore import QDate

sotrudniki = []
with open('sotrudniki.csv','r') as url:
  for line in url:
    sotrudniki.append(line.strip().split(';'))

sotrudniki_fio = []
for i in range(len(sotrudniki)):
    s = sotrudniki[i][3]+'. '+sotrudniki[i][0]+' '+sotrudniki[i][1]+' '+sotrudniki[i][2]
    sotrudniki_fio.append(s)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Вкладки для замен")

        self.tabs = QTabWidget()
        tab1 = QWidget()
        tab2 = QWidget()
        tab3 = QWidget()
        self.tabs.setMovable(True)
        self.tabs.addTab(tab1, 'Создание больничного листа')
        self.tabs.addTab(tab2, '2')
        self.tabs.addTab(tab3, '3')
        self.setCentralWidget(self.tabs)

        # поле поиска и выбора учителя
        lbl_find = QLabel(tab1)
        lbl_find.setText('Поиск: ')
        lbl_find.move(30,60)
        lbl_find.setFixedWidth(70)

        teach_find = QLineEdit(tab1)
        teach_find.textChanged.connect(self.teachFind)
        teach_find.move(100,60)
        teach_find.setFixedWidth(200)

        self.teach_select = QComboBox(tab1)
        self.teach_select.move(30,100)
        self.teach_select.addItems(sotrudniki_fio)
        self.teach_select.setFixedWidth(270)
        self.teach_select.currentIndex()

        # выбор даты начала замены
        zamena_dateStart = QDateEdit(tab1, calendarPopup=True)
        zamena_dateStart.move(100,140)
        zamena_dateStart.setFixedWidth(200)
        zamena_dateStart.setDate(QDate.currentDate())

        lbl_zamenaStart = QLabel(tab1)
        lbl_zamenaStart.setText('Начало: ')
        lbl_zamenaStart.move(30,142)
        lbl_zamenaStart.setFixedWidth(70)

        # выбор даты окончания замены
        zamena_dateEnd = QDateEdit(tab1, calendarPopup=True)
        zamena_dateEnd.move(100,180)
        zamena_dateEnd.setFixedWidth(200)
        zamena_dateEnd.setDate(QDate.currentDate())

        lbl_zamenaEnd = QLabel(tab1)
        lbl_zamenaEnd.move(30,182)
        lbl_zamenaEnd.setText('Окончание: ')
        lbl_zamenaEnd.setFixedWidth(70)

        # кнопки отмены и добавления замены в журнал
        okButton = QPushButton(tab1)
        okButton.setText("Добавить")
        okButton.move(30,220)
        okButton.setFixedWidth(100)
        okButton.clicked.connect(self.zamena_add)
        #okButton.addAction(self.zamena_add)
        cancelButton = QPushButton(tab1)
        cancelButton.setText("Назад")
        cancelButton.move(200,220)
        cancelButton.setFixedWidth(100)

        hBtn = QHBoxLayout()
        hBtn.addWidget(okButton)
        hBtn.addWidget(cancelButton)

    # вывод в выпадающий список учителей по фильтру из поля поиска
    def teachFind(self, text):
        fioSrez = text.capitalize()
        lenFioSrez = len(fioSrez)
        sotrudniki_find = []
        for i in range(len(sotrudniki_fio)):
            if sotrudniki_fio[i][6:6+lenFioSrez] == fioSrez:
                sotrudniki_find.append(sotrudniki_fio[i])
        if len(sotrudniki_find) > 0:
            self.teach_select.clear()
            self.teach_select.addItems(sotrudniki_find)
        else:
            self.teach_select.addItems(sotrudniki_fio)

    def zamena_add(self):
        fio_text = self.teach_select.currentText()
        id = fio_text[:4]
        for i in range (len(sotrudniki)):
            if sotrudniki[i][3] == id:
                print(i)
                break
        print(sotrudniki[i])


app = QApplication(sys.argv)

window = MainWindow()
window.setGeometry(400, 200, 800, 600)
window.show()

app.exec()
import sys
from PyQt6.QtWidgets import (
    QMainWindow, QApplication, QPushButton,
    QLabel, QComboBox, QListView, QLineEdit, QDateTimeEdit, QDateEdit,
    QLineEdit, QVBoxLayout, QWidget, QHBoxLayout
)
from PyQt6.QtCore import QDate

sotrudniki = []
with open('../sotrudniki.csv', 'r') as url:
  for line in url:
    sotrudniki.append(line.strip().split(';'))

sotrudniki_fio = []
for i in range(len(sotrudniki)):
    s = sotrudniki[i][0]+' '+sotrudniki[i][1]+' '+sotrudniki[i][2]
    sotrudniki_fio.append(s)
print(sotrudniki_fio)

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        teach_find = QLineEdit()
        teach_find.textChanged.connect(self.teachFind)

        self.teach_select = QComboBox()
        self.teach_select.addItems(sotrudniki_fio)

        # выбор даты начала замены
        zamena_dateStart = QDateEdit(calendarPopup=True)
        zamena_dateStart.setDate(QDate.currentDate())
        zamena_dateStart.setObjectName("zamena_dateStart")

        lbl_zamenaStart = QLabel()
        lbl_zamenaStart.setText('Начало: ')
        lbl_zamenaStart.setFixedWidth(70)

        hZamenaStart = QHBoxLayout()
        hZamenaStart.addWidget(lbl_zamenaStart)
        hZamenaStart.addWidget(zamena_dateStart)

        # выбор даты окончания замены
        zamena_dateEnd = QDateEdit(calendarPopup=True)
        zamena_dateEnd.setDate(QDate.currentDate())
        zamena_dateEnd.setObjectName("zamena_dateStart")

        lbl_zamenaEnd = QLabel()
        lbl_zamenaEnd.setText('Окончание: ')
        lbl_zamenaEnd.setFixedWidth(70)

        hZamenaEnd = QHBoxLayout()
        hZamenaEnd.addWidget(lbl_zamenaEnd)
        hZamenaEnd.addWidget(zamena_dateEnd)

        lbl_find = QLabel()
        lbl_find.setText('Поиск: ')
        lbl_find.setFixedWidth(70)

        hTeachFind = QHBoxLayout()
        hTeachFind.addWidget(lbl_find)
        hTeachFind.addWidget(teach_find)

        okButton = QPushButton("Добавить")
        cancelButton = QPushButton("Отмена")

        hBtn = QHBoxLayout()
        hBtn.addWidget(okButton)
        hBtn.addWidget(cancelButton)

        layout = QVBoxLayout()
        layout.addLayout(hTeachFind)
        layout.addWidget(self.teach_select)
        layout.addLayout(hZamenaStart)
        layout.addLayout(hZamenaEnd)
        layout.addLayout(hBtn)

        container = QWidget()
        container.setLayout(layout)

        container.setFixedWidth(250)
        self.setCentralWidget(container)

    def teachFind(self, text):
        fioSrez = text.capitalize()
        lenFioSrez = len(fioSrez)
        sotrudniki_find = []
        for i in range(len(sotrudniki_fio)):
            if sotrudniki_fio[i][:lenFioSrez] == fioSrez:
                sotrudniki_find.append(sotrudniki_fio[i])
        if len(sotrudniki_find) > 0:
            self.teach_select.clear()
            self.teach_select.addItems(sotrudniki_find)
        else:
            self.teach_select.addItems(sotrudniki_fio)


app = QApplication(sys.argv)
w = MainWindow()
w.show()
app.exec()
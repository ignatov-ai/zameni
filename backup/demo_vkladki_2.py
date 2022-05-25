import sys
from PyQt6.QtWidgets import (
    QMainWindow, QApplication, QPushButton,
    QLabel, QComboBox, QListView, QLineEdit, QDateTimeEdit, QDateEdit,
    QLineEdit, QVBoxLayout, QWidget, QHBoxLayout
)
from PyQt6.QtCore import QDate
from PyQt6.QtWidgets import *
from PyQt6.uic.properties import QtGui

window = QtGui.QWidget()
tab = QtGui.QTabWidget()
tab.addTab(QtGui.QLabel("Содержимое вкладки 1"), "Вкладка &1")
tab.addTab(QtGui.QLabel("Содержимое вкладки 2"), "Вкладка &2")
tab.addTab(QtGui.QLabel("Содержимое вкладки 3"), "Вкладка &3")
tab.setCurrentIndex(0)
vbox = QtGui.QVBoxLayout()
vbox.addWidget(tab)
window.setLayout(vbox)
window.show()
import sys
from random import randint
from PyQt5 import QtWidgets, QtCore, QtGui


class CustomTableView(QtWidgets.QTableView):
    def __init__(self, parent=None):
        super(CustomTableView, self).__init__(parent)
        self.setSortingEnabled(True)

    def KeyPressEvent(self, event: QtGui.QKeyEvent):
        if event.key() == QtCore.Qt.Key_Enter:
            print("Key_Enter ")
        elif event.key() == QtCore.Qt.Key_Return:
            print("Key_Return ")


class NumberSortModel(QtCore.QSortFilterProxyModel):

    def lessThan(self, left_index: "QModelIndex",
                 right_index: "QModelIndex") -> bool:

        left_var: str = left_index.data(QtCore.Qt.EditRole)
        right_var: str = right_index.data(QtCore.Qt.EditRole)

        try:
            return float(left_var) < float(right_var)
        except (ValueError, TypeError):
            pass

        try:
            return left_var < right_var
        except TypeError:
            return True


class Counter(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(Counter, self).__init__(parent)
        self.setWindowFlags(QtCore.Qt.Window)
        QtWidgets.QMainWindow.__init__(self)

        font = QtGui.QFont("Formula1", 10, QtGui.QFont.Bold)
        self.setFont(font)

        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)

        grid_layout = QtWidgets.QGridLayout()
        central_widget.setLayout(grid_layout)

        self.model = QtGui.QStandardItemModel(self)
        self.model.setHorizontalHeaderLabels(["Name", "Points"])

        self.proxy = NumberSortModel()
        self.proxy.setSourceModel(self.model)

        self.table = CustomTableView(self)
        self.table.setModel(self.proxy)
        for i in range(10):
            self.model.appendRow([QtGui.QStandardItem(f'Name{randint(10, 99)}'),
                                  QtGui.QStandardItem(str(randint(1, 100)))])

        sort_button = QtWidgets.QPushButton("Sort")

        qlineedit_name = QtWidgets.QLineEdit()
        qlineedit_name.resize(24, 80)
        qlineedit_name.setText("Name")
        qlineedit_name.selectAll()
        qlineedit_points = QtWidgets.QLineEdit()
        qlineedit_points.resize(24, 80)
        qlineedit_points.setText("Points")
        qlineedit_points.selectAll()

        horisontal_layout = QtWidgets.QHBoxLayout()
        horisontal_layout.addStretch(1)
        horisontal_layout.addWidget(qlineedit_name)
        horisontal_layout.addStretch(1)
        horisontal_layout.addWidget(qlineedit_points)
        horisontal_layout.addStretch(1)
        horisontal_layout.addWidget(sort_button)
        horisontal_layout.addStretch(1)
        horisontal_layout.setAlignment(QtCore.Qt.AlignRight)

        grid_layout.addLayout(horisontal_layout, 0, 0)
        grid_layout.addWidget(self.table, 1, 0)


if __name__ == "__main__":
    application = QtWidgets.QApplication([])
    window = Counter()
    window.setWindowTitle("Counter")
    window.setMinimumSize(480, 380)
    window.show()
    sys.exit(application.exec_())
import sys
from PyQt5.QtWidgets import QWidget, QLabel, QPushButton, QLineEdit, QApplication, QGridLayout, QMessageBox
from PyQt5.QtCore import Qt
# from functools import partial
from settings import *
from app import write_file


class Example(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        file = QLabel('File:', self)
        report1 = QLabel('Report 1:', self)
        report2 = QLabel('Report 2:', self)
        address = QLabel('Address:', self)
        last = QLabel('Last:', self)
        address.setAlignment(Qt.AlignCenter)
        last.setAlignment(Qt.AlignCenter)

        grid = QGridLayout()
        grid.setSpacing(15)

        self.fileEdit = QLineEdit(file_name, self)
        self.rep1AddrEdit = QLineEdit(str(rep1_addr_default), self)
        self.rep1AddrEdit.setAlignment(Qt.AlignCenter)
        self.rep1LastEdit = QLineEdit(str(last_matches_report1_default), self)
        self.rep1LastEdit.setAlignment(Qt.AlignCenter)
        self.rep2AddrEdit = QLineEdit(str(rep2_addr_default), self)
        self.rep2AddrEdit.setAlignment(Qt.AlignCenter)
        self.rep2LastEdit = QLineEdit(str(last_matches_report2_default), self)
        self.rep2LastEdit.setAlignment(Qt.AlignCenter)

        self.btnCalc = QPushButton('Calculate', self)
        self.btnCalc.clicked.connect(self.calc)

        grid.addWidget(file, 1, 1)
        grid.addWidget(self.fileEdit, 1, 2, 1, 3)
        grid.addWidget(address, 2, 2)
        grid.addWidget(last, 2, 3)
        grid.addWidget(report1, 3, 1)
        grid.addWidget(self.rep1AddrEdit, 3, 2)
        grid.addWidget(self.rep1LastEdit, 3, 3)
        grid.addWidget(report2, 4, 1)
        grid.addWidget(self.rep2AddrEdit, 4, 2)
        grid.addWidget(self.rep2LastEdit, 4, 3)
        grid.addWidget(self.btnCalc, 5, 3)

        self.setLayout(grid)
        self.setGeometry(300, 300, 250, 150)
        self.setWindowTitle('Analysis')
        self.show()

    def message(self):
        QMessageBox.information(self, "Information", "Done")

    def calc(self):
        excel_file = self.fileEdit.text()
        rep1_addr = self.rep1AddrEdit.text()
        rep2_addr = self.rep2AddrEdit.text()
        rep1_last = int(self.rep1LastEdit.text())
        rep2_last = int(self.rep2LastEdit.text())
        write_file(excel_file, rep1_addr, rep2_addr, rep1_last, rep2_last)
        self.message()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())


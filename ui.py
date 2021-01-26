import sys
import main
from PyQt5 import QtWidgets
import PyQt5.QtWidgets


class Example(PyQt5.QtWidgets.QWidget):

    def __init__(self):
        super().__init__()
        self.directory_selection = None
        self.setGeometry(500, 300, 600, 300)
        self.setWindowTitle("ISL / SSL Close Program")
        self.initUI()

    def initUI(self):
        self.label = QtWidgets.QLabel(self)
        self.label.move(50, 50)
        self.label.setText("HELLLOOOOO")

        self.btn_browse = PyQt5.QtWidgets.QPushButton(self)
        self.btn_browse.setObjectName("btn_browse")
        self.btn_browse.setText("Browse...")
        self.btn_browse.move(50, 150)
        self.btn_browse.clicked.connect(self.directory_dialog)

        self.btn_cancel = PyQt5.QtWidgets.QPushButton(self)
        self.btn_cancel.setObjectName("btn_cancel")
        self.btn_cancel.setText("Cancel")
        self.btn_cancel.move(150, 150)

        self.btn_Start = PyQt5.QtWidgets.QPushButton(self)
        self.btn_Start.setObjectName("btn_start")
        self.btn_Start.setText("Start")
        self.btn_Start.move(250, 150)
        self.btn_Start.clicked.connect(self.start_close)
        self.btn_Start.setVisible(False)

    def directory_dialog(self):
        directory_string = str(PyQt5.QtWidgets.QFileDialog.getExistingDirectory(self, "Select Directory")) + "/"
        self.directory_selection = directory_string.replace("/", "\\")
        print(self.directory_selection)
        if self.directory_selection is not None:
            print(self.directory_selection)
            self.label.setText(self.directory_selection)
            self.label.adjustSize()
            self.btn_Start.setVisible(True)

    def string_from_inputdialog(self):
        string, ok_pressed = QtWidgets.QInputDialog.getText(self, "Get text", "Your name:", QtWidgets.QLineEdit.Normal,
                                                            "")
        if ok_pressed and string != '':
            return string

    def start_close(self):
        # community_name = str(QtWidgets.QInputDialog.getText(self, 'Input Dialog', 'Community Name:'))
        community_name = self.string_from_inputdialog()
        print("Conversion Start")
        main.convert_files(str(self.directory_selection))
        print("Conversion End")
        print("Book Transfers begin")
        main_book = main.create_main_book(community_name)
        print("Book transfer complete")
        print("Data move begin")
        main.move_data_to_main_file(main_book, community_name, self.directory_selection)
        print("Data Move complete")
        print("Close Starting")
        main.complete_month_end_close(main_book, community_name)
        print("We be closed now!")

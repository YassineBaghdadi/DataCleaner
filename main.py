from PyQt5 import QtWidgets, uic
import os, sys


class Main(QtWidgets.QWidget):
    def __init__(self):
        super(Main, self).__init__()
        uic.loadUi(os.path.join(os.getcwd(), 'main.ui'), self)


        self.
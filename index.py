import ntpath
import os

import pandas as pd

from PyQt5 import QtCore, QtGui, QtWidgets
from win10toast import ToastNotifier


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(480, 160)
        Form.setMaximumSize(QtCore.QSize(480, 160))
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.filePath = QtWidgets.QLineEdit(Form)
        self.filePath.setReadOnly(True)
        self.filePath.setObjectName("filePath")
        self.horizontalLayout.addWidget(self.filePath)
        self.brows = QtWidgets.QPushButton(Form)
        self.brows.setObjectName("brows")
        self.horizontalLayout.addWidget(self.brows)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.verticalLayout_2.addLayout(self.verticalLayout)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setContentsMargins(15, -1, 15, -1)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.cols = QtWidgets.QComboBox(Form)
        self.cols.setObjectName("cols")
        self.horizontalLayout_3.addWidget(self.cols)
        self.verticalLayout_2.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.start = QtWidgets.QPushButton(Form)
        self.start.setMaximumSize(QtCore.QSize(100, 16777215))
        self.start.setObjectName("start")
        self.horizontalLayout_2.addWidget(self.start)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        self.brows.setText("...")
        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

        self.brows.clicked.connect(self.browes)
        self.refresh()
        self.start.clicked.connect(self.StartCleaning)

    def refresh(self):
        if self.filePath.text():
            self.start.setEnabled(True)
        else:
            self.start.setEnabled(False)

    def browes(self):
        self.fileName, _ = QtWidgets.QFileDialog.getOpenFileName(None, 'Single File', QtCore.QDir.rootPath(), '*.xlsx')
        if self.fileName:
            print(self.fileName)
            self.filePath.setText(self.fileName)
            self.data = pd.read_excel(self.fileName)
            self.cols.addItems(list(self.data.columns))

            print(list(self.data.columns))

        self.refresh()


    def StartCleaning(self):
        col = int(self.cols.currentIndex())
        self.data.drop_duplicates(subset=[self.cols.currentText()])
        sCh = ['*', '/', '\\', '-', '+', '@', '#', '$', '!', '~', '`', '%', '^', '&', '(', ')', '_', '=', '.', ',', '"', "'", '[', ']', '{', '}', '|', '­']
        print(col)
        dataOut = []

        failed = []
        ignored = []
        for i, v in enumerate(self.data.values.tolist()):
            if v[col] != '' or v[col] != ' ':
                if str(v[col]).strip()[:3] != '+44':
                    try:
                        correct_number = ''.join([str(i) for i in str(v[col]) if i.isdigit()])
                        v[col] = int(correct_number[-10:])
                    except:
                        failed.append(v)
                    dataOut.append(v)
                else:
                    ignored.append(v)

        out = pd.DataFrame(dataOut)

        outfailed = pd.DataFrame(failed)

        outignored = pd.DataFrame(ignored)
        filename, _ = QtWidgets.QFileDialog.getSaveFileName(None, "Save :", f'{str(ntpath.basename(self.fileName)).split(".")[0]}_clean.xlsx', "Excel Files (*.xlsx)")
        if filename:
            out.to_excel(filename)
            outfailed.to_excel(os.path.join(os.path.dirname(os.path.abspath(filename)), f'{str(ntpath.basename(self.fileName)).split(".")[0]}_Failed.xlsx'))
            outignored.to_excel(os.path.join(os.path.dirname(os.path.abspath(filename)),f'{str(ntpath.basename(self.fileName)).split(".")[0]}_Ignored.xlsx'))

        toast = ToastNotifier()
        txt = f"""
        The Data Has Been Cleaned Successfully :
        Done : {len(dataOut)}
        Ignored : {len(ignored)}
        Failed : {len(failed)}
        
        """
        toast.show_toast(f"{ntpath.basename(self.fileName)} Has Finished Cleaning ", txt, duration=1500)




    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "DataCleaner SACCOM CC"))
        self.brows.setText(_translate("Form", "..."))
        self.start.setText(_translate("Form", "Start"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

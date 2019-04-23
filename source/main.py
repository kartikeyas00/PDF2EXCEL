# -*- coding: utf-8 -*-

import sys
from functools import partial
from PyQt5 import QtCore, QtGui, QtWidgets, uic
import pandas as pd
from PandasModel import PandasModel
from PdfObject import PdfObject
import os

class App(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(App, self).__init__(parent)
        designUI=self.resource_path("resources\designUI.ui")
        uic.loadUi(designUI, self)

        self.filename = ""
        self.converted_file = None

        thread = QtCore.QThread(self)
        thread.start()
        self.pdf_object = PdfObject()
        self.pdf_object.moveToThread(thread)
        self.pdf_object.maximumChanged.connect(self.progressBar.setMaximum)
        self.pdf_object.progressChanged.connect(self.progressBar.setValue)
        self.pdf_object.pandasChanged.connect(self.on_pandasChanged)
        self.pushButton.clicked.connect(self.openFileNameDialog) #SELECT FILE
        self.pushButton_3.clicked.connect(self.convert)  #CONVERT
        self.pushButton_2.clicked.connect(self.view)     #VIEW
        self.pushButton_4.clicked.connect(self.saveFileDialog) #EXPORT
        
    def resource_path(self,relative_path):
        """ Get absolute path to resource, works 
        for dev and for PyInstaller """
        
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
    
        return os.path.join(base_path, relative_path)    
    
    def openFileNameDialog(self):        
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Open File", "", ".pdf(*.pdf)"
        )  
        if fileName:
            self.filename = fileName

    def view(self):
        if self.converted_file is not None:
            model = PandasModel(self.converted_file)
            self.tableView.setModel(model)

    def convert(self):
        if self.filename:
            wrapper = partial(self.pdf_object.pdf2excel, self.filename)
            QtCore.QTimer.singleShot(0, wrapper)

    @QtCore.pyqtSlot(pd.DataFrame)
    def on_pandasChanged(self, df):
        self.converted_file = df.copy()

    def saveFileDialog(self):        
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save File", "", ".xls(*.xls)"
        )
        if fileName and self.converted_file is not None:
            self.converted_file.to_excel(fileName,index=False)
            msg = QtWidgets.QMessageBox()
            msg.setText("File is Saved")
            msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msg.exec_()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    ex = App()
    app.setWindowIcon(QtGui.QIcon(ex.resource_path("resources\pdf-to-excel-icon.png")))
    ex.show()
    app.exec_()
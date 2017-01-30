# encoding: utf-8
import sys

from PyQt5.QtCore import pyqtSlot, QDir
from PyQt5.QtWidgets import (QApplication, QDialog, QFileDialog)
from PyQt5.uic import loadUiType
from excelreader import process_excel_1

app = QApplication(sys.argv)
form_class, base_class = loadUiType('excel.ui')


class DemoImpl(QDialog, form_class):
    def __init__(self, *args):
        super(DemoImpl, self).__init__(*args)
        self.setupUi(self)
        self.outdir = '.'
    
    @pyqtSlot()
    def on_add_excel_clicked(self):
        file_dir = QFileDialog.getOpenFileName(self, u"选择源数据Excel", QDir.currentPath())
        self.in_file.setText(file_dir[0])
        process_excel_1(file_dir[0], 'template.docx', self.outdir)

    @pyqtSlot()
    def on_sel_out_dir_clicked(self):
        dir = QFileDialog.getExistingDirectory(self, u"选择输出目录", QDir.currentPath())
        self.out_dir.setText(dir)
        self.outdir = dir

form = DemoImpl()
form.show()
sys.exit(app.exec_())
import sys
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
)
from form import Ui_MainWindow
from docxtpl import DocxTemplate


class AppWindow(QMainWindow):
    def __init__(self):
        super(AppWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.creatWord)

    def creatWord(self):
        Name = self.ui.lineEdit.text()
        Fam = self.ui.lineEdit_2.text()
        Oth = self.ui.lineEdit_3.text()
        doc = DocxTemplate('шаблон.docx')
        context = { 'Name' : Name, 'Fam' : Fam, 'Oth' : Oth}
        try:
            doc.render(context)
            doc.save("шаблон-final.docx")
        except:
            print("Закройте документ - шаблон-final.docx")

if __name__ == '__main__':
    app = QApplication([])
    AppWindow = AppWindow()
    AppWindow.show()
    sys.exit(app.exec())
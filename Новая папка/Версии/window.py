import sys
import time
from PyQt5.QtWidgets import QWidget, QApplication, QLabel


import sys
from PyQt5.QtWidgets import QMainWindow, QApplication


class Example(QMainWindow):

    def __init__(self):
        super().__init__()

        self.initUI()


    def initUI(self):

        self.statusBar().showMessage('Ready')

        self.setGeometry(300, 300, 250, 150)
        self.setWindowTitle('Statusbar')
        self.show()
    def UpdateWindow(self, arg1):
        self.statusBar().showMessage(arg1)##        self.setWindowTitle(arg1)
##        self.lbl.setText(arg1)
##        self.lbl.adjustSize()

if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = Example()
    for ii in range(0,10):
        ex.UpdateWindow(str(ii))
        time.sleep(1)
    
##class MyWindow(QWidget):
##
##    def __init__(self):
##        super().__init__()
##        self.initWindow()
##    def initWindow(self):
##        self.lbl = QLabel(self)
##        self.lbl.setText("lll")
##        self.setGeometry(300, 300, 250, 150)
##        self.show()
##    def UpdateWindow(self, arg1):
##        self.setWindowTitle(arg1)
##        self.lbl.setText(arg1)
##        self.lbl.adjustSize()
##    def CloseWindow(self):
##        self.destroy()
##
##if __name__ == '__main__':
##
##    app = QApplication(sys.argv)
##    ex = MyWindow()
    ##sys.exit(app.exec_())
    ##for ii in range(0,10):
      ##  ex.UpdateWindow(str(ii))
        ##time.sleep(1)
    ##ex.CloseWindow()
    ##sys.exit(app.exec_())
    ##ex.CloseWindow()
##import sys
##from PyQt5.QtWidgets import (QWidget, QLabel,
##    QLineEdit, QApplication)
##
##
##class Example(QWidget):
##
##    def __init__(self):
##        super().__init__()
##
##        self.initUI()
##
##
##    def initUI(self):
##
##        self.lbl = QLabel(self)
##        qle = QLineEdit(self)
##
##        qle.move(60, 100)
##        self.lbl.move(60, 40)
##
##        qle.textChanged[str].connect(self.onChanged)
##
##        self.setGeometry(300, 300, 280, 170)
##        self.setWindowTitle('QLineEdit')
##        self.show()
##
##
##    def onChanged(self, text):
##
##        self.lbl.setText(text)
##        self.lbl.adjustSize()
##
##
##if __name__ == '__main__':
##
##    app = QApplication(sys.argv)
##    ex = Example()
##    sys.exit(app.exec_())

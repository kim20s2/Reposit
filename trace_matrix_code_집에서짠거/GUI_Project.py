import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMainWindow, QAction, qApp
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QDate, Qt


class MyApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.date = QDate.currentDate()
        self.initUI()

    def initUI(self):
        exitAction = QAction('Exit', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('Exit application')
        exitAction.triggered.connect(qApp.quit)

        self.statusBar().showMessage(self.date.toString(Qt.DefaultLocaleLongDate))

        menubar = self.menuBar()
        menubar.setNativeMenuBar(False)
        filemenu = menubar.addMenu('&File')
        filemenu.addAction(exitAction)

        btn = QPushButton('DocID_Info 파일 생성', self)
        btn.move(300, 50)
        btn.resize(btn.sizeHint())

        btn = QPushButton('SysID_Info 파일 생성', self)
        btn.move(300, 100)
        btn.resize(btn.sizeHint())

        btn = QPushButton('SwID_Info 파일 생성', self)
        btn.move(300, 150)
        btn.resize(btn.sizeHint())

        btn = QPushButton('SysSwIT_Info 파일 생성', self)
        btn.move(300, 200)
        btn.resize(btn.sizeHint())

        btn = QPushButton('TestResult_Info 파일 생성', self)
        btn.move(300, 250)
        btn.resize(btn.sizeHint())

        btn = QPushButton('Write 파일 생성', self)
        btn.move(300, 300)
        btn.resize(btn.sizeHint())


        self.setWindowTitle('Trace Matrix 생성 Tool')
        self.setWindowIcon(QIcon('web.png'))
        self.setGeometry(300, 300, 600, 400)
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())

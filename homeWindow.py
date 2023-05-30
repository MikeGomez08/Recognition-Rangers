from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtGui import QGuiApplication
from PyQt5.QtCore import pyqtSlot, QTimer
from PyQt5.uic import loadUi
from main import MainWindow
from adminWindow import AdminWindow
from loginDialog import LogIn

import sys

class HomeWindow(QMainWindow):
    # initialize MyWindow Class
    def __init__(self):
        # call parent constructor
        super(HomeWindow, self).__init__()

        # load user interface for login window
        loadUi("homeWindowSmaller.ui", self)

        # call start function when employee button is clicked
        self.employeeButton.clicked.connect(self.start)

        # call sign in function when admin button is clicked
        self.adminButton.clicked.connect(self.admin_start)

        # declare attribute for the main window
        self.mainWindow = None
        self.adminWindow = None
        self.loginDialogue = None

        # cancel load of program when x button is pressed
        self.exitButton.clicked.connect(lambda: sys.exit("Program Cancelled"))

        # remove title bar and add style to form
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        # timer for delay to show splash screen
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self.delay_show)
        self.clock_timer.start(5000)

        # return to home window when other windows are closed




    def reinitialize(self):
        self.show()

    # python decorator for functions connected through signals (ex. click of a button)
    @pyqtSlot()
    # switch to the main program window
    def start(self):

        # open main window
        self.mainWindow = MainWindow()
        self.mainWindow.show()
        self.mainWindow.destroyed.connect(self.reinitialize)

        # close login window
        self.hide()


    # python decorator for functions connected through signals (ex. click of a button)
    @pyqtSlot()
    def admin_start(self):

        # open LogIn dialogue
        self.loginDialogue = LogIn()
        self.loginDialogue.show()

        # hide home window
        self.hide()

        # open admin window if login is successful
        self.loginDialogue.adminWindowOpened.connect(self.openAdmin)
        self.loginDialogue.loginExited.connect(self.reinitialize)


    def openAdmin(self):
        self.adminWindow = AdminWindow()
        self.adminWindow.show()
        self.hide()

        self.adminWindow.destroyed.connect(self.reinitialize)

    def delay_show(self):
        self.clock_timer.stop()
        self.show()
        splashWindow.close()


class SplashScreen(QMainWindow):
    def __init__(self):
        super(SplashScreen, self).__init__()

        loadUi("splashscreen.ui", self)

        # cancel load of program when x button is pressed
        self.exitButton.clicked.connect(lambda: sys.exit("Program Cancelled"))

        # remove title bar
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

# main function
if __name__ == "__main__":
    # set up qt app
    app = QApplication(sys.argv)
    # prevent from automatic closing
    app.setQuitOnLastWindowClosed(False)

    # splash screen
    splashWindow = SplashScreen()
    splashWindow.show()

    # create instance of login window
    homeWindow = HomeWindow()



    # cleanly exits when x button is clicked
    sys.exit(app.exec_())


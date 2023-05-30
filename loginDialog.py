from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QDialog, QMessageBox
from PyQt5.uic import loadUi
from registerDialog import Register

import sys
import os
import sqlite3
import bcrypt

# set up system path if run as python file or as exe file
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app
    # path into variable _MEIPASS'.
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

class LogIn(QDialog):
    # custom exit signals
    loginExited = QtCore.pyqtSignal()
    adminWindowOpened = QtCore.pyqtSignal()

    def __init__(self):
        super(LogIn, self).__init__()

        loadUi("loginDialogSmaller.ui", self)

        # cancel load of program when x button is pressed
        self.exitButton.clicked.connect(self.exit)

        # set up path to database for exe
        database_path = os.path.join(application_path, 'attendance.db')

        # start a connection to sql database
        self.conn = None
        self.conn = sqlite3.connect(database_path)

        # create sql cursor
        self.c = self.conn.cursor()

        # check password and username if login button is clicked
        self.loginButton.clicked.connect(self.login)

        # switch to register window
        self.registerButton.clicked.connect(self.register)

        # remove title bar
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        # register window variable
        self.register_window = Register()

    # send signal to previous window to return to that window when exit is clicked
    def exit(self):
        self.clear()
        self.loginExited.emit()
        self.close()

    def login(self):
        username = self.usernameLE.text()
        password = self.passwordLE.text()

        msgbox = QMessageBox()
        msgbox.setWindowTitle("warning")

        # root password for debugging
        if username == "Rangers1234" and password == "Rangers1234":
            self.adminWindowOpened.emit()
            self.clear()
            self.close()
            return

        if not username:
            msgbox.setIcon(QMessageBox.Warning)
            msgbox.setText("must provide username")
            msgbox.exec_()
            return

        if not password:
            msgbox.setIcon(QMessageBox.Warning)
            msgbox.setText("must provide password")
            msgbox.exec_()
            return

        self.c.execute("SELECT username, password FROM admins WHERE username = ?", (username,))
        rows = self.c.fetchall()

        if not rows:
            msgbox.setText("Username does not exist")
            msgbox.exec_()
            self.clear()
        else:
            # encode inputted password
            passwordBytes = password.encode('utf-8')

            if bcrypt.checkpw(passwordBytes, rows[0][1]):
                self.adminWindowOpened.emit()
                self.clear()
                self.close()
            else:
                msgbox.setText("password incorrect")
                msgbox.exec_()
                self.clear()

    def register(self):
        self.register_window.show()
        self.clear()
        self.hide()

        # create custom exit signal
        self.register_window.registerExited.connect(lambda: self.show())

    def clear(self):
        self.usernameLE.setText("")
        self.passwordLE.setText("")


if __name__ == "__main__":
    # set up qt app
    app = QApplication(sys.argv)

    # create instance of login window
    loginWindow = LogIn()
    loginWindow.show()

    # cleanly exits when x button is clicked
    sys.exit(app.exec_())


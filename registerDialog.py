from PyQt5 import QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QDialog, QMessageBox
from PyQt5.uic import loadUi

import sqlite3
import sys, os

# set up system path if run as python file or as exe file
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app
    # path into variable _MEIPASS'.
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

class Register(QDialog):
    # custom exit signal
    registerExited = QtCore.pyqtSignal()

    def __init__(self):
        super(Register, self).__init__()

        loadUi("registerDialogSmaller.ui", self)

        # remove title bar
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        # set up path to database for exe
        database_path = os.path.join(application_path, 'attendance.db')

        # start a connection to sql database
        self.conn = None
        self.conn = sqlite3.connect(database_path)
        # create sql cursor
        self.c = self.conn.cursor()

        # return to login screen if exit button is pressed
        self.exitButton.clicked.connect(self.close)

        self.registerButton.clicked.connect(self.register)

    # send signal to previous window to return to that window when exit is clicked
    def closeEvent(self, event):
        self.registerExited.emit()

    def register(self):
        # set up message box for displaying errors
        msgbox = QMessageBox()
        msgbox.setWindowTitle("warning")
        msgbox.setIcon(QMessageBox.Warning)

        # get user input
        employeeID = self.empIDLE.text()
        username = self.usernameLE.text()
        password = self.passwordLE.text()
        confirmPass = self.confirmLE.text()

        if password != confirmPass:
            msgbox.setText("Passwords do not match")
            return

        self.c.execute("SELECT hr_managers.manager_ID, hr_managers.employee_ID FROM hr_managers WHERE employee_ID = ? ", (employeeID,))
        rows = self.c.fetchall()

        self.c.execute("SELECT username, manager_ID FROM admins")
        admin_rows = self.c.fetchall()
        username_taken = False
        already_admin = False

        i = 0
        for row in admin_rows:
            if row[0] == username:
                username_taken = True

            if rows[i][0] == row[1]:
                already_admin = True

            i += 1

        # if employeeID is not in managers list then cancel
        if not rows:
            msgbox.setText("""The employee ID entered is not eligible for admin registration. 
                            Please ask a manager or an admin for assistance.""")
            msgbox.exec_()
            self.clear()
            return
        else:
            if username_taken:
                msgbox.setText("Error: Username already taken")
                msgbox.exec_()
                self.clear()
            elif already_admin:
                msgbox.setText("Error: Employee is already an admin")
                msgbox.exec_()
                self.clear()
            else:
                # convert to bytes for hashing
                passwordBytes = password.encode('utf-8')
                # generate password salt
                salt = bcrypt.gensalt()
                # Hashing the password
                hash = bcrypt.hashpw(passwordBytes, salt)

                self.c.execute("INSERT INTO admins (manager_ID, username, password) VALUES(?, ?, ?)", (rows[0][0], username, hash))
                self.conn.commit()

                msgbox.setText("Registration Successful")
                msgbox.exec_()
                self.clear()

    # function for clearing input fields
    def clear(self):
        self.empIDLE.setText("")
        self.usernameLE.setText("")
        self.passwordLE.setText("")
        self.confirmLE.setText("")
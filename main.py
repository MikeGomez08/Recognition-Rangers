from PyQt5.QtGui import QImage, QPixmap, QColor
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSlot, QTimer, Qt
from PyQt5.QtWidgets import QMainWindow, QMessageBox
from numpy import array

import sqlite3
import cv2
import face_recognition
import numpy as np
import datetime
import sys
import os

# set up system path if run as python file or as exe file
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app
    # path into variable _MEIPASS'.
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

class MainWindow(QMainWindow):
    # initialize MainWindow Class
    def __init__(self):
        # call parent constructor
        super(MainWindow, self).__init__()

        # load user interface for main window
        loadUi("mainSmaller.ui", self)

        # start maximized
        self.setWindowState(Qt.WindowMaximized)

        # variable for video frame
        self.current_frame = None

        # timer for returning picture label to default
        self.picture_timer = QTimer(self)

        # list for employee dictionaries and list of face encodings
        self.employees = []
        self.face_encodings = []

        # set up path to database for exe
        database_path = os.path.join(application_path, 'attendance.db')

        # start a connection to sql database
        self.conn = None
        self.conn = sqlite3.connect(database_path)

        # create sql cursor
        self.c = self.conn.cursor()

        # execute query to get all employee information
        self.c.execute("""SELECT employee_ID, departments.name, employees.name, face_image_path, face_encoding 
            FROM employees 
            INNER JOIN departments ON employees.department_ID = departments.department_ID
        """)

        # fetch results
        rows = self.c.fetchall()

        # load employee information into a dictionary and add to employee list
        for row in rows:
            employee = {
                "employee_ID": row[0],
                "department_name": row[1],
                "name": row[2],
                "face_path": row[3]
            }

            # find and load face encoding in the database if it does not exist yet
            if row[4] is None:
                face_image = cv2.imread(os.path.join(application_path, row[3]))
                # convert to rgb format
                face_image = cv2.cvtColor(face_image, cv2.COLOR_BGR2RGB)
                face_encoding = face_recognition.face_encodings(face_image)[0]

                self.c.execute("UPDATE employees SET face_encoding = ? WHERE employee_ID = ?", (str(face_encoding.tolist()), employee["employee_ID"]))
                self.conn.commit()

                self.face_encodings.append(face_encoding)
            else:
                self.face_encodings.append(np.array(eval(row[4])))

            self.employees.append(employee)



    # function that is called when window is showed
    def showEvent(self, event):
        # call functions
        self.startClock()
        self.startVideo()

    # method for updating clock
    @pyqtSlot()
    def startClock(self):
        """
        :return: no return value
        """

        now = datetime.datetime.now()
        current_date = now.strftime("%d/%m/%y")
        current_time = now.strftime("%I:%M %p")
        self.Date_Label.setText(current_date)
        self.Time_Label.setText(current_time)

        # timer to change time every 10 seconds (for synchronization)
        clock_timer = QTimer(self)
        clock_timer.timeout.connect(self.startClock)
        clock_timer.start(10000)

    # method to start video capture
    def startVideo(self):
        """
        :return: no return value
        """

        # index for which camera to use (0 for default system camera)
        camera_id = 0

        # Open camera for video capture
        self.capture = cv2.VideoCapture(camera_id)

        # call update frame method
        self.update_frame()

        # set up a timer to update frame every 10 milliseconds
        timer = QTimer(self)
        timer.timeout.connect(self.update_frame)
        timer.start(1)

    # method for updating video frame
    @pyqtSlot()
    def update_frame(self):
        # get frame from video capture, ret is the return value indicating whether a frame has been returned or not
        ret, self.current_frame = self.capture.read()

        # if the frame exists, set it to instance variable and display the frame in the ui
        if ret:
            # call find_faces method to detect faces
            self.find_faces()

            # convert frame to proper rgb format
            outImage = self.convertToRGB(self.current_frame)
            outImage = outImage.mirrored(horizontal=True, vertical=False)

            # display image to imgLabel in user interface
            self.imgLabel.setPixmap(QPixmap.fromImage(outImage))

    # method to put boxes in found faces
    def find_faces(self):

        # resize the frame to a smaller size to make computations faster
        img_small = cv2.resize(self.current_frame, (0, 0), None, 0.25, 0.25)

        # convert to rgb
        img_small = cv2.cvtColor(img_small, cv2.COLOR_BGR2RGB)

        # find locations of faces in the image
        faces_cur_frame = face_recognition.face_locations(img_small)

        # put boxes around the found faces
        for location in faces_cur_frame:
            y1, x2, y2, x1 = location

            # multiply by 4 to get original ratio
            y1, x2, y2, x1 = y1 * 4, x2 * 4, y2 * 4, x1 * 4

            cv2.rectangle(self.current_frame, (x1, y1), (x2, y2), (255, 255, 255), 2)

        # find for match in known faces when clock in or clock out button is clicked
        if self.ClockInButton.isChecked() or self.ClockOutButton.isChecked():
            encodings_cur_frame = face_recognition.face_encodings(img_small, faces_cur_frame)

            if self.face_encodings:
                for encoding in encodings_cur_frame:
                    match = face_recognition.compare_faces(self.face_encodings, encoding, tolerance=0.50)
                    face_dis = face_recognition.face_distance(self.face_encodings, encoding)
                    best_match_index = np.argmin(face_dis)

                    # pass in best_match_index, -1 if found face was unknown
                    if match[best_match_index]:
                        self.mark_attendance(best_match_index)
                    else:
                        self.mark_attendance(-1)
            else:
                self.mark_attendance(-1)

    # method for marking attendance
    def mark_attendance(self, best_match_index):
        """
        :param: best_match_index: index of found face match
        :return:
        """
        now = datetime.datetime.now()

        date_string = now.strftime("%Y-%m-%d")
        time_string = now.strftime("%H:%M:%S")

        datetime.datetime.strptime(date_string + " 12:00:00", '%Y-%m-%d %H:%M:%S')

        if best_match_index >= 0:
            employee_ID = self.employees[best_match_index]["employee_ID"]
            name = self.employees[best_match_index]["name"]
            department = self.employees[best_match_index]["department_name"]
            picture = self.employees[best_match_index]["face_path"]
            picture = cv2.imread(os.path.join(application_path, picture))

            # converts picture to proper rgb format
            picture = self.convertToRGB(picture)

        if self.ClockInButton.isChecked() or self.ClockOutButton.isChecked():
            if self.picture_timer.isActive():
                self.picture_timer.stop()
            self.ClockInButton.setEnabled(False)
            self.ClockOutButton.setEnabled(False)

            if best_match_index >= 0:
                # get employee time_in and time_out from database
                self.c.execute("SELECT time_in, time_out, break_in, break_out FROM attendance WHERE date = ? AND employee_ID = ?",
                               (date_string, employee_ID))
                rows = self.c.fetchall()

                if self.ClockInButton.isChecked():
                    # if rows is empty (employee has not timed in yet) then add attendance
                    if not rows:
                        self.c.execute("INSERT OR IGNORE INTO attendance (date,employee_ID,time_in) VALUES(?,?,?)",
                                       (date_string, employee_ID, time_string))
                        self.conn.commit()

                        # get employee time_in and time_out from database
                        self.c.execute("SELECT time_in, time_out FROM attendance WHERE date = ? AND employee_ID = ?",
                                       (date_string, employee_ID))
                        rows = self.c.fetchall()

                        self.update_labels(name, employee_ID, department, 'Successfully Logged in', '-', '-',
                                           rows[0][0])
                    # else if the employee already has a record and time in, but no time out, then output that user has already logged in
                    elif rows[0][1] is None:
                        self.update_labels(name, employee_ID, department, 'Already Logged in', '-', '-', rows[0][0])
                    # else if employee has complete record for the day, then output that user has already logged in and out
                    else:
                        time_in = datetime.datetime.strptime(date_string + " " + rows[0][0], '%Y-%m-%d %H:%M:%S')
                        time_out = datetime.datetime.strptime(date_string + " " + rows[0][1], '%Y-%m-%d %H:%M:%S')
                        elapsed_hours = time_out - time_in
                        hours = abs(elapsed_hours.total_seconds() / 60 ** 2)
                        minutes = abs(elapsed_hours.total_seconds() / 60) % 60

                        self.update_labels(name, employee_ID, department, 'Already Logged out',
                                           "{:.0f}".format(hours) + 'h',
                                           "{:.0f}".format(minutes) + 'm', rows[0][0], rows[0][1])

                elif self.ClockOutButton.isChecked():

                    # if rows is empty then employee has not timed in, notify the user of this
                    if not rows:
                        self.update_labels(name, employee_ID, department, 'Please clock in First',
                                           "0h", "0m", "--:-- --", "--:-- --")
                    else:
                        if rows[0][1] is None:
                            time_out = datetime.datetime.strptime(date_string + " " + time_string, '%Y-%m-%d %H:%M:%S')
                        else:
                            time_out = datetime.datetime.strptime(date_string + " " + rows[0][1], '%Y-%m-%d %H:%M:%S')

                        time_in = datetime.datetime.strptime(date_string + " " + rows[0][0], '%Y-%m-%d %H:%M:%S')
                        elapsed_hours = time_out - time_in
                        hours = abs(elapsed_hours.total_seconds() / 60 ** 2)
                        minutes = abs(elapsed_hours.total_seconds() / 60) % 60
                        total_time = "{:.0f}".format(hours) + "h" + "{:.0f}".format(minutes) + "m"

                        if rows[0][1] is None:
                            self.c.execute("UPDATE attendance SET time_out = ?, total_time = ? WHERE employee_ID = ?",
                                           (time_string, total_time, employee_ID))
                            self.conn.commit()
                            self.update_labels(name, employee_ID, department, 'Successfully Logged out',
                                               "{:.0f}".format(hours) + 'h',
                                               "{:.0f}".format(minutes) + 'm', rows[0][0], time_string)
                        else:
                            self.update_labels(name, employee_ID, department, 'Already Logged out',
                                               "{:.0f}".format(hours) + 'h',
                                               "{:.0f}".format(minutes) + 'm', rows[0][0], rows[0][1])

                # display employee image
                self.picLabel.setPixmap(QPixmap.fromImage(picture))
                self.export_attendance()

            else:
                self.NameLabel.setText("unknown")

            # enable buttons again
            self.ClockInButton.setEnabled(True)
            self.ClockInButton.setChecked(False)
            self.ClockOutButton.setEnabled(True)
            self.ClockOutButton.setChecked(False)

            # timer to remove labels after 10 seconds
            self.picture_timer.timeout.connect(self.remove_labels)
            self.picture_timer.start(5000)

    # method for removing labels
    @pyqtSlot()
    def remove_labels(self):
        self.picture_timer.stop()
        self.update_labels("-", "-", "-", "- status -", "0h", "0m", "--:-- --", "--:-- --")

        picture = cv2.imread(os.path.join(application_path, "icons/profile.png"))

        # converts picture to proper rgb format
        picture = self.convertToRGB(picture)

        self.picLabel.setPixmap(QPixmap.fromImage(picture))

    # method for updating labels
    def update_labels(self, name, employee_ID, department, status, hours, minutes, timeIn, timeOut = "--:-- --"):
        self.NameLabel.setText(name)
        self.Emp_ID.setText(employee_ID)
        self.Dep_Name.setText(department)
        self.StatusLabel.setText(status)
        self.HoursLabel.setText(hours)
        self.MinLabel.setText(minutes)
        self.TimeInLabel.setText("TIME IN: " + timeIn)
        self.TimeOutLabel.setText("TIME OUT: " + timeOut)

    # function to export to csv file from sql
    def export_attendance(self):
        # get the attendance records
        self.c.execute("""
        SELECT date, employees.name, employees.employee_ID, departments.name, time_in, time_out, total_time
        FROM employees 
        INNER JOIN departments ON employees.department_ID = departments.department_ID
        INNER JOIN attendance ON employees.employee_ID = attendance.employee_ID
        """)
        rows = self.c.fetchall()

        with open('Attendance.csv', 'w') as f:
            # write the headers
            f.writelines("Date, Name, Employee ID, Department, Clock in, Clock Out, Total Time")

            # write each row
            for row in rows:
                f.writelines(f'\n{row[0]},{row[1]},{row[2]},{row[3]},{row[4]},{row[5]},{row[6]}')

    # convert image to proper rgb format
    def convertToRGB(self, image):
        qformat = QImage.Format_Indexed8
        if len(image.shape) == 3:
            if image.shape[2] == 4:
                qformat = QImage.Format_RGBA8888
            else:
                qformat = QImage.Format_RGB888
        outImage = QImage(image, image.shape[1], image.shape[0], image.strides[0], qformat)

        # convert rgb to bgr for proper output
        outImage = outImage.rgbSwapped()

        return outImage

    # send signal to login window to return to that window
    def closeEvent(self, event):
        self.picture_timer.stop()
        self.capture.release()
        self.destroyed.emit()
        self.deleteLater()
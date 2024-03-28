import sys

import pymysql
import serial
import time
import threading
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QStandardItem, QStandardItemModel, QColor
from PyQt5.QtWidgets import QApplication, QMessageBox, QTableView, QTableWidgetItem

from view_activity import Ui_MainWindow


def koneksi():
    try:
        connection = pymysql.connect(
            host="localhost",
            user="admin",
            password="Marsel@010799",
            database="db_sari",
            charset="utf8mb4",
            cursorclass=pymysql.cursors.DictCursor,

        )
        return connection
    except pymysql.err.OperationalError as e:
        QMessageBox.critical(None, "Kesalahan", f"Kesalahan koneksi: {e}")
        return None

ser = serial.Serial(port='/dev/ttyUSB0', baudrate=9600, timeout=.1)
ser.flush()

currentX = 0
currentY = 0
currentZ = 0
currentK = 0
delay = 0

class WorkerThread(QThread):
    stepCompleted = pyqtSignal()

    def __init__(self, parent=None):
        super(WorkerThread, self).__init__(parent)

    def run(self):
        while True:
            self.stepCompleted.emit() 

class MainWindow(QtWidgets.QMainWindow, QThread):
    finished = pyqtSignal()
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.integrationModel = QStandardItemModel()
        self.createProjectModel = QStandardItemModel()
        self.saveActivityModel = QStandardItemModel()
        self.openProjectModel = QStandardItemModel()
        self.chooseProjectModel = QStandardItemModel()
        self.chooseActivityModel = QStandardItemModel()
        self.runningModel_1 = QStandardItemModel()
        self.runningModel_2 = QStandardItemModel()
        self.font = QtGui.QFont()
        self.font.setPointSize(14)
        
        self.kdProjectSignal = QtCore.pyqtSignal(str)
        self.defaultMenu()
        # Logout
        self.ui.btLogout.clicked.connect(self.setLogout)

        # Komponen bagian create project
        self.ui.tbCreateProjectTabel.setModel(self.createProjectModel)
        self.ui.txtCreateProjectCari.keyReleaseEvent = self.searchData
        self.ui.btToGetStarted.clicked.connect(self.setGetStarted)
        self.ui.btCreateProjectKembali.clicked.connect(self.setDashboard)
        self.ui.btCreateProjectSimpan.clicked.connect(self.setProyek)
        self.refreshCreateProjectTable()
        self.getCreateProjectCode()
        self.ui.tbCreateProjectTabel.resizeColumnsToContents()
        self.ui.tbCreateProjectTabel.verticalHeader().setDefaultSectionSize(50)

        # Komponen bagian save activity
        self.ui.btSaveActivityBack.clicked.connect(self.setDashboard)
        self.ui.tbSaveActivityTabel.setModel(self.saveActivityModel)
        self.ui.txtSaveActivityCari.keyReleaseEvent = self.searchDataActivity
        self.refreshSaveActivityTable()
        self.getCreateActivityCode()
        self.setupActivityTable()
        self.ui.btSaveActivitySimpan.clicked.connect(self.setActivityData)
        self.ui.tbSaveActivityTabel.resizeColumnsToContents()
        self.ui.tbSaveActivityTabel.verticalHeader().setDefaultSectionSize(50)
        self.ui.btSaveActivityOpen.clicked.connect(self.goToIntegration)
        self.ui.btSaveActivityDelete.clicked.connect(self.deleteActivity)

        # Komponen bagian Integration
        self.ui.btIntegrationKembali.clicked.connect(self.setDashboard)
        self.ui.slIntegrationX.valueChanged.connect(self.updateLineEditX)
        self.ui.slIntegrationY.valueChanged.connect(self.updateLineEditY)
        self.ui.slIntegrationZ.valueChanged.connect(self.updateLineEditZ)
        self.ui.slIntegrationK.valueChanged.connect(self.updateLineEditK)

        self.ui.slIntegrationX.sliderReleased.connect(self.releasedSlider)
        self.ui.slIntegrationY.sliderReleased.connect(self.releasedSlider)
        self.ui.slIntegrationZ.sliderReleased.connect(self.releasedSlider)
        self.ui.slIntegrationK.sliderReleased.connect(self.releasedSlider)

        self.ui.txtIntegrationX.editingFinished.connect(self.releasedText)
        self.ui.txtIntegrationY.editingFinished.connect(self.releasedText)
        self.ui.txtIntegrationZ.editingFinished.connect(self.releasedText)
        self.ui.txtIntegrationK.editingFinished.connect(self.releasedText)

        self.ui.txtIntegrationX.valueChanged.connect(self.updateSliderX)
        self.ui.txtIntegrationY.valueChanged.connect(self.updateSliderY)
        self.ui.txtIntegrationZ.valueChanged.connect(self.updateSliderZ)
        self.ui.txtIntegrationK.valueChanged.connect(self.updateSliderK)

        self.ui.tbIntegrationTabel.setModel(self.integrationModel)
        self.ui.btintegrationSetData.clicked.connect(self.setCoordinates)
        self.ui.txtIntegrationCari.keyReleaseEvent = self.searchDataIntegration
        self.refreshIntegrationTable()
        self.setupIntegrationTable()
        self.getIntegrationCode()
        self.ui.btIntegrationHapus.clicked.connect(self.deleteIntegration)
        self.ui.btIntegrationReset.clicked.connect(self.setIntegrationEmptyColumn)

        # Komponen bagian Open Project
        self.ui.btOpenProjectKembali.clicked.connect(self.setDashboard)
        self.ui.btOpenProject.clicked.connect(self.setOpenProject)
        self.ui.tbOpenProjectTabel.setModel(self.openProjectModel)
        self.ui.txtOpenProjectCari.keyReleaseEvent = self.searchOpenProjectData
        self.refreshOpenProjectTable()
        self.setupOpenProjectTable()
        self.ui.btOpenProjectBuka.clicked.connect(self.goToActivityMenu2)
        self.ui.btIntegrationSimpan.clicked.connect(self.setActivitySaved)
        self.ui.btOpenProjectDelete.clicked.connect(self.deleteProject)

        #  komponen bagian run program
        self.clicked_times = 0
        self.stacked_data = []
        self.timer = QTimer()
        self.is_process_running = False
        self.is_process_paused = False
        self.looping = False
        self.loop_thread = None
        self.last_paused_step = -1
        self.workerThread = WorkerThread()
        self.workerThread.stepCompleted.connect(self.sendLoopStep)
        self.timer.timeout.connect(self.sendNextStep)
        self.current_step = -1
        self.previous_z = None
        self.previous_data_row = None
        self.ui.tbRunningProjectTabel_1.setModel(self.runningModel_1)
        self.ui.tbRunningProjectTabel_2.setModel(self.runningModel_2)
        self.refreshRunningProjectData_1()
        self.ui.btRunning.clicked.connect(self.setRunProject)
        self.ui.btRunningProjectKembali.clicked.connect(self.setDashboard)
        self.ui.btRunningProjectShowProject.clicked.connect(self.setRunShowProject)
        self.ui.btRunningProjectShowActivity.clicked.connect(self.setRunShowActivity)
        self.ui.tbRunningProjectTabel_1.clicked.connect(self.rowClicked)
        self.ui.btRunningProjectRight.clicked.connect(self.moveToTable2)
        self.ui.btRunningProjectLeft.clicked.connect(self.removeFromTable2)
        self.startSendingSteps()
        self.ui.btRunningProjectRun.clicked.connect(self.toggleSendingSteps)
        self.ui.btRunningProjectStop.clicked.connect(self.actionSendeingSteps)
        self.ui.btRunningProjectLoop.clicked.connect(self.toggleSendingLoopSteps)
        self.ui.btRunningProjectStatus.setText("Not Starting")
        

        # Komponen bagian show project
        self.ui.tbChooseProjectTabel.setModel(self.chooseProjectModel)
        self.ui.txtChooseProjectCari.keyReleaseEvent = self.searchChooseProjectData
        self.refreshChooseProjectTable()
        self.setupChooseProjectTable()
        self.ui.btChooseProjectOpen.clicked.connect(self.goToRunProjectGetProject)
        self.ui.btChooseProjectBack.clicked.connect(self.setRunProject)

        # Komponen bagian show activity
        self.refreshChooseActivityTable()
        self.ui.tbChooseActivityTabel.setModel(self.chooseActivityModel)
        self.ui.txtChooseActivityCari.keyReleaseEvent = self.searchChooseActivity
        self.setupChooseActivityTable()
        self.ui.btChooseActivityOpen.clicked.connect(self.goToRunActivityGetProject)
        self.ui.btChooseActivityBack.clicked.connect(self.setRunProject)

        # Testing 
        # self.ui.btRunningProjectRun.clicked.connect(self.runLoopStep)

    def defaultMenu(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.dashboard_1)

    # Start Create Project==========================================================================================================

    def getCreateProjectCode(self):
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT MAX(kd_project) FROM project"
                    cursor.execute(sql)
                    result = cursor.fetchone()

                    if result["MAX(kd_project)"]:
                        latest_code = result["MAX(kd_project)"]
                        kode_int = int(latest_code[1:]) + 1
                        new_kode = f"P{kode_int:04d}"
                    else:
                        new_kode = "P0001"

                    self.ui.lbCreateProjectKode.setText(new_kode)
                    return new_kode
        finally:
            if connection:
                connection.close()

    def setProyek(self):
        a = self.ui.lbCreateProjectKode.text()
        b = self.ui.txtCreateProjectNama.text()
        c = self.ui.txtCreateProjectKet.toPlainText()

        if not a:
            QMessageBox.warning(self, "Warning", "Please input the project name.")
            return

        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM project WHERE nama_project=%s"
                    cursor.execute(sql, (b,))
                    result = cursor.fetchone()
                    if result:
                        QMessageBox.warning(
                            self,
                            "Warning",
                            f"The project named '{b}' already exists.",
                        )
                        return

                    sql = "INSERT INTO project (kd_project, nama_project, keterangan) VALUES (%s, %s, %s)"
                    cursor.execute(sql, (a, b, c))
                    connection.commit()

                    QMessageBox.information(
                        self, "Sukses", "The Project has been successfully saved."
                    )

                    self.goToActivityMenu1()
                    self.refreshCreateProjectTable()
                    self.setEmptyColumn()
                    self.getCreateProjectCode()

        finally:
            if connection:
                connection.close()

    def refreshCreateProjectTable(self):
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM project"
                    cursor.execute(sql)
                    result = cursor.fetchall()

                    self.createProjectModel.clear()
                    self.createProjectModel.setHorizontalHeaderLabels(
                        ["Project Code", "Project Name", "Description"]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    header_view = self.ui.tbCreateProjectTabel.horizontalHeader()

                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    for i, row_data in enumerate(result):
                        a = QStandardItem(row_data["kd_project"])
                        b = QStandardItem(row_data["nama_project"])
                        c = QStandardItem(row_data["keterangan"])

                        a.setFont(font)
                        b.setFont(font)
                        c.setFont(font)

                        a.setTextAlignment(QtCore.Qt.AlignCenter)
                        b.setTextAlignment(QtCore.Qt.AlignCenter)
                        c.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.createProjectModel.appendRow([a, b, c])
                        self.ui.tbCreateProjectTabel.resizeColumnsToContents()
                        self.ui.tbCreateProjectTabel.verticalHeader().setDefaultSectionSize(
                            50
                        )

        finally:
            if connection:
                connection.close()

    def searchData(self, cariData):
        cariData = self.ui.txtCreateProjectCari.text()

        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM project WHERE nama_project LIKE %s OR kd_project LIKE %s"
                    cursor.execute(
                        sql,
                        ("%" + cariData + "%", "%" + cariData + "%"),
                    )
                    result = cursor.fetchall()

                    self.createProjectModel.clear()
                    self.createProjectModel.setHorizontalHeaderLabels(
                        ["Project Code", "Project Name", "Description"]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    header_view = self.ui.tbCreateProjectTabel.horizontalHeader()

                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    for i, row_data in enumerate(result):
                        a = QStandardItem(row_data["kd_project"])
                        b = QStandardItem(row_data["nama_project"])
                        c = QStandardItem(row_data["keterangan"])

                        a.setFont(font)
                        b.setFont(font)
                        c.setFont(font)

                        a.setTextAlignment(QtCore.Qt.AlignCenter)
                        b.setTextAlignment(QtCore.Qt.AlignCenter)
                        c.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.createProjectModel.appendRow([a, b, c])
                        self.ui.tbCreateProjectTabel.resizeColumnsToContents()
                        self.ui.tbCreateProjectTabel.verticalHeader().setDefaultSectionSize(
                            50
                        )

        finally:
            if connection:
                connection.close()

    def keyReleaseEvent(self, event):
        if event.key() == QtCore.Qt.Key_Enter or event.key() == QtCore.Qt.Key_Return:
            self.searchData(event)
        else:
            super().keyReleaseEvent(event)

    def setEmptyColumn(self):
        self.ui.txtCreateProjectNama.setText("")
        self.ui.txtCreateProjectKet.setText("")
        self.ui.txtCreateProjectCari.setText("")

    # END Create Poject==========================================================================================================

    def setLogout(self):
        a = QMessageBox.question(
            self,
            "Exit Confirmation",
            "Are you sure you want to exit ?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if a == QMessageBox.Yes:
            QApplication.instance().quit()

    def setDashboard(self):
        a = QMessageBox.question(
            self,
            "Confirmation",
            "You will return to main course, Are you sure?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if a == QMessageBox.Yes:
            self.setIntegrationEmptyColumn()
            self.ui.stackedWidget.setCurrentWidget(self.ui.dashboard_1)

    def setGetStarted(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.createProject_2)

    def setIntegration(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.Integration_3)
        self.refreshIntegrationTable()
        self.setIntegrationEmptyColumn()

    def setActivitySaved(self):
        a = QMessageBox.question(
            self,
            "Confirmation",
            "Yeaay, The coordinats has been saved....",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if a == QMessageBox.Yes:
            self.setIntegrationEmptyColumn1()
            self.ui.stackedWidget.setCurrentWidget(self.ui.saveActivity_4)

    def setActivity(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.saveActivity_4)
        self.refreshSaveActivityTable()

    def setOpenProject(self):
        self.ui.stackedWidget.setCurrentWidget(self.ui.openProject_5)
        self.refreshOpenProjectTable()

    def setRunProject(self):
        self.refreshRunningProjectData_1()
        self.ui.stackedWidget.setCurrentWidget(self.ui.runProject_6)

    def setRunShowProject(self):
        self.refreshChooseProjectTable()
        self.ui.stackedWidget.setCurrentWidget(self.ui.getProject_7)
        
    def setRunShowActivity(self):
        self.refreshChooseActivityTable()
        self.refreshRunningProjectData_1()
        self.ui.stackedWidget.setCurrentWidget(self.ui.getActivity_8)

    # START Activity Menu=====================================================================================================

    def goToActivityMenu1(self):
        a = self.ui.lbCreateProjectKode.text()
        b = self.ui.txtCreateProjectNama.text()
        if a:
            self.ui.lbSaveActivityGetCode.setText(a)
            self.ui.lbCreateAcvtivityGetProjectName.setText(b)
            self.setActivity()

    def getCreateActivityCode(self):
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT MAX(kd_activity) FROM activity"
                    cursor.execute(sql)
                    result = cursor.fetchone()

                    if result["MAX(kd_activity)"]:
                        latest_code = result["MAX(kd_activity)"]
                        kode_int = int(latest_code[1:]) + 1
                        new_kode = f"A{kode_int:04d}"
                    else:
                        new_kode = "A0001"

                    self.ui.txtSaveActivityCode.setText(new_kode)
                    return new_kode
        finally:
            if connection:
                connection.close()

    def goToIntegration(self):
        a = self.ui.txtSaveActivityCode.text()
        b = self.ui.txtSaveActivityNama.text()
        if a:
            self.ui.lbIntegrationGetCode.setText(a)
            self.ui.txtIntegrationGetActivityName.setText(b)
            self.setIntegration()
            self.setActivityEmptyColumn()
            self.getCreateActivityCode()

    def setActivityData(self):
        a = self.ui.txtSaveActivityCode.text()
        b = self.ui.txtSaveActivityNama.text()
        c = self.ui.txtSaveActivityKet.toPlainText()
        d = self.ui.lbSaveActivityGetCode.text()

        if not a or not b or not c or not d:
            QMessageBox.warning(self, "Warning", "Please input your activity data.")
            return

        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM activity WHERE nama_activity=%s AND kd_project=%s"
                    cursor.execute(sql, (b, a))
                    result = cursor.fetchone()
                    if result:
                        QMessageBox.warning(
                            self,
                            "Warning",
                            f"The activity named '{b}' already exists.",
                        )
                        return

                    sql = "INSERT INTO activity (kd_activity, nama_activity, keterangan, kd_project) VALUES (%s, %s, %s, %s)"
                    cursor.execute(sql, (a, b, c, d))
                    connection.commit()

                    QMessageBox.information(
                        self, "Sukses", "Your Activity has been successfully saved."
                    )
                    self.goToIntegration()
                    self.refreshSaveActivityTable()
                    self.getCreateActivityCode()
                    self.setActivityEmptyColumn()

        finally:
            if connection:
                connection.close()

    def deleteActivity(self):
        a = self.ui.txtSaveActivityCode.text()

        if not a:
            QMessageBox.warning(
                self, "Warning", "Please input the activity code to delete."
            )
            return

        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql_select = "SELECT * FROM activity WHERE kd_activity = %s"
                    cursor.execute(sql_select, (a,))
                    result = cursor.fetchone()

                    if not result:
                        QMessageBox.warning(
                            self,
                            "Warning",
                            f"The activity with code '{a}' does not exist.",
                        )
                        return

                    confirmation = QMessageBox.question(
                        self,
                        "Confirmation",
                        f"Are you sure you want to delete the activity with code '{a}'?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No,
                    )

                    if confirmation == QMessageBox.Yes:
                        sql_delete = "DELETE FROM activity WHERE kd_activity = %s"
                        cursor.execute(sql_delete, (a,))
                        connection.commit()

                        QMessageBox.information(
                            self,
                            "Success",
                            "The activity has been successfully deleted.",
                        )

                        self.refreshSaveActivityTable()
                    else:
                        return
        finally:
            if connection:
                connection.close()

    def refreshSaveActivityTable(self):
        a = self.ui.lbSaveActivityGetCode.text()
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM activity WHERE kd_project=%s"
                    cursor.execute(sql, (a))
                    result = cursor.fetchall()

                    self.saveActivityModel.clear()
                    self.saveActivityModel.setHorizontalHeaderLabels(
                        ["Code", "Activity Name", "Description", ""]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    header_view = self.ui.tbSaveActivityTabel.horizontalHeader()

                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    self.ui.tbSaveActivityTabel.setColumnWidth(0, 100)
                    self.ui.tbSaveActivityTabel.setColumnWidth(1, 400)
                    self.ui.tbSaveActivityTabel.setColumnWidth(2, 500)
                    self.ui.tbSaveActivityTabel.setColumnWidth(3, 1)

                    for i, row_data in enumerate(result):
                        a = QStandardItem(row_data["kd_activity"])
                        b = QStandardItem(row_data["nama_activity"])
                        c = QStandardItem(row_data["keterangan"])
                        d = QStandardItem(row_data["kd_project"])

                        a.setFont(font)
                        b.setFont(font)
                        c.setFont(font)
                        d.setFont(font)

                        a.setTextAlignment(QtCore.Qt.AlignCenter)
                        b.setTextAlignment(QtCore.Qt.AlignCenter)
                        c.setTextAlignment(QtCore.Qt.AlignCenter)
                        d.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.saveActivityModel.appendRow([a, b, c, d])
                        self.ui.tbSaveActivityTabel.verticalHeader().setDefaultSectionSize(
                            50
                        )

        finally:
            if connection:
                connection.close()

    def searchDataActivity(self, cariData):
        getCode = self.ui.lbSaveActivityGetCode.text()
        cariData = self.ui.txtSaveActivityCari.text()

        if cariData.strip():
            try:
                connection = koneksi()
                if connection:
                    with connection.cursor() as cursor:
                        sql = "SELECT * FROM activity WHERE kd_project=%s AND (nama_activity LIKE %s OR kd_activity LIKE %s)"
                        cursor.execute(
                            sql,
                            (getCode, "%" + cariData + "%", "%" + cariData + "%"),
                        )
                        result = cursor.fetchall()

                        self.saveActivityModel.clear()
                        self.saveActivityModel.setHorizontalHeaderLabels(
                            ["Code", "Activity Name", "Description", ""]
                        )
                        font = QtGui.QFont()
                        font.setPointSize(14)

                        header_view = self.ui.tbSaveActivityTabel.horizontalHeader()

                        header_view.setStyleSheet(
                            "background-color: rgb(114, 159, 207);"
                        )
                        font = header_view.font()
                        font.setPointSize(18)
                        header_view.setFont(font)

                        self.ui.tbSaveActivityTabel.setColumnWidth(0, 100)
                        self.ui.tbSaveActivityTabel.setColumnWidth(1, 400)
                        self.ui.tbSaveActivityTabel.setColumnWidth(2, 500)
                        self.ui.tbSaveActivityTabel.setColumnWidth(3, 1)

                        for i, row_data in enumerate(result):
                            a = QStandardItem(row_data["kd_activity"])
                            b = QStandardItem(row_data["nama_activity"])
                            c = QStandardItem(row_data["keterangan"])
                            d = QStandardItem(row_data["kd_project"])

                            a.setFont(font)
                            b.setFont(font)
                            c.setFont(font)
                            d.setFont(font)

                            a.setTextAlignment(QtCore.Qt.AlignCenter)
                            b.setTextAlignment(QtCore.Qt.AlignCenter)
                            c.setTextAlignment(QtCore.Qt.AlignCenter)
                            d.setTextAlignment(QtCore.Qt.AlignCenter)

                            self.saveActivityModel.appendRow([a, b, c, d])
                            self.ui.tbSaveActivityTabel.verticalHeader().setDefaultSectionSize(
                                50
                            )

            finally:
                if connection:
                    connection.close()

        else:
            self.refreshSaveActivityTable()

    def keyReleaseEventActivity(self, event):
        if event.key() == QtCore.Qt.Key_Enter or event.key() == QtCore.Qt.Key_Return:
            self.searchDataActivity(event)
        else:
            super().keyReleaseEvent(event)

    def setActivityEmptyColumn(self):
        self.ui.txtSaveActivityKet.setText("")
        self.ui.txtSaveActivityNama.setText("")
        self.ui.txtSaveActivityCari.setText("")

    def setupActivityTable(self):
        self.ui.tbSaveActivityTabel.clicked.connect(self.displaySelectedActivityCode)

    def displaySelectedActivityCode(self):
        a = self.ui.tbSaveActivityTabel.selectedIndexes()
        if a:
            getRow = a[0].row()
            getCode = self.saveActivityModel.index(getRow, 0).data()
            getName = self.saveActivityModel.index(getRow, 1).data()
            getDesc = self.saveActivityModel.index(getRow, 2).data()
            self.ui.txtSaveActivityCode.setText(getCode)
            self.ui.txtSaveActivityNama.setText(getName)
            self.ui.txtSaveActivityKet.setText(getDesc)

    # END Activity Menu=====================================================================================================

    # START Integration Menu================================================================================================

    def updateLineEditX(self):
        x = self.ui.slIntegrationX.value()
        self.ui.txtIntegrationX.setValue(x)

    def updateLineEditY(self):
        y = self.ui.slIntegrationY.value()
        self.ui.txtIntegrationY.setValue(y)

    def updateLineEditZ(self):
        z = self.ui.slIntegrationZ.value()
        self.ui.txtIntegrationZ.setValue(z)

    def updateLineEditK(self):
        k = self.ui.slIntegrationK.value()
        self.ui.txtIntegrationK.setValue(k)

    def updateSliderX(self, value):
        self.ui.slIntegrationX.setValue(int(value))

    def updateSliderY(self, value):
        self.ui.slIntegrationY.setValue(int(value))

    def updateSliderZ(self, value):
        self.ui.slIntegrationZ.setValue(int(value))

    def updateSliderK(self, value):
        self.ui.slIntegrationK.setValue(int(value))

    def setCoordinates(self):
        a = self.ui.txtIntegrationCode.text()
        b = self.ui.txtIntegrationX.value()
        c = self.ui.txtIntegrationY.value()
        d = self.ui.txtIntegrationZ.value()
        e = self.ui.txtIntegrationK.value()
        f = 0
        g = self.ui.txtIntegrationKet.text()
        h = self.ui.lbIntegrationGetCode.text()

        if not g:
            QMessageBox.warning(self, "Warning", "Please input your description.")
            return

        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql_check = "SELECT * FROM koordinat WHERE kd_kor = %s"
                    cursor.execute(sql_check, (a,))
                    exist = cursor.fetchone()

                    if exist:
                        sql_update = "UPDATE koordinat SET x = %s, y = %s, z = %s, k = %s, delay = %s, keterangan = %s WHERE kd_kor = %s"
                        cursor.execute(sql_update, (b, c, d, e, f, g, a))
                        QMessageBox.information(
                            self,
                            "Success",
                            "Your Coordinats has been successfully updated.",
                        )
                    else:
                        sql_insert = "INSERT INTO koordinat (kd_kor, x, y, z, k, delay, keterangan, kd_activity) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"
                        cursor.execute(sql_insert, (a, b, c, d, e, f, g, h))
                        QMessageBox.information(
                            self,
                            "Success",
                            "Your Coordinats has been successfully saved.",
                        )

                    connection.commit()
                    self.refreshIntegrationTable()
                    self.getIntegrationCode()
                    # self.setIntegrationEmptyColumn()

        finally:
            if connection:
                connection.close()

    def deleteIntegration(self):
        a = self.ui.txtIntegrationCode.text()

        if not a:
            QMessageBox.warning(
                self, "Warning", "Please input the coordinat code to delete."
            )
            return

        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM koordinat WHERE kd_kor = %s"
                    cursor.execute(sql, (a,))
                    result = cursor.fetchone()

                    if not result:
                        QMessageBox.warning(
                            self,
                            "Warning",
                            f"The coordinat with code '{a}' does not exist.",
                        )
                        return

                    confirmation = QMessageBox.question(
                        self,
                        "Confirmation",
                        f"Are you sure you want to delete the coordinat with code '{a}'?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No,
                    )

                    if confirmation == QMessageBox.Yes:
                        sql_delete = "DELETE FROM koordinat WHERE kd_kor = %s"
                        cursor.execute(sql_delete, (a,))
                        connection.commit()

                        QMessageBox.information(
                            self,
                            "Success",
                            "The coordinat has been successfully deleted.",
                        )

                        self.refreshIntegrationTable()
                    else:
                        return
        finally:
            if connection:
                connection.close()

    def refreshIntegrationTable(self):
        a = self.ui.lbIntegrationGetCode.text()
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM koordinat WHERE kd_activity=%s"
                    cursor.execute(sql, (a,))
                    result = cursor.fetchall()

                    self.integrationModel.clear()
                    self.integrationModel.setHorizontalHeaderLabels(
                        ["C-Code", "X", "Y", "Z", "K", "Delay", "Description", ""]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    header_view = self.ui.tbIntegrationTabel.horizontalHeader()
                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    self.ui.tbIntegrationTabel.setColumnWidth(0, 100)
                    self.ui.tbIntegrationTabel.setColumnWidth(1, 100)
                    self.ui.tbIntegrationTabel.setColumnWidth(2, 100)
                    self.ui.tbIntegrationTabel.setColumnWidth(3, 100)
                    self.ui.tbIntegrationTabel.setColumnWidth(4, 100)
                    self.ui.tbIntegrationTabel.setColumnWidth(5, 100)
                    self.ui.tbIntegrationTabel.setColumnWidth(6, 400)
                    self.ui.tbIntegrationTabel.setColumnWidth(7, 100)

                    for row_data in result:
                        items = [
                            QStandardItem(str(row_data["kd_kor"])),
                            QStandardItem(str(row_data["x"])),
                            QStandardItem(str(row_data["y"])),
                            QStandardItem(str(row_data["z"])),
                            QStandardItem(str(row_data["k"])),
                            QStandardItem(str(row_data["delay"])),
                            QStandardItem(str(row_data["keterangan"])),
                            QStandardItem(str(row_data["kd_activity"])),
                        ]
                        for item in items:
                            item.setFont(font)
                            item.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.integrationModel.appendRow(items)

                    self.ui.tbIntegrationTabel.resizeColumnsToContents()
                    self.ui.tbIntegrationTabel.verticalHeader().setDefaultSectionSize(
                        50
                    )

        finally:
            if connection:
                connection.close()

    def searchDataIntegration(self, cariData):
        a = self.ui.lbIntegrationGetCode.text()
        cariData = self.ui.txtIntegrationCari.text()

        if cariData.strip():
            try:
                connection = koneksi()
                if connection:
                    with connection.cursor() as cursor:
                        sql = "SELECT * FROM koordinat WHERE kd_activity=%s AND (keterangan LIKE %s OR kd_kor LIKE %s)"
                        cursor.execute(
                            sql,
                            (a, "%" + cariData + "%", "%" + cariData + "%"),
                        )
                        result = cursor.fetchall()

                        self.integrationModel.clear()
                        self.integrationModel.setHorizontalHeaderLabels(
                            ["C-Code", "X", "Y", "Z", "K", "Delay", "Description", ""]
                        )
                        font = QtGui.QFont()
                        font.setPointSize(14)

                        header_view = self.ui.tbIntegrationTabel.horizontalHeader()
                        header_view.setStyleSheet(
                            "background-color: rgb(114, 159, 207);"
                        )
                        font = header_view.font()
                        font.setPointSize(18)
                        header_view.setFont(font)

                        self.ui.tbIntegrationTabel.setColumnWidth(0, 100)
                        self.ui.tbIntegrationTabel.setColumnWidth(1, 100)
                        self.ui.tbIntegrationTabel.setColumnWidth(2, 100)
                        self.ui.tbIntegrationTabel.setColumnWidth(3, 100)
                        self.ui.tbIntegrationTabel.setColumnWidth(4, 100)
                        self.ui.tbIntegrationTabel.setColumnWidth(5, 100)
                        self.ui.tbIntegrationTabel.setColumnWidth(6, 400)
                        self.ui.tbIntegrationTabel.setColumnWidth(7, 100)

                        for row_data in result:
                            items = [
                                QStandardItem(str(row_data["kd_kor"])),
                                QStandardItem(str(row_data["x"])),
                                QStandardItem(str(row_data["y"])),
                                QStandardItem(str(row_data["z"])),
                                QStandardItem(str(row_data["k"])),
                                QStandardItem(str(row_data["delay"])),
                                QStandardItem(str(row_data["keterangan"])),
                                QStandardItem(str(row_data["kd_activity"])),
                            ]
                            for item in items:
                                item.setFont(font)
                                item.setTextAlignment(QtCore.Qt.AlignCenter)

                            self.integrationModel.appendRow(items)

                        self.ui.tbIntegrationTabel.resizeColumnsToContents()
                        self.ui.tbIntegrationTabel.verticalHeader().setDefaultSectionSize(
                            50
                        )

            finally:
                if connection:
                    connection.close()

        else:
            self.refreshIntegrationTable()

    def getIntegrationCode(self):
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT MAX(kd_kor) FROM koordinat"
                    cursor.execute(sql)
                    result = cursor.fetchone()

                    if result["MAX(kd_kor)"]:
                        latest_code = result["MAX(kd_kor)"]
                        kode_int = int(latest_code[1:]) + 1
                        new_kode = f"C{kode_int:04d}"
                    else:
                        new_kode = "C0001"

                    self.ui.txtIntegrationCode.setText(new_kode)
                    return new_kode
        finally:
            if connection:
                connection.close()

    def setupIntegrationTable(self):
        self.ui.tbIntegrationTabel.clicked.connect(self.displaySelectedIntegration)

    def displaySelectedIntegration(self):
        a = self.ui.tbIntegrationTabel.selectedIndexes()
        if a:
            getRow = a[0].row()
            getCode = self.integrationModel.index(getRow, 0).data()
            getX = int(self.integrationModel.index(getRow, 1).data())
            getY = int(self.integrationModel.index(getRow, 2).data())
            getZ = int(self.integrationModel.index(getRow, 3).data())
            getK = int(self.integrationModel.index(getRow, 4).data())
            getDelay = int(self.integrationModel.index(getRow, 5).data())
            getDesc = self.integrationModel.index(getRow, 6).data()

            self.ui.txtIntegrationCode.setText(getCode)
            self.ui.txtIntegrationKet.setText(getDesc)
            self.ui.txtIntegrationDelay.setNum(getDelay)

            self.ui.txtIntegrationX.setValue(getX)
            self.ui.txtIntegrationY.setValue(getY)
            self.ui.txtIntegrationZ.setValue(getZ)
            self.ui.txtIntegrationK.setValue(getK)

            self.ui.slIntegrationX.setValue(getX)
            self.ui.slIntegrationY.setValue(getY)
            self.ui.slIntegrationZ.setValue(getZ)
            self.ui.slIntegrationK.setValue(getK)

            sendSerial(getX, getY, getK, getZ, getDelay)

    def keyReleaseEventIntegration(self, event):
        if event.key() == QtCore.Qt.Key_Enter or event.key() == QtCore.Qt.Key_Return:
            self.searchDataIntegration(event)
        else:
            super().keyReleaseEvent(event)

    def setIntegrationEmptyColumn(self):
        self.ui.txtIntegrationKet.setText("")
        self.ui.txtIntegrationDelay.setText("0")
        self.ui.txtIntegrationX.setValue(0)
        self.ui.txtIntegrationY.setValue(0)
        self.ui.txtIntegrationZ.setValue(0)
        self.ui.txtIntegrationK.setValue(0)

        self.ui.slIntegrationX.setValue(0)
        self.ui.slIntegrationY.setValue(0)
        self.ui.slIntegrationZ.setValue(0)
        self.ui.slIntegrationK.setValue(0)

        sendSerial(0, 0, 0, 0, 0)

    def setIntegrationEmptyColumn1(self):
        self.ui.txtIntegrationKet.setText("")
        self.ui.txtIntegrationDelay.setText("0")
        self.ui.txtIntegrationX.setValue(0)
        self.ui.txtIntegrationY.setValue(0)
        self.ui.txtIntegrationZ.setValue(0)
        self.ui.txtIntegrationK.setValue(0)

        self.ui.slIntegrationX.setValue(0)
        self.ui.slIntegrationY.setValue(0)
        self.ui.slIntegrationZ.setValue(0)
        self.ui.slIntegrationK.setValue(0)

    # END Integration Menu ==================================================================================

    # START Open Project ====================================================================================

    def refreshOpenProjectTable(self):
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM project"
                    cursor.execute(sql)
                    result = cursor.fetchall()

                    self.openProjectModel.clear()
                    self.openProjectModel.setHorizontalHeaderLabels(
                        ["P-Code", "Project Name", "Description"]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    header_view = self.ui.tbOpenProjectTabel.horizontalHeader()

                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    self.ui.tbOpenProjectTabel.setColumnWidth(0, 100)
                    self.ui.tbOpenProjectTabel.setColumnWidth(1, 400)
                    self.ui.tbOpenProjectTabel.setColumnWidth(2, 800)

                    for i, row_data in enumerate(result):
                        a = QStandardItem(row_data["kd_project"])
                        b = QStandardItem(row_data["nama_project"])
                        c = QStandardItem(row_data["keterangan"])

                        a.setFont(font)
                        b.setFont(font)
                        c.setFont(font)

                        a.setTextAlignment(QtCore.Qt.AlignCenter)
                        b.setTextAlignment(QtCore.Qt.AlignCenter)
                        c.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.openProjectModel.appendRow([a, b, c])
                        self.ui.tbOpenProjectTabel.verticalHeader().setDefaultSectionSize(
                            50
                        )

        finally:
            if connection:
                connection.close()

    def setOpenProjectEmptyColumn(self):
        self.ui.txtOpenProjectGetData.setText("")
        self.ui.txtOpenProjectGetName.setText("")

    def deleteProject(self):
        a = self.ui.txtOpenProjectGetData.text()

        if not a:
            QMessageBox.warning(
                self, "Warning", "Please input the activity code to delete."
            )
            return

        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql_select = "SELECT * FROM project WHERE kd_project = %s"
                    cursor.execute(sql_select, (a,))
                    result = cursor.fetchone()

                    if not result:
                        QMessageBox.warning(
                            self,
                            "Warning",
                            f"The project with code '{a}' does not exist.",
                        )
                        return

                    confirmation = QMessageBox.question(
                        self,
                        "Confirmation",
                        f"Are you sure you want to delete the project with code '{a}'?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.No,
                    )

                    if confirmation == QMessageBox.Yes:
                        sql_delete = "DELETE FROM project WHERE kd_project = %s"
                        cursor.execute(sql_delete, (a,))
                        connection.commit()

                        QMessageBox.information(
                            self,
                            "Success",
                            "The project has been successfully deleted.",
                        )
                        self.setOpenProjectEmptyColumn()
                        self.refreshOpenProjectTable()
                    else:
                        return
        finally:
            if connection:
                connection.close()

    def searchOpenProjectData(self, cariData):
        cariData = self.ui.txtOpenProjectCari.text()

        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM project WHERE nama_project LIKE %s OR kd_project LIKE %s"
                    cursor.execute(
                        sql,
                        ("%" + cariData + "%", "%" + cariData + "%"),
                    )
                    result = cursor.fetchall()

                    self.openProjectModel.clear()
                    self.openProjectModel.setHorizontalHeaderLabels(
                        ["Project Code", "Project Name", "Description"]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    header_view = self.ui.tbOpenProjectTabel.horizontalHeader()

                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    for i, row_data in enumerate(result):
                        a = QStandardItem(row_data["kd_project"])
                        b = QStandardItem(row_data["nama_project"])
                        c = QStandardItem(row_data["keterangan"])

                        a.setFont(font)
                        b.setFont(font)
                        c.setFont(font)

                        a.setTextAlignment(QtCore.Qt.AlignCenter)
                        b.setTextAlignment(QtCore.Qt.AlignCenter)
                        c.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.openProjectModel.appendRow([a, b, c])
                        self.ui.tbOpenProjectTabel.resizeColumnsToContents()
                        self.ui.tbOpenProjectTabel.verticalHeader().setDefaultSectionSize(
                            50
                        )

        finally:
            if connection:
                connection.close()

    def keyOpenReleaseEvent(self, event):
        if event.key() == QtCore.Qt.Key_Enter or event.key() == QtCore.Qt.Key_Return:
            self.searchOpenProjectData(event)
        else:
            super().keyReleaseEvent(event)

    def setupOpenProjectTable(self):
        self.ui.tbOpenProjectTabel.clicked.connect(self.displaySelectedProjectCode)

    def displaySelectedProjectCode(self):
        a = self.ui.tbOpenProjectTabel.selectedIndexes()
        if a:
            getRow = a[0].row()
            getCode = self.openProjectModel.index(getRow, 0).data()
            getName = self.openProjectModel.index(getRow, 1).data()
            self.ui.txtOpenProjectGetData.setText(getCode)
            self.ui.txtOpenProjectGetName.setText(getName)

    def goToActivityMenu2(self):
        a = self.ui.txtOpenProjectGetData.text()
        b = self.ui.txtOpenProjectGetName.text()
        if a:
            self.ui.lbSaveActivityGetCode.setText(a)
            self.ui.lbCreateAcvtivityGetProjectName.setText(b)
            self.setActivity()

    # END Open Project ===================================================================================================
                
    # START Run Program ==================================================================================================

    def setRunProgramEmptyColumn(self):
        self.ui.txtOpenProjectGetData.setText("")
        self.ui.txtOpenProjectGetName.setText("")

    def refreshRunningProjectData_1(self):
        d = self.ui.txtRunningProjectGetActivity.text()
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM koordinat WHERE kd_activity=%s"
                    cursor.execute(sql, (d,))
                    result = cursor.fetchall()

                    self.runningModel_1.clear()
                    self.runningModel_1.setHorizontalHeaderLabels(     
                    ["Description", "X", "Y", "Z", "K", "Delay", "", "C-Code"]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    header_view = self.ui.tbRunningProjectTabel_1.horizontalHeader()
                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    self.ui.tbRunningProjectTabel_1.setColumnWidth(0, 500)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(1, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(2, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(3, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(4, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(5, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(6, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(7, 100)

                    for row_data in result:
                        items = [
                            QStandardItem(str(row_data["keterangan"])),
                            QStandardItem(str(row_data["x"])),
                            QStandardItem(str(row_data["y"])),
                            QStandardItem(str(row_data["z"])),
                            QStandardItem(str(row_data["k"])),
                            QStandardItem(str(row_data["delay"])),
                            QStandardItem(str(row_data["kd_activity"])),
                            QStandardItem(str(row_data["kd_kor"])),
                        ]
                        for item in items:
                            item.setFont(font)
                            item.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.runningModel_1.appendRow(items)

                    self.ui.tbRunningProjectTabel_1.verticalHeader().setDefaultSectionSize(
                        50
                    )

                    self.runningModel_2.setHorizontalHeaderLabels(
                    ["Description", "X", "Y", "K", "Z", "Delay", "", "C-Code"]
                    )

                    font_2 = QtGui.QFont()
                    font_2.setPointSize(14)
                    header_view_2 = self.ui.tbRunningProjectTabel_2.horizontalHeader()
                    header_view_2.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font_2 = header_view_2.font()
                    font_2.setPointSize(18)
                    header_view_2.setFont(font_2)

                    self.ui.tbRunningProjectTabel_2.setColumnWidth(0, 600)
                    self.ui.tbRunningProjectTabel_2.setColumnWidth(1, 100)
                    self.ui.tbRunningProjectTabel_2.setColumnWidth(2, 100)
                    self.ui.tbRunningProjectTabel_2.setColumnWidth(3, 100)
                    self.ui.tbRunningProjectTabel_2.setColumnWidth(4, 100)
                    self.ui.tbRunningProjectTabel_2.setColumnWidth(5, 100)
                    self.ui.tbRunningProjectTabel_2.setColumnWidth(6, 100)
                    self.ui.tbRunningProjectTabel_2.setColumnWidth(7, 100)

        finally:
            if connection:
                connection.close()

    
    def rowClicked(self, index):
        self.resetColorRunningModel_1()

        selected_row_data = []
        for column_index in range(self.runningModel_1.columnCount()):
            item = self.runningModel_1.item(index.row(), column_index)
            selected_row_data.append(item.text())

        if selected_row_data not in self.stacked_data: 
            self.stacked_data.append(selected_row_data)
            self.highlight_selected_row(index.row(), "blue")
        else:
            print("zero")


    def moveToTable2(self):
        if self.stacked_data:
            for data in self.stacked_data:
                self.add_row_to_table_2(data)

            self.stacked_data.clear()
            self.resetColorRunningModel_1()


    def removeFromTable2(self):
        selected_index = self.ui.tbRunningProjectTabel_2.currentIndex()
        if selected_index.isValid():
            self.runningModel_2.removeRow(selected_index.row())


    def resetColorRunningModel_1(self):
        for row in range(self.runningModel_1.rowCount()):
            for column in range(self.runningModel_1.columnCount()):
                item = self.runningModel_1.item(row, column)
                if item is not None:
                    item.setBackground(QColor("white"))

    def resetColorRunningModel_2(self, row_index):
        for column_index in range(self.runningModel_2.columnCount()):
            item = self.runningModel_2.item(row_index, column_index)
            if item is not None:
                item.setBackground(QColor("white"))

    def highlight_selected_row(self, row_index, color):
        for column_index in range(self.runningModel_1.columnCount()):
            item = self.runningModel_1.item(row_index, column_index)
            item.setBackground(QColor(color))


    def add_row_to_table_2(self, row_data):
        code = QStandardItem(row_data[0])  
        x = QStandardItem(row_data[1])  
        y = QStandardItem(row_data[2])  
        z = QStandardItem(row_data[3])  
        k = QStandardItem(row_data[4])  
        delay = QStandardItem(row_data[5])  
        ket = QStandardItem(row_data[6])  
        pr = QStandardItem(row_data[7]) 
        
        font_2 = QtGui.QFont()
        font_2.setPointSize(14)
        for item in [ket, x, y, z, k, delay, pr, code]:
            item.setFont(font_2)  
            item.setTextAlignment(QtCore.Qt.AlignCenter)

        items = [code, x, y, k, z, delay, pr, ket] + [QStandardItem('') for _ in range(0)]
        
        self.runningModel_2.appendRow(items)

    def startSendingSteps(self):
        self.current_step = -1
        self.timer = QTimer()
        self.timer.timeout.connect(self.sendNextStep)
        self.timer.start(2000)


    def highlightRow(self, row_index, color):
        view = self.ui.tbRunningProjectTabel_2
        for column_index in range(self.runningModel_2.columnCount()):
            item = self.runningModel_2.item(row_index, column_index)
            
            if item is not None:
                # self.timer.start(delay)  
                item.setBackground(color)
                view.viewport().update()  
 

    def stopSendingSteps(self):
        self.timer.stop()
        sendSerial(0,0,0,0,0)


    # def sendLoopStep(self):
    #     self.handleDuplicateCoordinates()
        
    #     row_count = self.runningModel_2.rowCount()
    #     delay = int(self.runningModel_2.data(self.runningModel_2.index(0, 5)))
        
    #     if self.current_step < row_count:
    #         if self.current_step >= 0:
    #             self.resetColorRunningModel_2(self.current_step)

    #         self.current_step += 1
    #         if self.current_step < row_count and delay > 0:
    #             self.timer.start(delay)
    #             self.highlightRow(self.current_step, QColor("yellow"))
            
    #         elif delay == 0:
    #             self.timer.start(2000)
    #             self.highlightRow(self.current_step, QColor("yellow"))
            

    #         data_row = []
    #         for column in range(1, 6):
    #             item = self.runningModel_2.item(self.current_step, column)
    #             if item is not None and item.text():
    #                 data_row.append(int(item.text()))
    #             else:
    #                 data_row.append(0)
    #         sendSerial(*data_row)
    #     else:
    #         self.current_step = -1  
    #         self.setupTimer()


    def sendLoopStep(self):
        self.handleDuplicateCoordinates()
        
        row_count = self.runningModel_2.rowCount()
        delay = int(self.runningModel_2.data(self.runningModel_2.index(0, 5)))
        
        if self.current_step < row_count:
            if self.current_step >= 0:
                self.resetColorRunningModel_2(self.current_step)

            self.current_step += 1
            
            if delay > 1:
                self.timer.start(delay)
            else:
                self.timer.start(2000)

            if self.current_step < row_count:
                self.highlightRow(self.current_step, QColor("yellow"))
                
            data_row = []
            for column in range(1, 6):
                item = self.runningModel_2.item(self.current_step, column)
                if item is not None and item.text():
                    data_row.append(int(item.text()))
                else:
                    data_row.append(0)
            sendSerial(*data_row)
        else:
            self.current_step = -1  
            self.setupTimer()

        
    def sendNextStep(self):
        row_count = self.runningModel_2.rowCount()
        if self.current_step < row_count:
            if self.current_step >= 0:
                self.resetColorRunningModel_2(self.current_step)

            self.current_step += 1
            if self.current_step < row_count:
                self.highlightRow(self.current_step, QColor("yellow"))

            data_row = []
            for column in range(1, 6):
                item = self.runningModel_2.item(self.current_step, column)
                if item is not None and item.text():
                    data_row.append(int(item.text()))
                else:
                    data_row.append(0)
            sendSerial(*data_row)
        else:
            self.stopSendingSteps()  

        if self.current_step == row_count-1:
            self.ui.btRunningProjectStatus.setText("Has Finished")
            self.ui.btRunningProjectRun.setText("Run")
            self.resetColorRunningModel_2(self.current_step)
            self.timer.stop()
            sendSerial(0,0,0,0,0) 


    def handleDuplicateCoordinates(self):
        row_count = self.runningModel_2.rowCount()
        columns_count = self.runningModel_2.columnCount()

        current_row = 0
        while current_row < row_count - 1:
            current_data = []
            next_data = []

            for column in range(min(columns_count, 8)):  
                current_item = self.runningModel_2.item(current_row, column)
                if current_item:
                    current_data.append(current_item.text())
                else:
                    current_data.append("")

                next_item = self.runningModel_2.item(current_row + 1, column)
                if next_item:
                    next_data.append(next_item.text())
                else:
                    next_data.append("")

            # print("CD:", current_data)
            # print("ND:", next_data)

            if current_data[:8] == next_data[:8]:
                if not self.dataCopied(current_row + 1):
                    print("Inserting row at:", current_row + 1)
                    self.runningModel_2.insertRow(current_row + 1)

                    for column in range(columns_count):
                        current_item = QStandardItem(current_data[column])
                        self.runningModel_2.setItem(current_row + 1, column, current_item)

                    row_count += 1

            current_row += 1
            

    def dataCopied(self, row):
        columns_count = min(self.runningModel_2.columnCount(), 8)

        current_data = []
        for column in range(columns_count):
            item = self.runningModel_2.item(row, column)
            if item and item.text():
                current_data.append(item.text())

        if len(current_data) == columns_count:
            for column in range(columns_count):
                if current_data[column] != self.runningModel_2.item(row - 1, column).text():
                    return False
            return True
        return False
 

    def setupTimer(self):
        self.current_step = -1
        self.timer = QTimer()
        self.timer.timeout.connect(self.sendLoopStep)
        self.timer.start(2000)   

    def resetColorRunningModel_2(self, row):
        for column in range(0, 8):
            item = self.runningModel_2.item(row, column)
            if item is not None:
                item.setBackground(QColor("white"))

    def highlightRowLoop(self, row, color):
        for column in range(0, 8):
            item = self.runningModel_2.item(row, column)
            if item is not None:
                item.setBackground(color)

    def moveDataToNextColumn(self, data_row):
        for i in range(len(data_row)):
            next_column_index = (self.current_step + 1) % self.runningModel_2.columnCount()
            item = QStandardItem(data_row[i])
            self.runningModel_2.setItem(self.current_step, next_column_index, item)


    def actionSendeingSteps(self):
        self.pauseProcess()  
        sendSerial(0,0,0,0,0)
        reply = QMessageBox.question(self, 'Caution', 
                        "If you click this button, the process will be stopped. Do you agree?",
                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.pauseProcess()  
            sendSerial(0,0,0,0,0)
            self.ui.btRunningProjectStatus.setText("Has Stopped")
            self.resetColorRunningModel_2(self.current_step)


    def toggleSendingSteps(self):
        if self.ui.btRunningProjectRun.text() == "Run":
            self.ui.btRunningProjectRun.setText("Pause")
            self.startSendingSteps()
            self.ui.btRunningProjectStatus.setText("Has Running")
        elif self.ui.btRunningProjectRun.text() == "Pause":
            self.ui.btRunningProjectRun.setText("Resume")
            self.pauseProcess()
            self.ui.btRunningProjectStatus.setText("Has Paused")
        elif self.ui.btRunningProjectRun.text() == "Resume":
            self.ui.btRunningProjectRun.setText("Pause")
            self.resumeProcess()
            self.ui.btRunningProjectStatus.setText("Has Running")
        elif self.ui.btRunningProjectStop.text() == "Stop":
            self.ui.btRunningProjectRun.setText("Run")
            self.actionSendeingSteps()
            self.ui.btRunningProjectStatus.setText("Has Stopped")
            if self.current_step < self.runningModel_2.rowCount():
                self.highlightRow(self.current_step, QColor("yellow"))

    def toggleSendingLoopSteps(self):
        if self.ui.btRunningProjectLoop.text() == "Loop":
            self.setupTimer()
            self.ui.btRunningProjectStatus.setText("Has Looping")
        elif self.ui.btRunningProjectStop.text() == "Stop":
            self.actionSendeingSteps()
            self.ui.btRunningProjectStatus.setText("Has Stopped")
            if self.current_step < self.runningModel_2.rowCount():
                self.highlightRow(self.current_step, QColor("yellow"))

    def pauseProcess(self):
        self.timer.stop()

    def resumeProcess(self):
        self.timer.start() 

    def executeDelay(delay):
        time.sleep(delay * 1000)

    def sendWithDelay(data_row):
        x, y, k, z, delay = map(int, data_row[1:6])  
        sendSerial(x, y, k, z, delay) 
          
            
    # END Run Program ====================================================================================================

    # START Choose Project ===============================================================================================
    
    def refreshChooseProjectTable(self):
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM project"
                    cursor.execute(sql)
                    result = cursor.fetchall()

                    self.chooseProjectModel.clear()
                    self.chooseProjectModel.setHorizontalHeaderLabels(
                        ["Project Code", "Project Name", "Description"]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    self.ui.tbChooseProjectTabel.setColumnWidth(0, 100)
                    self.ui.tbChooseProjectTabel.setColumnWidth(1, 400)
                    self.ui.tbChooseProjectTabel.setColumnWidth(2, 700)

                    header_view = self.ui.tbChooseProjectTabel.horizontalHeader()

                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    for i, row_data in enumerate(result):
                        a = QStandardItem(row_data["kd_project"])
                        b = QStandardItem(row_data["nama_project"])
                        c = QStandardItem(row_data["keterangan"])

                        a.setFont(font)
                        b.setFont(font)
                        c.setFont(font)

                        a.setTextAlignment(QtCore.Qt.AlignCenter)
                        b.setTextAlignment(QtCore.Qt.AlignCenter)
                        c.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.chooseProjectModel.appendRow([a, b, c])
                        self.ui.tbChooseProjectTabel.verticalHeader().setDefaultSectionSize(
                            50
                        )

        finally:
            if connection:
                connection.close()

    def searchChooseProjectData(self, cariData):
        cariData = self.ui.txtChooseProjectCari.text()

        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM project WHERE nama_project LIKE %s OR kd_project LIKE %s"
                    cursor.execute(
                        sql,
                        ("%" + cariData + "%", "%" + cariData + "%"),
                    )
                    result = cursor.fetchall()

                    self.chooseProjectModel.clear()
                    self.chooseProjectModel.setHorizontalHeaderLabels(
                        ["Project Code", "Project Name", "Description"]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    self.ui.tbChooseProjectTabel.setColumnWidth(0, 100)
                    self.ui.tbChooseProjectTabel.setColumnWidth(1, 400)
                    self.ui.tbChooseProjectTabel.setColumnWidth(2, 700)
                    
                    header_view = self.ui.tbChooseProjectTabel.horizontalHeader()

                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    for i, row_data in enumerate(result):
                        a = QStandardItem(row_data["kd_project"])
                        b = QStandardItem(row_data["nama_project"])
                        c = QStandardItem(row_data["keterangan"])

                        a.setFont(font)
                        b.setFont(font)
                        c.setFont(font)

                        a.setTextAlignment(QtCore.Qt.AlignCenter)
                        b.setTextAlignment(QtCore.Qt.AlignCenter)
                        c.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.chooseProjectModel.appendRow([a, b, c])
                        self.ui.tbChooseProjectTabel.verticalHeader().setDefaultSectionSize(
                            50
                        )

        finally:
            if connection:
                connection.close()

    def keyChooseProjectReleaseEvent(self, event):
        if event.key() == QtCore.Qt.Key_Enter or event.key() == QtCore.Qt.Key_Return:
            self.searchChooseProjectData(event)
        else:
            super().keyReleaseEvent(event)

    def setupChooseProjectTable(self):
        self.ui.tbChooseProjectTabel.clicked.connect(self.displaySelectedChooseProject)

    def displaySelectedChooseProject(self):
        a = self.ui.tbChooseProjectTabel.selectedIndexes()
        if a:
            getRow = a[0].row()
            getCode = self.chooseProjectModel.index(getRow, 0).data()
            getName = self.chooseProjectModel.index(getRow, 1).data()
            self.ui.lbChooseProjectGetCode.setText(getCode)
            self.ui.lbChooseProjectGetName.setText(getName)

    def goToRunProjectGetProject(self):
        a = self.ui.lbChooseProjectGetCode.text()
        b = self.ui.lbChooseProjectGetName.text()
        if a:
            self.ui.txtRunningProjectGetProject.setText(a)
            self.ui.txtRunProjectName.setText(b)
            self.ui.lbChooseActivityGetCodeProject.setText(a)
            self.setRunProject()


    # END Choose Project =================================================================================================
         

    # START Choose Activity ==============================================================================================
    
    def refreshChooseActivityTable(self):
        eii = self.ui.lbChooseActivityGetCodeProject.text()
        try:
            connection = koneksi()
            if connection:
                with connection.cursor() as cursor:
                    sql = "SELECT * FROM activity WHERE kd_project=%s"
                    cursor.execute(sql, (str(eii)))
                    result = cursor.fetchall()

                    self.chooseActivityModel.clear()
                    self.chooseActivityModel.setHorizontalHeaderLabels(
                        ["Code", "Activity Name", "Description", "."]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    header_view = self.ui.tbChooseActivityTabel.horizontalHeader()
                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font.setPointSize(18)
                    header_view.setFont(font)

                    self.ui.tbChooseActivityTabel.setColumnWidth(0, 100)
                    self.ui.tbChooseActivityTabel.setColumnWidth(1, 400)
                    self.ui.tbChooseActivityTabel.setColumnWidth(2, 700)
                    self.ui.tbChooseActivityTabel.setColumnWidth(3, 1)

                    for i, row_data in enumerate(result):
                        a = QStandardItem(row_data["kd_activity"])
                        b = QStandardItem(row_data["nama_activity"])
                        c = QStandardItem(row_data["keterangan"])
                        d = QStandardItem(row_data["kd_project"])

                        a.setFont(font)
                        b.setFont(font)
                        c.setFont(font)
                        d.setFont(font)

                        a.setTextAlignment(QtCore.Qt.AlignCenter)
                        b.setTextAlignment(QtCore.Qt.AlignCenter)
                        c.setTextAlignment(QtCore.Qt.AlignCenter)
                        d.setTextAlignment(QtCore.Qt.AlignCenter)

                        self.chooseActivityModel.appendRow([a, b, c, d])
                        self.ui.tbChooseActivityTabel.verticalHeader().setDefaultSectionSize(50)

        finally:
            if connection:
                connection.close()

    def searchChooseActivity(self, cariData):
        getCode = self.ui.lbChooseActivityGetCode.text()
        cariData = self.ui.txtChooseActivityCari.text()

        if cariData.strip():
            try:
                connection = koneksi()
                if connection:
                    with connection.cursor() as cursor:
                        sql = "SELECT * FROM activity WHERE kd_project=%s AND (nama_activity LIKE %s OR kd_activity LIKE %s)"
                        cursor.execute(
                            sql,
                            (getCode, "%" + cariData + "%", "%" + cariData + "%"),
                        )
                        result = cursor.fetchall()

                        self.chooseActivityModel.clear()
                        self.chooseActivityModel.setHorizontalHeaderLabels(
                            ["Code", "Activity Name", "Description", ""]
                        )
                        font = QtGui.QFont()
                        font.setPointSize(14)

                        header_view = self.ui.tbChooseActivityTabel.horizontalHeader()

                        header_view.setStyleSheet(
                            "background-color: rgb(114, 159, 207);"
                        )
                        font = header_view.font()
                        font.setPointSize(18)
                        header_view.setFont(font)

                        self.ui.tbChooseActivityTabel.setColumnWidth(0, 100)
                        self.ui.tbChooseActivityTabel.setColumnWidth(1, 400)
                        self.ui.tbChooseActivityTabel.setColumnWidth(2, 700)
                        self.ui.tbChooseActivityTabel.setColumnWidth(3, 1)

                        for i, row_data in enumerate(result):
                            a = QStandardItem(row_data["kd_activity"])
                            b = QStandardItem(row_data["nama_activity"])
                            c = QStandardItem(row_data["keterangan"])
                            d = QStandardItem(row_data["kd_project"])

                            a.setFont(font)
                            b.setFont(font)
                            c.setFont(font)
                            d.setFont(font)

                            a.setTextAlignment(QtCore.Qt.AlignCenter)
                            b.setTextAlignment(QtCore.Qt.AlignCenter)
                            c.setTextAlignment(QtCore.Qt.AlignCenter)
                            d.setTextAlignment(QtCore.Qt.AlignCenter)

                            self.chooseActivityModel.appendRow([a, b, c, d])
                            self.ui.tbChooseActivityTabel.verticalHeader().setDefaultSectionSize(
                                50
                            )

            finally:
                if connection:
                    connection.close()

        else:
            self.refreshChooseActivityTable()

    def keyReleaseEventShowActivity(self, event):
        if event.key() == QtCore.Qt.Key_Enter or event.key() == QtCore.Qt.Key_Return:
            self.searchDataShowActivity(event)
        else:
            super().keyReleaseEvent(event)

    def setupChooseActivityTable(self):
        self.ui.tbChooseActivityTabel.clicked.connect(self.displaySelectedShowActivityCode)

    def displaySelectedShowActivityCode(self):
        a = self.ui.tbChooseActivityTabel.selectedIndexes()
        if a:
            getRow = a[0].row()
            getCode = self.chooseActivityModel.index(getRow, 0).data()
            getName = self.chooseActivityModel.index(getRow, 1).data()
            getPro = self.chooseActivityModel.index(getRow, 3).data()
            self.ui.lbChooseActivityGetCode.setText(getCode)
            self.ui.lbChooseActivityGetName.setText(getName)
            self.ui.lbChooseActivityGetCodeProject.setText(getPro)

    def goToRunActivityGetProject(self):
        a = self.ui.lbChooseActivityGetCode.text()
        b = self.ui.lbChooseActivityGetName.text()
        if a:
            self.ui.txtRunningProjectGetActivity.setText(a)
            self.ui.txtRunProjectActivity.setText(b)
            
            self.setRunProject()
            self.refreshRunningProjectData_1()
            self.refreshChooseActivityTable()


    # END Choose Activity ================================================================================================

    # Bagian set up serial ===============================================================================================

    def releasedSlider(self):
        global currentX, currentY, currentZ, currentK, delay
        valueX = self.ui.slIntegrationX.value()
        valueY = self.ui.slIntegrationY.value()
        valueZ = self.ui.slIntegrationZ.value()
        valueK = self.ui.slIntegrationK.value()

        if (valueX != currentX or valueY != currentY or valueZ != currentZ or valueK != currentK):
            sendSerial(valueX, valueY, valueK, valueZ, delay)
            currentX, currentY, currentZ, currentK, delay = valueX, valueY, valueZ, valueK, delay


    def releasedText(self):
        global delay
        valueX = self.ui.txtIntegrationX.value()
        valueY = self.ui.txtIntegrationY.value()
        valueZ = self.ui.txtIntegrationZ.value()
        valueK = self.ui.txtIntegrationK.value()
        sendSerial(valueX, valueY, valueK, valueZ, delay)


    def sendDataSerial(self):
        data_list = []
        for row in range(self.runningModel_2.rowCount()):
            data_row = []
            for column in range(self.runningModel_2.columnCount()):
                item = self.runningModel_2.item(row, column)
                if item is not None:
                    data_row.append(item.text())
            data_list.append(data_row)

        for data_row in data_list:
            if len(data_row) == 5:
                sendSerial(*data_row) 

def writeSerial(data):
    ser.write(bytes(data, 'utf-8'))
    time.sleep(0.05)

def sendSerial(x, y, k, z, delay):
    data = "config:%d,%d,%d,%d,%d;" % (x, y, k, z, delay)
    writeSerial(data)
    print(data)

    
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.showMaximized()
    sys.exit(app.exec_())

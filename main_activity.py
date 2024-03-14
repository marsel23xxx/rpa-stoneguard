import sys

import pymysql
import serial
import time
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QStandardItem, QStandardItemModel
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
            # 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
            # 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
            # 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
        )
        return connection
    except pymysql.err.OperationalError as e:
        QMessageBox.critical(None, "Kesalahan", f"Kesalahan koneksi: {e}")
        return None

ser = serial.Serial(port='/dev/ttyUSB0', baudrate=9600, timeout=.1)
ser.flush()

class MainWindow(QtWidgets.QMainWindow):
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

        self.ui.txtIntegrationX.textChanged.connect(self.updateSliderX)
        self.ui.txtIntegrationY.textChanged.connect(self.updateSliderY)
        self.ui.txtIntegrationZ.textChanged.connect(self.updateSliderZ)
        self.ui.txtIntegrationK.textChanged.connect(self.updateSliderK)

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

        # Komponen bagian show project
        self.ui.tbChooseProjectTabel.setModel(self.chooseProjectModel)
        self.ui.txtChooseProjectCari.keyReleaseEvent = self.searchChooseProjectData
        self.refreshChooseProjectTable()
        self.setupChooseProjectTable()
        self.ui.btChooseProjectOpen.clicked.connect(self.goToRunProjectGetProject)

        # Komponen bagian show activity
        self.refreshChooseActivityTable()
        self.ui.tbChooseActivityTabel.setModel(self.chooseActivityModel)
        self.ui.txtChooseActivityCari.keyReleaseEvent = self.searchChooseActivity
        self.setupChooseActivityTable()
        self.ui.btChooseActivityOpen.clicked.connect(self.goToRunActivityGetProject)
        self.ui.btChooseActivityBack.clicked.connect(self.setRunProject)


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
            self.setIntegrationEmptyColumn()
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

    def updateLineEditX(self, value):
        self.ui.txtIntegrationX.setText(str(value))

    def updateLineEditY(self, value):
        self.ui.txtIntegrationY.setText(str(value))

    def updateLineEditZ(self, value):
        self.ui.txtIntegrationZ.setText(str(value))

    def updateLineEditK(self, value):
        self.ui.txtIntegrationK.setText(str(value))

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
        b = self.ui.txtIntegrationX.text()
        c = self.ui.txtIntegrationY.text()
        d = self.ui.txtIntegrationZ.text()
        e = self.ui.txtIntegrationK.text()
        f = self.ui.txtIntegrationDelay.text()
        g = self.ui.txtIntegrationKet.text()
        h = self.ui.lbIntegrationGetCode.text()

        if not b or not c or not d or not e or not g:
            QMessageBox.warning(self, "Warning", "Please input your data.")
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
            getDesc = self.integrationModel.index(getRow, 5).data()
            getDelay = self.integrationModel.index(getRow, 6).data()

            self.ui.txtIntegrationCode.setText(getCode)
            self.ui.txtIntegrationKet.setText(getDelay)
            self.ui.txtIntegrationDelay.setText(getDesc)

            self.ui.txtIntegrationX.setText(str(getX))
            self.ui.txtIntegrationY.setText(str(getY))
            self.ui.txtIntegrationZ.setText(str(getZ))
            self.ui.txtIntegrationK.setText(str(getK))

            self.ui.slIntegrationX.setValue(getX)
            self.ui.slIntegrationY.setValue(getY)
            self.ui.slIntegrationZ.setValue(getZ)
            self.ui.slIntegrationK.setValue(getK)

            sendSerial(getX, getY, getK, getZ)

    def keyReleaseEventIntegration(self, event):
        if event.key() == QtCore.Qt.Key_Enter or event.key() == QtCore.Qt.Key_Return:
            self.searchDataIntegration(event)
        else:
            super().keyReleaseEvent(event)

    def setIntegrationEmptyColumn(self):
        self.ui.txtIntegrationKet.setText("")
        self.ui.txtIntegrationDelay.setText("0")
        self.ui.txtIntegrationX.setText("0")
        self.ui.txtIntegrationY.setText("0")
        self.ui.txtIntegrationZ.setText("0")
        self.ui.txtIntegrationK.setText("0")

        x = self.ui.slIntegrationX.setValue(0)
        y = self.ui.slIntegrationY.setValue(0)
        z = self.ui.slIntegrationZ.setValue(0)
        k = self.ui.slIntegrationK.setValue(0)

        sendSerial(0, 0, 0, 0)

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
                        ["C-Code", "X", "Y", "Z", "K", "Delay", "Description", ""]
                    )
                    font = QtGui.QFont()
                    font.setPointSize(14)

                    header_view = self.ui.tbRunningProjectTabel_1.horizontalHeader()
                    header_view.setStyleSheet("background-color: rgb(114, 159, 207);")
                    font = header_view.font()
                    font.setPointSize(18)
                    header_view.setFont(font)

                    self.ui.tbRunningProjectTabel_1.setColumnWidth(0, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(1, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(2, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(3, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(4, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(5, 100)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(6, 400)
                    self.ui.tbRunningProjectTabel_1.setColumnWidth(7, 100)

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

                        self.runningModel_1.appendRow(items)

                    self.ui.tbRunningProjectTabel_1.resizeColumnsToContents()
                    self.ui.tbRunningProjectTabel_1.verticalHeader().setDefaultSectionSize(
                        50
                    )

        finally:
            if connection:
                connection.close()

    def rowClicked(self, index):
        selected_row_index = index.row()
        column_count = self.runningModel_1.columnCount()
        selected_row_data = []
        for column_index in range(column_count):
            item = self.runningModel_1.item(selected_row_index, column_index)
            selected_row_data.append(item.text())

        self.stacked_data.append(selected_row_data)


    def moveToTable2(self):
        if self.stacked_data:
            for data in self.stacked_data:
                items = [QStandardItem(str(item)) for item in data]
                self.runningModel_2.appendRow(items)

            self.clicked_times = 0
            self.stacked_data.clear()


    def removeFromTable2(self):
        selected_index = self.ui.tbRunningProjectTabel_2.currentIndex()
        if selected_index.isValid():
            selected_row = selected_index.row()
            self.runningModel_2.removeRow(selected_row)
    
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
                        self.ui.tbChooseProjectTabel.resizeColumnsToContents()
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
                        self.ui.tbChooseProjectTabel.resizeColumnsToContents()
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
                        self.ui.tbChooseActivityTabel.setColumnWidth(2, 500)
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
        global currentX
        global currentY
        global currentZ
        global currentD
        global currentK
        valueX = self.ui.slIntegrationX.value()
        valueY = self.ui.slIntegrationY.value()
        valueZ = self.ui.slIntegrationZ.value()
        valueK = self.ui.slIntegrationK.value()
        sendSerial(valueX, valueY, valueK, valueZ)

def writeSerial(data):
    ser.write(bytes(data, 'utf-8'))
    time.sleep(0.05)

def sendSerial(x, y, z, k):
    data = "config:%d,%d,%d,%d;" % (x, y, z, k)
    writeSerial(data)
    print(data)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.showMaximized()
    sys.exit(app.exec_())

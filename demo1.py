import time
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QWidget, QGroupBox, QMenu, QAction,
                             QPushButton, QCheckBox, QRadioButton,
                             QVBoxLayout, QGridLayout)
import mid
#import check
from PyQt5.QtWidgets import QMessageBox

"""class EmitStr(QObject):
    '''
    定义一个信号类，
    sys.stdout有个write方法，通过重定向，
    每当有新字符串输出时就会触发下面定义的write函数，
    进而发出信号
    text：新字符串，会通过信号传递出去
    '''
    textWrit  = pyqtSignal(str)
    def write(self, text):
        self.textWrit.emit(str(text))"""

class Ui_MainWindow(object):
    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.file_path = None
        self.rexcel =None
        self.money =None
        self.str_in =None
        # self.setupUi(self,MainWindow)
        # sys.stdout = EmitStr(textWrit=self.outputWrite)  # 输出结果重定向
        # sys.stderr = EmitStr(textWrit=self.outputWrite)  # 错误输出重定向

        # 实时显示输出, 将控制台的输出重定向到界面中


    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(800, 600)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.verticalLayout_2.addWidget(self.label)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.comboBox = QtWidgets.QPushButton(self.centralwidget)
        self.comboBox.setObjectName("comboBox")



        self.horizontalLayout.addWidget(self.comboBox)
        self.toolButton = QtWidgets.QToolButton(self.centralwidget)
        self.toolButton.setObjectName("toolButton")
        self.horizontalLayout.addWidget(self.toolButton)
        self.verticalLayout_2.addLayout(self.horizontalLayout)
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.verticalLayout_2.addWidget(self.line)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.radioButton = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton.setChecked(True)
        self.radioButton.setObjectName("radioButton")
        self.horizontalLayout_2.addWidget(self.radioButton)
        self.radioButton_2 = QtWidgets.QRadioButton(self.centralwidget)
        self.radioButton_2.setObjectName("radioButton_2")
        self.horizontalLayout_2.addWidget(self.radioButton_2)
        self.radioButton_3 = QtWidgets.QRadioButton(self.centralwidget)

        self.radioButton_3.setObjectName("radioButton_3")
        self.horizontalLayout_2.addWidget(self.radioButton_3)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.verticalLayout_2.addWidget(self.line_2)

        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_5.addWidget(self.label_7)
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setObjectName("checkBox")
        self.horizontalLayout_5.addWidget(self.checkBox)
        self.checkBox_2 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_2.setObjectName("checkBox_2")
        self.horizontalLayout_5.addWidget(self.checkBox_2)
        self.checkBox_3 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_3.setObjectName("checkBox_3")
        self.horizontalLayout_5.addWidget(self.checkBox_3)
        self.checkBox_4 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_4.setObjectName("checkBox_4")
        self.horizontalLayout_5.addWidget(self.checkBox_4)
        self.checkBox_5 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_5.setObjectName("checkBox_5")
        self.horizontalLayout_5.addWidget(self.checkBox_5)
        self.checkBox_6 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_6.setObjectName("checkBox_6")
        self.horizontalLayout_5.addWidget(self.checkBox_6)
        self.button2 = QtWidgets.QPushButton(self.centralwidget)
        self.button2.setObjectName("button2")
        self.horizontalLayout_5.addWidget(self.button2)
        self.verticalLayout_2.addLayout(self.horizontalLayout_5)

        self.line_3 = QtWidgets.QFrame(self.centralwidget)
        self.line_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.verticalLayout_2.addWidget(self.line_3)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_4.addWidget(self.label_5)

        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setObjectName("lineEdit")
        self.button =QtWidgets.QPushButton(self.centralwidget)
        self.button.setObjectName("button")

        self.horizontalLayout_4.addWidget(self.lineEdit)
        self.horizontalLayout_4.addWidget(self.button)

        self.verticalLayout_2.addLayout(self.horizontalLayout_4)
        self.line_4 = QtWidgets.QFrame(self.centralwidget)
        self.line_4.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_4.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_4.setObjectName("line_4")
        self.verticalLayout_2.addWidget(self.line_4)
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setObjectName("label_3")
        self.verticalLayout.addWidget(self.label_3)
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setEnabled(True)
        self.textEdit.setObjectName("textEdit")
        self.verticalLayout.addWidget(self.textEdit)
        self.button3 = QtWidgets.QPushButton(self.centralwidget)
        self.button3.setObjectName("button3")
        self.verticalLayout.addWidget(self.button3)
        self.verticalLayout_2.addLayout(self.verticalLayout)

        self.line_6 = QtWidgets.QFrame(self.centralwidget)
        self.line_6.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_6.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_6.setObjectName("line_6")
        self.verticalLayout_2.addWidget(self.line_6)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setObjectName("label_8")

        self.verticalLayout_3.addWidget(self.label_8)

        self.gridLayout = QtWidgets.QGridLayout()
        self.gridLayout.setObjectName("gridLayout")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.gridLayout.addWidget(self.lineEdit_2, 0, 0, 1, 1)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.gridLayout.addWidget(self.lineEdit_3, 0, 1, 1, 1)
        self.lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.gridLayout.addWidget(self.lineEdit_4, 1, 0, 1, 1)
        self.lineEdit_5 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.gridLayout.addWidget(self.lineEdit_5, 1, 1, 1, 1)
        self.lineEdit_6 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.gridLayout.addWidget(self.lineEdit_6, 2, 0, 1, 2)
        self.verticalLayout_3.addLayout(self.gridLayout)

        self.button4 = QtWidgets.QPushButton(self.centralwidget)
        self.button4.setObjectName("button4")
        self.verticalLayout_3.addWidget(self.button4)
        self.verticalLayout_2.addLayout(self.verticalLayout_3)

        self.line_5 = QtWidgets.QFrame(self.centralwidget)
        self.line_5.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_5.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_5.setObjectName("line_5")
        self.verticalLayout_2.addWidget(self.line_5)
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setObjectName("label_6")
        self.verticalLayout_2.addWidget(self.label_6)




        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setObjectName("textBrowser")
        self.verticalLayout_2.addWidget(self.textBrowser)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.font2 = QtGui.QFont()
        self.font2.setFamily('微软雅黑')
        self.font2.setBold(True)
        self.font2.setPointSize(9)

        self.font1 = QtGui.QFont()
        self.font1.setFamily('微软雅黑')
        self.font1.setBold(True)
        self.font1.setPointSize(12)
        # self.font1.setWeight(20)
        self.label.setFont(self.font2)
        self.label_2.setFont(self.font2)
        self.label_3.setFont(self.font2)
        self.label_5.setFont(self.font2)
        self.label_6.setFont(self.font2)
        self.label_7.setFont(self.font2)
        self.label_8.setFont(self.font2)
        self.button2.setStyleSheet(

            "QPushButton{color: white}"
            "QPushButton{background:#2e486c}"
            "QPushButton:hover{background:#8FAADC;}"
            "QPushButton{border:black;border-width=2;border-style:groove;border-radius: 15px;}"


        )
        self.comboBox.setStyleSheet(
            "QPushButton{background:white}"
        )
        self.button.setStyleSheet(


            "QPushButton:hover{background:#8FAADC;}"
            "QPushButton{color: white;background:#2e486c;border-radius: 15px;border:black;border-width=2;border-style:groove;}"

        )
        self.lineEdit.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=10, xOffset=3, yOffset=2))
        self.textEdit.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=10, xOffset=3, yOffset=2))
        self.lineEdit_2.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=10, xOffset=3, yOffset=2))
        self.lineEdit_3.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=10, xOffset=3, yOffset=2))
        self.lineEdit_6.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=10, xOffset=3, yOffset=2))
        self.lineEdit_5.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=10, xOffset=3, yOffset=2))
        self.lineEdit_4.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=10, xOffset=3, yOffset=2))
        self.textBrowser.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=10, xOffset=3, yOffset=2))
        self.button2.setFont(self.font1)
        self.button.setFont(self.font1)
        self.toolButton.setFixedSize(80 ,60)
        self.toolButton.setToolTip("选择文件夹")
        # self.toolButton.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.button.setFixedSize(200 ,60)
        self.button2.setFixedSize(200, 60)
        self.button3.setFont(self.font1)
        self.button4.setFont(self.font1)
        self.button3.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=30, xOffset=5, yOffset=3))
        self.button4.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=30, xOffset=5, yOffset=3))
        # self.label.setGraphicsEffect(QtWidgets.QGraphicsDropShadowEffect(blurRadius=25, xOffset=0, yOffset=0))
        """self.button2.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=30, xOffset=5, yOffset=3))
        self.button2.setFont(self.font1)
        self.button.setFont(self.font1)
        self.button.setGraphicsEffect(
            QtWidgets.QGraphicsDropShadowEffect(blurRadius=30, xOffset=5, yOffset=3))"""


        op = QtWidgets.QGraphicsOpacityEffect()
        # 设置透明度的值，0.0到1.0，最小值0是透明，1是不透明
        op.setOpacity(0.8)
        self.button2.setGraphicsEffect(op)
        self.button.setGraphicsEffect(op)

        self.button3.setEnabled(False)
        self.textEdit.setEnabled(False)
        self.radioButton.toggled.connect(self.close1)
        self.radioButton_2.toggled.connect(self.close1)
        self.radioButton_3.toggled.connect(self.close1)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    """def outputWrite(self, text):
        self.textBrowser.append(text)       # 输出的字符追加到 QTextEdit 中"""
    def close1(self):
        if self.radioButton.isChecked( )==True:
            self.button3.setEnabled(False)
            self.button.setEnabled(True)
            self.button2.setEnabled(True)
            self.checkBox.setEnabled(True)
            self.checkBox_2.setEnabled(True)
            self.checkBox_3.setEnabled(True)
            self.checkBox_4.setEnabled(True)
            self.checkBox_5.setEnabled(True)
            self.checkBox_6.setEnabled(True)
            self.textEdit.setEnabled(False)
            self.lineEdit.setEnabled(True)

            self.lineEdit_2.setEnabled(False)
            self.lineEdit_3.setEnabled(False)
            self.lineEdit_4.setEnabled(False)
            self.lineEdit_5.setEnabled(False)
            self.lineEdit_6.setEnabled(False)
            self.button4.setEnabled(False)
        elif self.radioButton_2.isChecked()==True :
            self.button3.setEnabled(True)
            self.button.setEnabled(False)
            self.button2.setEnabled(False)
            self.checkBox.setEnabled(False)
            self.checkBox_2.setEnabled(False)
            self.checkBox_3.setEnabled(False)
            self.checkBox_4.setEnabled(False)
            self.checkBox_5.setEnabled(False)
            self.checkBox_6.setEnabled(False)
            self.textEdit.setEnabled(True)
            self.lineEdit.setEnabled(False)

            self.lineEdit_2.setEnabled(False)
            self.lineEdit_3.setEnabled(False)
            self.lineEdit_4.setEnabled(False)
            self.lineEdit_5.setEnabled(False)
            self.lineEdit_6.setEnabled(False)
            self.button4.setEnabled(False)
        elif self.radioButton_3.isChecked()==True:
            self.lineEdit_2.setEnabled(True)
            self.lineEdit_3.setEnabled(True)
            self.lineEdit_4.setEnabled(True)
            self.lineEdit_5.setEnabled(True)
            self.lineEdit_6.setEnabled(True)
            self.button4.setEnabled(True)

            self.button3.setEnabled(False)
            self.button.setEnabled(False)
            self.button2.setEnabled(False)
            self.checkBox.setEnabled(False)
            self.checkBox_2.setEnabled(False)
            self.checkBox_3.setEnabled(False)
            self.checkBox_4.setEnabled(False)
            self.checkBox_5.setEnabled(False)
            self.checkBox_6.setEnabled(False)
            self.textEdit.setEnabled(False)
            self.lineEdit.setEnabled(False)

    def msg(self ,Filepath):
        m = QtWidgets.QFileDialog.getExistingDirectory(None, "选取文件夹", "../")  # 起始路径
        self.comboBox.setText(m)
        self.rexcel =mid.read(m)
        self.file_path =m




    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "选择发票存储文件夹"))
        self.toolButton.setText(_translate("MainWindow", "..."))
        self.toolButton.setIcon(QIcon("file.png"))

        self.label_2.setText(_translate("MainWindow", "选择你需要的工作模式"))
        self.radioButton.setText(_translate("MainWindow", "正常报销"))
        self.radioButton_2.setText(_translate("MainWindow", "回滚模式"))


        self.label_5.setText(_translate("MainWindow", "输入需要报销的价格"))
        self.button.setText(_translate("MainWindow" ,"确认"))
        self.button2.setText(_translate("MainWindow", "确认"))
        self.button3.setText(_translate("MainWindow", "确认"))
        self.button4.setText(_translate("MainWindow", "确认"))
        self.label_3.setText(_translate("MainWindow", "输入需要回滚的发票代码（注意不是发票号码）,如需回滚多张请用英文逗号隔开.例如：1234566,1234567"))
        self.label_6.setText(_translate("MainWindow", "输出框"))
        self.label_7.setText(_translate("MainWindow", "选择需要报销的类别"))
        self.label_8.setText(_translate("MainWindow" ,"输入需要真伪核验的发票"))
        self.checkBox.setText(_translate("MainWindow", "交通费"))
        self.checkBox_2.setText(_translate("MainWindow", "差旅费"))
        self.checkBox_3.setText(_translate("MainWindow", "耗材费"))
        self.checkBox_4.setText(_translate("MainWindow", "印刷费"))
        self.checkBox_5.setText(_translate("MainWindow", "培训费"))
        self.checkBox_6.setText(_translate("MainWindow", "其他"))
        self.radioButton_3.setText(_translate("MainWindow", "真伪核验"))
        self.toolButton.clicked.connect(self.msg)
        self.button.clicked.connect(self.sure)
        self.button2.clicked.connect(self.getvalue1)
        self.button3.clicked.connect(self.getvalue2)
        self.lineEdit_2.setPlaceholderText("请输入发票代码：")
        self.lineEdit_3.setPlaceholderText("请输入发票号码：")
        self.lineEdit_4.setPlaceholderText("请输入校验码后6位：")
        self.lineEdit_5.setPlaceholderText("请输入开票日期,格式为20221107：")
        self.lineEdit_6.setPlaceholderText("请输入金额：")

        self.button4.clicked.connect(self.getvalue3)

        # sys.stdout = Signal()
        # sys.stdout.text_update.connect(self.updatetext)


    def getvalue3(self):
        fpdm = self.lineEdit_2.toPlainText()
        fphm = self.lineEdit_2.toPlainText()
        checkCode = self.lineEdit_4.toPlainText()
        date = self.lineEdit_5.toPlainText()
        noTaxAmount = self.lineEdit_6.toPlainText()
        flag = check.check_invoice(checkCode,fpdm,fphm,date,noTaxAmount)
        if flag == 1:
            QtWidgets.QMessageBox.information(self.button3, "真伪核验", "发票鉴定为真！", QMessageBox.Yes)
            self.textBrowser.append("发票鉴定为真！")
            self.textBrowser.moveCursor(self.textBrowser.textCursor().End)  # 文本框显示到底部
            time.sleep(0.1)
        else:
            QtWidgets.QMessageBox.warning(self.button3, "真伪核验", '发票信息错误或发票为假！', QMessageBox.Yes)
            self.textBrowser.append("发票信息错误或发票为假！")
            self.textBrowser.moveCursor(self.textBrowser.textCursor().End)  # 文本框显示到底部
            time.sleep(0.1)

    def sure(self ,MainWindow):
        self.money =self.lineEdit.text()
        list ,num =mid.select(self.str_in ,self.file_path ,self.rexcel ,self.money)
        if num==0:
            QtWidgets.QMessageBox.warning(self.button3, "报销的发票", list + '\n', QMessageBox.Yes )
        else:
            QtWidgets.QMessageBox.information(self.button3, "报销的发票", list + '\n', QMessageBox.Yes | QMessageBox.No,
                                              QMessageBox.Yes)
        # todo:
        self.textBrowser.append(list)
        self.textBrowser.moveCursor(self.textBrowser.textCursor().End)  # 文本框显示到底部
        time.sleep(0.1)

    def getvalue2(self):
        Inovice_ID = self.textEdit.toPlainText().split(',')
        self.textBrowser.append(mid.RollBack(Inovice_ID, self.file_path))
        self.textBrowser.moveCursor(self.textBrowser.textCursor().End)  # 文本框显示到底部

    def getvalue1(self):
        # from PyQt5.QtWidgets import QMessageBox
        list = ""
        m=''

        if self.checkBox.isChecked():  # 判断复选框是否被选中
            list += "\n" + self.checkBox.text()  # 记录选中的权限
            m='1 '
        if self.checkBox_2.isChecked():
            list += "\n" + self.checkBox_2.text()
            m=m+'2 '
        if self.checkBox_3.isChecked():
            list += "\n" + self.checkBox_3.text()
            m=m+'3 '
        if self.checkBox_4.isChecked():
            list += "\n" + self.checkBox_4.text()
            m=m+'4 '
        if self.checkBox_5.isChecked():
            list += "\n" + self.checkBox_5.text()
            m=m+'5 '
        if self.checkBox_6.isChecked():
            list += "\n" + self.checkBox_6.text()
            m = m + '6 '
        self.str_in=m
        QtWidgets.QMessageBox.information(self.button3, "选择的发票种类", list + '\n',QMessageBox.Yes|QMessageBox.No,QMessageBox.Yes)
        # QMessageBox.information(self,'选择的发票种类',list,QMessageBox.Yes|QMessageBox.No,QMessageBox.Yes)
        self.list=list

    """def updatetext(self, text):
        cursor = self.textBrowser.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.textBrowser.append(text)
        self.textBrowser.setTextCursor(cursor)
        self.textBrowser.ensureCursorVisible()"""

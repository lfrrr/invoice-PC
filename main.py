# 开发第一个基于PyQt5的桌面应用

import sys

import demo1
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *

from PyQt5.QtWidgets import QApplication, QMainWindow





if __name__ == '__main__':
    # 创建QApplication类的实例

    app = QApplication([])
    # 创建一个窗口
    w = QMainWindow()
    w.setWindowIcon(QIcon('pic1.png'))
    #w.setStyleSheet("#MainWindow{border-image:url(唐峰素材 (14).jpg)}")  # 这里使用相对路径，也可以使用绝对路径

    #w.setWindowOpacity(0.85)  # 设置窗口透明度
    # 设置窗口尺寸   宽度300，高度150
    ui = demo1.Ui_MainWindow()
    ui.setupUi(w)
    # 设置窗口的标题
    w.setWindowTitle('闪易发票')

    # 显示窗口
    w.show()

    # 进入程序的主循环，并通过exit函数确保主循环安全结束(该释放资源的一定要释放)
    app.exec_()

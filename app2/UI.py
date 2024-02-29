from PyQt5.QtWidgets import QApplication, QPushButton, QComboBox, QLabel, QMessageBox, QDesktopWidget, \
    QWidget
from PyQt5.QtCore import Qt, pyqtSlot
from PyQt5.QtGui import QIcon, QPalette, QColor, QFont
from function import *
import sys
import pandas as pd
import warnings
warnings.filterwarnings("ignore")
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        # 设置窗口标题和大小
        self.setWindowTitle('CAP程序')
        self.resize(400, 400)
        # 设置窗口背景颜色
        self.setAutoFillBackground(True)
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(230, 230, 250))
        self.setPalette(palette)

        self.button_WX_K = QPushButton('WX_K', self)
        self.button_WX_K.move(150, 50)
        self.button_WX_K.resize(100, 50)
        self.button_WX_K.clicked.connect(self.open_WX_K)
        self.button_WX_K.setFont(QFont('Arial', 12))
        self.button_WX_K.setStyleSheet('QPushButton {background-color: lightblue; }')

        self.button_MS = QPushButton('MS', self)
        self.button_MS.move(150, 150)
        self.button_MS.resize(100, 50)
        self.button_MS.clicked.connect(self.open_MS)
        self.button_MS.setFont(QFont('Arial', 12))
        self.button_MS.setStyleSheet('QPushButton {background-color: lightblue; }')

        self.button_MS_WK = QPushButton('MS_WK', self)
        self.button_MS_WK.move(150, 250)
        self.button_MS_WK.resize(100, 50)
        self.button_MS_WK.clicked.connect(self.open_MS_WK)
        self.button_MS_WK.setFont(QFont('Arial', 12))
        self.button_MS_WK.setStyleSheet('QPushButton {background-color: lightblue; }')

    @pyqtSlot()
    def open_WX_K(self):
        self.WX_K_window = WX_K_Window()
        self.WX_K_window.show()
        self.hide()  # 隐藏主窗口

    @pyqtSlot()
    def open_MS(self):
        self.MS_window = MS_Window()
        self.MS_window.show()
        self.hide()  # 隐藏主窗口

    @pyqtSlot()
    def open_MS_WK(self):
        self.MS_WK_window = MS_WK_Window()
        self.MS_WK_window.show()
        self.hide()  # 隐藏主窗口
class WX_K_Window(QWidget):
    # 初始化方法
    def __init__(self):
        # 调用父类的初始化方法
        super().__init__()
        # 设置窗口标题和大小
        self.setWindowTitle("WX&K")
        self.resize(600, 400)
        self.setWindowIcon(QIcon('copilot.png'))
        self.setAutoFillBackground(True)
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(255, 240, 245))
        self.setPalette(palette)
        self.center()
        self.create_widgets()
        # 创建一个返回按钮
        self.button_back = QPushButton('返回', self)
        self.button_back.move(250, 150)
        self.button_back.resize(100, 40)
        self.button_back.clicked.connect(self.go_back)
        self.button_back.setFont(QFont('Arial', 12))
        self.button_back.setStyleSheet('QPushButton {background-color: pink; }')
    # 创建窗口控件的方法
    def create_widgets(self):
        self.label1 = QLabel("输入：", self)
        self.label1.setGeometry(50, 50, 150, 30)
        self.label1.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.label1.setFont(QFont('Arial', 12))
        self.combo1 = QComboBox(self)
        self.combo1.setGeometry(200, 50, 250, 30)
        self.combo1.addItems(["WX_pre_get_name", "WX_get_name", "K_pre_get_name", "K_get_name"])
        self.combo1.setFont(QFont('Arial', 12))
        self.label2 = QLabel("输出：", self)
        self.label2.setGeometry(50, 100, 150, 30)
        self.label2.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.label2.setFont(QFont('Arial', 12))
        self.combo2 = QComboBox(self)
        self.combo2.setGeometry(200, 100, 250, 30)
        self.combo2.addItems(["wx_pre_output", "wx_output", "k_pre_output", "k_output"])
        self.combo2.setFont(QFont('Arial', 12))
        self.button = QPushButton("运行", self)
        self.button.setGeometry(150, 150, 100, 40)
        self.button.clicked.connect(self.run)
        self.button.setFont(QFont('Arial', 12))
        self.button.setStyleSheet('QPushButton {background-color: pink; }')
    def run(self):
        sheet_name = self.combo1.currentText()
        output_name = self.combo2.currentText()
        filenames = pd.read_excel('./middle_file/map.xlsx', sheet_name=sheet_name)
        filenames = filenames.values.tolist()
        all_output = []
        for i in range(0, 6, 3):
            filenames[i][2] = process_data(filenames[i][0], filenames[i][1])
            filenames[i + 1][2] = process_data(filenames[i + 1][0], filenames[i + 1][1])
            filenames[i + 2][2] = process_data(filenames[i + 2][0], filenames[i + 2][1])
            filenames[i][3] = mergeadd(filenames[i][2], filenames[i + 1][2], filenames[i + 2][2])
            all_output.append(filenames[i][3])
        all_output = pd.concat(all_output, axis=0, ignore_index=True)
        all_output = all_output.sort_values(by=['DPN', 'Capacity', 'Heads', 'Discs'])
        all_output = all_output.rename(columns={'DPN': 'DPN', 'Capacity': 'CAP', 'Heads': 'HEADS', 'Discs': 'DISCS'})
        with pd.ExcelWriter(r'./middle_file/result.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            all_output.to_excel(writer, sheet_name=output_name, index=False)
        msg_box = QMessageBox(QMessageBox.Information, "提示", "运行成功，请退出")
        msg_box.exec_()
    # 定义窗口居中显示的方法
    def center(self):
        # 获取窗口的几何信息
        qr = self.frameGeometry()
        # 获取屏幕的中心点
        cp = QDesktopWidget().availableGeometry().center()
        # 将窗口的中心点移动到屏幕的中心点
        qr.moveCenter(cp)
        # 将窗口的左上角移动到qr的左上角
        self.move(qr.topLeft())
    @pyqtSlot()
    def go_back(self):
        self.MainWindow = MainWindow()
        self.MainWindow.show()
        self.hide()  # 隐藏当前窗口
class MS_Window(QWidget):
    # 初始化方法
    def __init__(self):
        super().__init__()
        self.setWindowTitle('MS')
        self.resize(600, 400)
        self.setWindowIcon(QIcon('copilot.png'))
        self.setAutoFillBackground(True)
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(255, 240, 245))
        self.setPalette(palette)
        self.center()
        self.create_widgets()
        # 创建一个返回按钮
        self.button_back = QPushButton('返回', self)
        self.button_back.move(250, 200)
        self.button_back.resize(100, 40)
        self.button_back.clicked.connect(self.go_back)
        self.button_back.setFont(QFont('Arial', 12))
        self.button_back.setStyleSheet('QPushButton {background-color: pink; }')

    # 创建窗口控件的方法
    def create_widgets(self):
        self.label = QLabel("LOC：", self)
        self.label.setGeometry(50, 50, 150, 30)
        self.label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.label.setFont(QFont('Arial', 12))
        self.combo = QComboBox(self)
        self.combo.setGeometry(200, 50, 150, 30)
        self.combo.addItems(["KOR", "WUX"])
        self.combo.setFont(QFont('Arial', 12))

        self.label1 = QLabel("name1：", self)
        self.label1.setGeometry(50, 100, 150, 30)
        self.label1.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.label1.setFont(QFont('Arial', 12))
        self.combo1 = QComboBox(self)
        self.combo1.setGeometry(200, 100, 150, 30)
        self.combo1.addItems(["k_jul", "k_aug", "w_jul", "w_aug"])
        self.combo1.setFont(QFont('Arial', 12))


        self.label2 = QLabel("name2：", self)
        self.label2.setGeometry(50, 150, 150, 30)
        self.label2.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.label2.setFont(QFont('Arial', 12))
        self.combo2 = QComboBox(self)
        self.combo2.setGeometry(200, 150, 150, 30)
        self.combo2.addItems(["k_jul", "k_aug", "w_jul", "w_aug"])
        self.combo2.setFont(QFont('Arial', 12))

        self.button = QPushButton("运行", self)
        self.button.setGeometry(150, 200, 100, 40)
        self.button.clicked.connect(self.run)
        self.button.setFont(QFont('Arial', 12))
        self.button.setStyleSheet('QPushButton {background-color: pink; }')

    def run(self):
        sheet_name=self.combo.currentText()
        sheet_name1 = self.combo1.currentText()
        sheet_name2 = self.combo2.currentText()
        output_jul, output_aug = filter_and_process(sheet_name)
        with pd.ExcelWriter('./middle_file/add_result.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            output_jul.to_excel(writer, sheet_name=sheet_name1, index=False)
            output_aug.to_excel(writer, sheet_name=sheet_name2, index=False)
        msg_box = QMessageBox(QMessageBox.Information, "提示", "运行成功，请退出")
        msg_box.exec_()
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    @pyqtSlot()
    def go_back(self):
        self.MainWindow = MainWindow()
        self.MainWindow.show()
        self.hide()  # 隐藏当前窗口
class MS_WK_Window(QWidget):
    # 初始化方法
    def __init__(self):
        # 调用父类的初始化方法
        super().__init__()
        self.setWindowTitle('MS_WK')
        self.resize(600, 400)
        self.setAutoFillBackground(True)
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(255, 240, 245))
        self.setPalette(palette)
        self.center()
        self.create_widgets()
        # 创建一个返回按钮
        self.button_back = QPushButton('返回', self)
        self.button_back.move(250, 150)
        self.button_back.resize(100, 40)
        self.button_back.clicked.connect(self.go_back)
        self.button_back.setFont(QFont('Arial', 12))
        self.button_back.setStyleSheet('QPushButton {background-color: pink; }')
    def create_widgets(self):
        self.label1 = QLabel("name1：", self)
        self.label1.setGeometry(50, 50, 150, 30)
        self.label1.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.label1.setFont(QFont('Arial', 12))
        self.combo1 = QComboBox(self)
        self.combo1.setGeometry(200, 50, 200, 30)
        self.combo1.addItems(["wx_pre_output", "wx_output", "k_pre_output", "k_output"])
        self.combo1.setFont(QFont('Arial', 12))
        self.label2 = QLabel("name2：", self)
        self.label2.setGeometry(50, 100, 150, 30)
        self.label2.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.label2.setFont(QFont('Arial', 12))
        self.combo2 = QComboBox(self)
        self.combo2.setGeometry(200, 100, 200, 30)
        self.combo2.addItems(["wx_pre_output", "wx_output", "k_pre_output", "k_output"])
        self.combo2.setFont(QFont('Arial', 12))
        self.button = QPushButton("运行", self)
        self.button.setGeometry(150, 150, 100, 40)
        self.button.clicked.connect(self.run)
        self.button.setFont(QFont('Arial', 12))
        self.button.setStyleSheet('QPushButton {background-color: pink; }')
    def run(self):
        sheet_name1 = self.combo1.currentText()
        sheet_name2 = self.combo2.currentText()
        result = MS_WK(sheet_name1, sheet_name2)
        with pd.ExcelWriter('./middle_file/add_result.xlsx', engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            result.to_excel(writer, sheet_name='for_model', index=False)
        msg_box = QMessageBox(QMessageBox.Information, "提示", "运行成功，请退出")
        msg_box.exec_()
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    @pyqtSlot()
    def go_back(self):
        self.MainWindow = MainWindow()
        self.MainWindow.show()
        self.hide()  # 隐藏当前窗口
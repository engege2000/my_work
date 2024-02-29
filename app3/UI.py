from process_function import *
from match_function import *
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QLineEdit, QFileDialog
from PyQt5.QtCore import pyqtSlot, Qt
from PyQt5.QtGui import QIcon, QFont, QPalette, QColor
# 定义主窗口类
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        # 设置窗口标题和大小
        self.setWindowTitle('HDA_MEAS程序')
        self.resize(400, 250)
        # 设置窗口图标
        self.setWindowIcon(QIcon('./need_file/welcome.jpeg'))
        # 设置窗口背景颜色
        self.setAutoFillBackground(True)
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(230, 230, 250))
        self.setPalette(palette)
        # 创建两个功能窗口的按钮
        self.button_YD_TT = QPushButton('YD_TT', self)
        self.button_YD_TT.move(150, 50)
        self.button_YD_TT.resize(100, 50)
        self.button_YD_TT.clicked.connect(self.open_YD_TT)
        self.button_YD_TT.setFont(QFont('Arial', 12))
        self.button_YD_TT.setStyleSheet('QPushButton {background-color: lightblue; }')
        self.button_merge = QPushButton('Merge', self)
        self.button_merge.move(150, 150)
        self.button_merge.resize(100, 50)
        self.button_merge.clicked.connect(self.open_merge)
        self.button_merge.setFont(QFont('Arial', 12))
        self.button_merge.setStyleSheet('QPushButton {background-color: lightblue; }')
    # 定义打开YD_TT功能窗口的槽函数
    @pyqtSlot()
    def open_YD_TT(self):
        self.YD_TT_window = YD_TT_Window()
        self.YD_TT_window.show()
        self.hide() # 隐藏主窗口
    # 定义打开合并功能窗口的槽函数
    @pyqtSlot()
    def open_merge(self):
        self.merge_window = Merge_Window()
        self.merge_window.show()
        self.hide() # 隐藏主窗口
# 定义YD_TT功能窗口类
class YD_TT_Window(QWidget):
    def __init__(self):
        super().__init__()
        # 设置窗口标题和大小
        self.setWindowTitle('YD_TT功能')
        self.resize(650, 400)
        # 设置窗口图标
        self.setWindowIcon(QIcon('copilot.png'))
        # 设置窗口背景颜色
        self.setAutoFillBackground(True)
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(240, 255, 240))
        self.setPalette(palette)
        # 创建上传文件按钮
        self.button_upload = QPushButton('上传文件', self)
        self.button_upload.move(50, 50)
        self.button_upload.resize(100, 50)
        self.button_upload.clicked.connect(self.upload_file)
        self.button_upload.setFont(QFont('Arial', 12))
        self.button_upload.setStyleSheet('QPushButton {background-color: lightblue; }')
        # 创建自动填充按钮
        self.button_fill = QPushButton('自动填充', self)
        self.button_fill.move(200, 50)
        self.button_fill.resize(100, 50)
        self.button_fill.clicked.connect(self.auto_fill)
        self.button_fill.setFont(QFont('Arial', 12))
        self.button_fill.setStyleSheet('QPushButton {background-color: lightblue; }')
        # 创建运行按钮
        self.button_run = QPushButton('运行', self)
        self.button_run.move(350, 50)
        self.button_run.resize(100, 50)
        self.button_run.clicked.connect(self.run_code)
        self.button_run.setFont(QFont('Arial', 12))
        self.button_run.setStyleSheet('QPushButton {background-color: pink; }')
        # 创建一个返回按钮
        self.button_back = QPushButton('返回', self)
        self.button_back.move(500, 50)
        self.button_back.resize(100, 50)
        self.button_back.clicked.connect(self.go_back)
        self.button_back.setFont(QFont('Arial', 12))
        self.button_back.setStyleSheet('QPushButton {background-color: pink; }')
        # 创建文件路径显示标签
        self.label_path = QLabel('文件路径：', self)
        self.label_path.move(50, 150)
        self.label_path.setFont(QFont('Arial', 12))
        # 创建文件路径输入框
        self.line_path = QLineEdit(self)
        self.line_path.move(150, 150)
        self.line_path.resize(300, 30)
        self.line_path.setFont(QFont('Arial', 12))
        # 创建填充内容显示标签
        self.label_content = QLabel('填充内容：', self)
        self.label_content.move(50, 200)
        self.label_content.setFont(QFont('Arial', 12))
        # 创建填充内容输入框
        self.line_content = QLineEdit(self)
        self.line_content.move(150, 200)
        self.line_content.resize(300, 30)
        self.line_content.setFont(QFont('Arial', 12))

        # 创建填充内容显示标签
        self.label_state = QLabel('运行状态：', self)
        self.label_state.move(50, 250)
        self.label_state.setFont(QFont('Arial', 12))
        # 创建填充内容输入框
        self.line_state = QLineEdit(self)
        self.line_state.move(150, 250)
        self.line_state.resize(300, 30)
        self.line_state.setFont(QFont('Arial', 12))
    # 定义上传文件的槽函数
    @pyqtSlot()
    def upload_file(self):
        # 打开文件选择对话框
        file_name, _ = QFileDialog.getOpenFileName(self, '选择文件', './need_file', 'Excel files (*.xlsx)')
        # 如果选择了文件，将文件路径显示在输入框中
        if file_name:
            self.line_path.setText(file_name)
    # 定义自动填充的槽函数
    @pyqtSlot()
    def auto_fill(self):
        # 获取文件路径
        file_path = self.line_path.text()
        # 如果文件路径中含有Q424，填充为'Q424_PSE3.5'，否则填充为'Q324_PSE3.5'
        if 'Q324' in file_path:
            self.line_content.setText('Q324_PSE3.5')
        elif 'Q424' in file_path :
            self.line_content.setText('Q424_PSE3.5')
        else:
            self.line_content.setText('Q125_PSE3.5')
    # 定义运行代码的槽函数
    @pyqtSlot()
    def run_code(self):
        # 导入需要的模块和函数
        # 获取文件路径和填充内容
        file_path = self.line_path.text()
        fill_content = self.line_content.text()
        # 执行YD_TT功能的代码
        match_list=['match1','match2','match3','match4','match5']
        Q_list=['Q1','Q2','Q3','Q4','Q5']
        for match_list,Q_list in zip(match_list,Q_list):
            Q=concat_(file_path)
            Q1=match(Q,match_list)
            save(Q1,Q_list)
        Q1=pd.read_excel(r'./middle_file/result.xlsx', sheet_name='Q1')
        Q1.to_csv('Q1.csv',index=False)
        #Create an instance of QUpdater
        qu = QUpdater()
        # Loop through the sheets you want to merge
        sheets = ['Q2', 'Q3', 'Q4', 'Q5']
        matches = ['match2', 'match3', 'match4', 'match5']
        for i in range(len(sheets)):
            # Merge each sheet with Q1
            qu.merge_Q(sheets[i], matches[i])
            # Write the merged Q1 to the result.xlsx file in append mode
            qu.Q1.to_csv('Q1.csv')
        step_add_map(fill_content)
        # 显示运行结果
        self.line_state.setText('运行完毕，请退出')

    # 定义返回按钮的槽函数
    @pyqtSlot()
    def go_back(self):
        self.MainWindow = MainWindow()
        self.MainWindow.show()
        self.hide()  # 隐藏当前窗口
class Merge_Window(QWidget):
    def __init__(self):
        super().__init__()
        # 设置窗口标题和大小
        self.setWindowTitle('合并功能')
        self.resize(400, 300)
        # 设置窗口图标
        self.setWindowIcon(QIcon('copilot.png'))
        # 设置窗口背景颜色
        self.setAutoFillBackground(True)
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(255, 240, 245))
        self.setPalette(palette)
        # 创建运行按钮
        self.button_run = QPushButton('运行', self)
        self.button_run.move(100, 50)
        self.button_run.resize(100, 50)
        self.button_run.clicked.connect(self.run_code)
        self.button_run.setFont(QFont('Arial', 12))
        self.button_run.setStyleSheet('QPushButton {background-color: pink; }')
        # 创建一个返回按钮
        self.button_back = QPushButton('返回', self)
        self.button_back.move(250, 50)
        self.button_back.resize(100, 50)
        self.button_back.clicked.connect(self.go_back)
        self.button_back.setFont(QFont('Arial', 12))
        self.button_back.setStyleSheet('QPushButton {background-color: pink; }')
        # 创建运行结果显示标签
        self.label_result = QLabel('运行结果：', self)
        self.label_result.move(100, 150)
        self.label_result.setFont(QFont('Arial', 12))
        # 创建运行结果输入框
        self.line_state = QLineEdit(self)
        self.line_state.move(200, 150)
        self.line_state.resize(150, 30)
        self.line_state.setFont(QFont('Arial', 12))
    # 定义运行代码的槽函数
    @pyqtSlot()
    def run_code(self):
        # 执行合并功能的代码
        Q_result=merge_()
        save(Q_result,'Q_result')
        # 显示运行结果
        self.line_state.setText('运行完毕，请退出')

    # 定义返回按钮的槽函数
    @pyqtSlot()
    def go_back(self):
        self.MainWindow=MainWindow()
        self.MainWindow.show()
        self.hide()  # 隐藏当前窗口

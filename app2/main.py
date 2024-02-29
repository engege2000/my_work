import sys
from PyQt5.QtWidgets import QApplication
from UI import MainWindow
app = QApplication(sys.argv)
# 创建主窗口对象
window = MainWindow()
# 显示主窗口
window.show()
# 进入应用程序的主循环
sys.exit(app.exec_())
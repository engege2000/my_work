from UI import QApplication,MainWindow
import sys
# 创建应用程序对象
app = QApplication(sys.argv)
# 创建主窗口对象
main_window = MainWindow()
# 显示主窗口
main_window.show()
# 进入应用程序的主循环
sys.exit(app.exec_())
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget
from tab1 import Tab1Widget
from tab2 import Tab2Widget

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.options = []
        # 從檔案讀取選項
        with open('products.txt', 'r') as f:
            for line in f:
                # 每行一個選項，存儲在list中
                self.options.append(line.strip())
        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 720, 460)
        self.setWindowTitle('My Window')

        self.tabs = QTabWidget(self)
        self.tabs.setGeometry(2, 2, 718, 456)
        self.tab1 = Tab1Widget(self.options, self)
        self.tab2 = Tab2Widget(self.options, self)
        self.tabs.addTab(self.tab1, "進貨")
        self.tabs.addTab(self.tab2, "銷貨")
        # 设置选项卡标签的样式表
        self.tabs.setStyleSheet("QTabBar::tab { font-size: 16px; padding: 16px; margin: 0 0px; }")



        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyWindow()
    sys.exit(app.exec_())
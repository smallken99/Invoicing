from PyQt5.QtWidgets import QWidget, QPushButton


class Tab2Widget(QWidget):
    def __init__(self, options, parent):
        super().__init__(parent)
        self.options = options
        #print(options)
        self.initUI()


    def initUI(self):
        btn = QPushButton("Button 1", self)
        btn.move(50, 50)
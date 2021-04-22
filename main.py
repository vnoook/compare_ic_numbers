import PyQt5
import PyQt5.QtWidgets
import sys

class Window(PyQt5.QtWidgets.QMainWindow):
    def __init__(self):
        super(Window, self).__init__()

        self.setWindowTitle('first exp in QT5')
        self.setGeometry(300, 300, 400, 300)

        self.app_window_main_text = PyQt5.QtWidgets.QLabel(self)
        self.app_window_main_text.setText('__ proverka__sdfsdfsdfsdfsdfsdfsdfsdf')
        self.app_window_main_text.move(50, 50)
        self.app_window_main_text.adjustSize()

        self.app_window_main_btn1 = PyQt5.QtWidgets.QPushButton(self)
        self.app_window_main_btn1.setText('ok')
        self.app_window_main_btn1.move(150, 150)
        self.app_window_main_btn1.setFixedWidth(50)
        self.app_window_main_btn1.clicked.connect(self.change_label_text)

    def change_label_text(self):
        x = input('дай текст для надписи : ')
        self.app_window_main_text.setText(x)
        self.app_window_main_text.move(70, 70)
        self.app_window_main_text.adjustSize()
        self.setGeometry(300+100, 300+100, 400+100, 300+100)

# создание основного окна
def main_app():
    app = PyQt5.QtWidgets.QApplication(sys.argv)
    app_window_main = Window()

    app_window_main.show()
    sys.exit(app.exec_())


# запуск основного окна
if __name__ == '__main__':
    main_app()


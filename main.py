import PyQt5
import PyQt5.QtWidgets
import sys

# класс главного окна
class Window(PyQt5.QtWidgets.QMainWindow):
    # описание главного окна
    def __init__(self):
        super(Window, self).__init__()

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('first exp in QT5')
        self.setGeometry(300, 300, 400, 300)

        # объекты на главном окне
        # текст label
        self.app_window_main_text = PyQt5.QtWidgets.QLabel(self)
        self.app_window_main_text.setText(' __ proverka __ sdfsdfsdfsdfsdfsdfsdfsdf')
        self.app_window_main_text.move(50, 50)
        self.app_window_main_text.adjustSize()

        # кнопка button btn1
        self.app_window_main_btn1 = PyQt5.QtWidgets.QPushButton(self)
        self.app_window_main_btn1.setText('ok')
        self.app_window_main_btn1.move(150, 150)
        self.app_window_main_btn1.setFixedWidth(50)
        self.app_window_main_btn1.clicked.connect(self.change_label_text)

        # кнопка button btn2
        self.app_window_main_btn2 = PyQt5.QtWidgets.QPushButton(self)
        self.app_window_main_btn2.setText('EXIT')
        self.app_window_main_btn2.move(250, 250)
        self.app_window_main_btn2.setFixedWidth(50)
        self.app_window_main_btn2.clicked.connect(self.click_on_btn_exit)

    # событие нажатие на кнопку EXIT
    def click_on_btn_exit(selfself):
        exit()
        pass

    # событие нажатие на кнопку OK
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


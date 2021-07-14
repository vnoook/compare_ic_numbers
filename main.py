# button_select_file_IC
# button_select_file_GASPS
#
# pushButton_do_fill_data
#
#
#
# Произвести заполнение


import PyQt5
import PyQt5.QtWidgets
import sys


# класс главного окна
class Window(PyQt5.QtWidgets.QMainWindow):
    # описание главного окна
    def __init__(self):
        super(Window, self).__init__()

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Сравнение номеров дел')
        self.setGeometry(300, 300, 400, 300)

        # объекты на главном окне
        # label_select_file_IC
        self.label_select_file_IC = PyQt5.QtWidgets.QLabel(self)
        self.label_select_file_IC.setObjectName("label_select_file_IC")
        self.label_select_file_IC.setText('Выберите файл ИЦ')
        self.label_select_file_IC.move(10, 10)
        self.label_select_file_IC.adjustSize()
        # label_select_file_GASPS
        self.label_select_file_GASPS = PyQt5.QtWidgets.QLabel(self)
        self.label_select_file_GASPS.setObjectName("label_select_file_GASPS")
        self.label_select_file_GASPS.setText('Выберите файл ГАСПС')
        self.label_select_file_GASPS.move(10, 30)
        self.label_select_file_GASPS.adjustSize()
        # label_path_file_IC
        self.label_path_file_IC = PyQt5.QtWidgets.QLabel(self)
        self.label_path_file_IC.setObjectName("label_path_file_IC")
        self.label_path_file_IC.setText('label_path_file_IC')
        self.label_path_file_IC.move(10, 50)
        self.label_path_file_IC.adjustSize()
        # label_path_file_GASPS
        self.label_path_file_GASPS = PyQt5.QtWidgets.QLabel(self)
        self.label_path_file_GASPS.setObjectName("label_path_file_GASPS")
        self.label_path_file_GASPS.setText('label_path_file_GASPS')
        self.label_path_file_GASPS.move(10, 70)
        self.label_path_file_GASPS.adjustSize()

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
        self.app_window_main_btn2.clicked.connect(self.click_on_btn2_exit)

    # событие нажатие на кнопку EXIT
    def click_on_btn2_exit(self):
        exit()

    # событие нажатие на кнопку OK
    def change_label_text(self):
        new_text = input('Введите новый текст для надписи на форме : ')
        self.label_select_file_IC.setText(new_text)


# создание основного окна
def main_app():
    app = PyQt5.QtWidgets.QApplication(sys.argv)
    app_window_main = Window()

    app_window_main.show()
    sys.exit(app.exec_())


# запуск основного окна
if __name__ == '__main__':
    main_app()

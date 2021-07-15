
import PyQt5
import PyQt5.QtWidgets
import sys

info_path_open_file = r'd:\downloads\2020_09_29\1em\1 em 09-2020'
info_path_open_file = ''

# класс главного окна
class Window(PyQt5.QtWidgets.QMainWindow):
    # описание главного окна
    def __init__(self):
        super(Window, self).__init__()

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Сравнение номеров дел')
        self.setGeometry(300, 300, 900, 300)

        # объекты на главном окне
        # label_select_file_IC 'Выберите файл ИЦ'
        self.label_select_file_IC = PyQt5.QtWidgets.QLabel(self)
        self.label_select_file_IC.setObjectName('label_select_file_IC')
        self.label_select_file_IC.setText('1. Выберите файл ИЦ')
        self.label_select_file_IC.setGeometry(PyQt5.QtCore.QRect(10, 10, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_select_file_IC.setFont(font)
        self.label_select_file_IC.adjustSize()

        # label_select_file_GASPS 'Выберите файл ГАСПС'
        self.label_select_file_GASPS = PyQt5.QtWidgets.QLabel(self)
        self.label_select_file_GASPS.setObjectName('label_select_file_GASPS')
        self.label_select_file_GASPS.setText('2. Выберите файл ГАС ПС')
        self.label_select_file_GASPS.setGeometry(PyQt5.QtCore.QRect(10, 80, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_select_file_GASPS.setFont(font)
        self.label_select_file_GASPS.adjustSize()

        # label_path_file_IC
        self.label_path_file_IC = PyQt5.QtWidgets.QLabel(self)
        self.label_path_file_IC.setObjectName('label_path_file_IC')
        self.label_path_file_IC.setText('файл пока не выбран')
        self.label_path_file_IC.setGeometry(PyQt5.QtCore.QRect(70, 42, 820, 16))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.label_path_file_IC.setFont(font)
        self.label_path_file_IC.adjustSize()

        # label_path_file_GASPS
        self.label_path_file_GASPS = PyQt5.QtWidgets.QLabel(self)
        self.label_path_file_GASPS.setObjectName('label_path_file_GASPS')
        self.label_path_file_GASPS.setText('файл пока не выбран')
        self.label_path_file_GASPS.setGeometry(PyQt5.QtCore.QRect(70, 112, 820, 20))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.label_path_file_GASPS.setFont(font)
        self.label_path_file_GASPS.adjustSize()

        # toolButton_select_file_IC
        self.toolButton_select_file_IC = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_file_IC.setObjectName('toolButton_select_file_IC')
        self.toolButton_select_file_IC.setText('...')
        self.toolButton_select_file_IC.setGeometry(PyQt5.QtCore.QRect(10, 40, 50, 20))
        self.toolButton_select_file_IC.setFixedWidth(50)
        self.toolButton_select_file_IC.clicked.connect(self.select_file)

        # toolButton_select_file_GASPS
        self.toolButton_select_file_GASPS = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_file_GASPS.setObjectName('toolButton_select_file_GASPS')
        self.toolButton_select_file_GASPS.setText('...')
        self.toolButton_select_file_GASPS.setGeometry(PyQt5.QtCore.QRect(10, 110, 50, 20))
        self.toolButton_select_file_GASPS.setFixedWidth(50)
        self.toolButton_select_file_GASPS.clicked.connect(self.select_file)

        # pushButton_do_fill_data
        self.pushButton_do_fill_data = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_do_fill_data.setObjectName('pushButton_do_fill_data')
        self.pushButton_do_fill_data.setText('Произвести заполнение')
        self.pushButton_do_fill_data.setGeometry(PyQt5.QtCore.QRect(10, 150, 180, 25))
        self.pushButton_do_fill_data.setFixedWidth(130)
        self.pushButton_do_fill_data.clicked.connect(self.do_fill_data)

        # # кнопка button btn1
        # self.app_window_main_btn1 = PyQt5.QtWidgets.QPushButton(self)
        # self.app_window_main_btn1.setText('ok')
        # self.app_window_main_btn1.move(150, 150)
        # self.app_window_main_btn1.setFixedWidth(50)
        # self.app_window_main_btn1.clicked.connect(self.change_label_text)

        # кнопка button_exit
        self.button_exit = PyQt5.QtWidgets.QPushButton(self)
        self.button_exit.setObjectName('button_exit')
        self.button_exit.setText('EXIT')
        self.button_exit.setGeometry(PyQt5.QtCore.QRect(10, 200, 180, 25))
        self.button_exit.setFixedWidth(50)
        self.button_exit.clicked.connect(self.click_on_btn_exit)


    # событие нажатие на кнопку выбора файла
    def select_file(self):
        print()
        print(f'выбрать файл с кнопки {self.sender().objectName()}')
        print()

        # запоминание старого значения пути выбора файлов
        old_path_of_selected_file_IC = self.label_path_file_IC.text()
        old_path_of_selected_file_GASPS = self.label_path_file_GASPS.text()
        print(f'путь до выбора файла {old_path_of_selected_file_IC = }')
        print(f'путь до выбора файла {old_path_of_selected_file_GASPS = }')
        print()

        # определение какая кнопка выбора файла нажата
        # если ИЦ, то выдать в окно про ИЦ
        if self.sender().objectName() == self.toolButton_select_file_IC.objectName():
            info_for_open_file = 'Выберите файл ИЦ формата Excel, версии старше 2007 года (.XLSX)'
        # если ГАСПС, то выдать в окно про ГАСПС
        elif self.sender().objectName() == self.toolButton_select_file_GASPS.objectName():
            info_for_open_file = 'Выберите файл ГАС ПС формата Excel, версии старше 2007 года (.XLSX)'

        # непосредственное окно выбора файла и переменная для хранения пути файла
        data_of_open_file_name = PyQt5.QtWidgets.QFileDialog.getOpenFileName(self,
                                                                info_for_open_file,
                                                                info_path_open_file,
                                                                'Файлы Excel xlsx (*.xlsx)'
                                                                )
        # вычленение пути файла из data_of_open_file_name
        file_name = data_of_open_file_name[0]
        print(f'выбран файл {file_name = }')
        print()

        # выбор где и что менять исходя из выбора пользователя
        if self.sender().objectName() == self.toolButton_select_file_IC.objectName():
            self.label_path_file_IC.setText(file_name)
            self.label_path_file_IC.adjustSize()

        elif self.sender().objectName() == self.toolButton_select_file_GASPS.objectName():
            self.label_path_file_GASPS.setText(file_name)
            self.label_path_file_GASPS.adjustSize()

        elif file_name == '':
            if self.sender().objectName() == self.toolButton_select_file_IC.objectName():
                self.label_path_file_IC.setText(old_path_of_selected_file_IC)
                self.label_path_file_IC.adjustSize()
            elif self.sender().objectName() == self.toolButton_select_file_GASPS.objectName():
                self.label_path_file_GASPS.setText(old_path_of_selected_file_GASPS)
                self.label_path_file_GASPS.adjustSize()
        else:
            # self.sender().objectName() == old_path_of_selected_file
            pass

        print(f'путь после выбора файла {old_path_of_selected_file_IC = }')
        print(f'путь после выбора файла {old_path_of_selected_file_GASPS = }')
        print()


    # событие нажатие на кнопку заполнения файла из ИЦ
    def do_fill_data(self):
        print(f'\nнажата кнопка {self.sender().objectName()}')


    # событие нажатие на кнопку EXIT
    def click_on_btn_exit(self):
        exit()


# создание основного окна
def main_app():
    app = PyQt5.QtWidgets.QApplication(sys.argv)
    app_window_main = Window()

    app_window_main.show()
    sys.exit(app.exec_())


# запуск основного окна
if __name__ == '__main__':
    main_app()

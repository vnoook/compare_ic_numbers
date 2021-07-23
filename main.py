# ...
# INSTALL
# pip install openpyxl
# INSTALL
# pip install PyQt5
# ...

import PyQt5
import PyQt5.QtWidgets
import sys
import openpyxl
import openpyxl.styles

# класс главного окна
class Window(PyQt5.QtWidgets.QMainWindow):
    # описание главного окна
    def __init__(self):
        super(Window, self).__init__()

        # переменные
        self.info_for_open_file = ''
        self.info_path_open_file = ''
        self.info_extention_open_file = 'Файлы Excel xlsx (*.xlsx)'
        self.text_empty_path_file = 'файл пока не выбран'
        self.file_IC = ''
        self.file_GASPS = ''

        # главное окно, надпись на нём и размеры
        self.setWindowTitle('Сравнение номеров дел')
        self.setGeometry(300, 300, 900, 300)

        # объекты на главном окне
        # label_select_file_IC
        self.label_select_file_IC = PyQt5.QtWidgets.QLabel(self)
        self.label_select_file_IC.setObjectName('label_select_file_IC')
        self.label_select_file_IC.setText('1. Выберите файл ИЦ')
        self.label_select_file_IC.setGeometry(PyQt5.QtCore.QRect(10, 10, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_select_file_IC.setFont(font)
        self.label_select_file_IC.adjustSize()
        self.label_select_file_IC.setToolTip(self.label_select_file_IC.objectName())

        # toolButton_select_file_IC
        self.toolButton_select_file_IC = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_file_IC.setObjectName('toolButton_select_file_IC')
        self.toolButton_select_file_IC.setText('...')
        self.toolButton_select_file_IC.setGeometry(PyQt5.QtCore.QRect(10, 40, 50, 20))
        self.toolButton_select_file_IC.setFixedWidth(50)
        self.toolButton_select_file_IC.clicked.connect(self.select_file)
        self.toolButton_select_file_IC.setToolTip(self.toolButton_select_file_IC.objectName())

        # label_path_file_IC
        self.label_path_file_IC = PyQt5.QtWidgets.QLabel(self)
        self.label_path_file_IC.setObjectName('label_path_file_IC')
        self.label_path_file_IC.setText(self.text_empty_path_file)
        self.label_path_file_IC.setGeometry(PyQt5.QtCore.QRect(70, 40, 820, 16))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.label_path_file_IC.setFont(font)
        self.label_path_file_IC.adjustSize()
        self.label_path_file_IC.setToolTip(self.label_path_file_IC.objectName())

        # comboBox_liter_IC
        self.comboBox_liter_IC = PyQt5.QtWidgets.QComboBox(self)
        self.comboBox_liter_IC.setObjectName('comboBox_liter_IC')
        self.comboBox_liter_IC.setGeometry(PyQt5.QtCore.QRect(10, 70, 70, 20))
        self.comboBox_liter_IC.addItem('Колонка')
        self.comboBox_liter_IC.setEnabled(False)
        self.comboBox_liter_IC.setToolTip(self.comboBox_liter_IC.objectName())

        # comboBox_digit_IC
        self.comboBox_digit_IC = PyQt5.QtWidgets.QComboBox(self)
        self.comboBox_digit_IC.setObjectName('comboBox_digit_IC')
        self.comboBox_digit_IC.setGeometry(PyQt5.QtCore.QRect(90, 70, 70, 20))
        self.comboBox_digit_IC.addItem('Строка')
        self.comboBox_digit_IC.setEnabled(False)
        self.comboBox_digit_IC.setToolTip(self.comboBox_digit_IC.objectName())

        # label_select_file_GASPS
        self.label_select_file_GASPS = PyQt5.QtWidgets.QLabel(self)
        self.label_select_file_GASPS.setObjectName('label_select_file_GASPS')
        self.label_select_file_GASPS.setText('2. Выберите файл ГАС ПС')
        self.label_select_file_GASPS.setGeometry(PyQt5.QtCore.QRect(10, 120, 150, 40))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(12)
        self.label_select_file_GASPS.setFont(font)
        self.label_select_file_GASPS.adjustSize()
        self.label_select_file_GASPS.setToolTip(self.label_select_file_GASPS.objectName())

        # label_path_file_GASPS
        self.label_path_file_GASPS = PyQt5.QtWidgets.QLabel(self)
        self.label_path_file_GASPS.setObjectName('label_path_file_GASPS')
        self.label_path_file_GASPS.setText(self.text_empty_path_file)
        self.label_path_file_GASPS.setGeometry(PyQt5.QtCore.QRect(70, 150, 820, 20))
        font = PyQt5.QtGui.QFont()
        font.setPointSize(10)
        self.label_path_file_GASPS.setFont(font)
        self.label_path_file_GASPS.adjustSize()
        self.label_path_file_GASPS.setToolTip(self.label_path_file_GASPS.objectName())

        # toolButton_select_file_GASPS
        self.toolButton_select_file_GASPS = PyQt5.QtWidgets.QPushButton(self)
        self.toolButton_select_file_GASPS.setObjectName('toolButton_select_file_GASPS')
        self.toolButton_select_file_GASPS.setText('...')
        self.toolButton_select_file_GASPS.setGeometry(PyQt5.QtCore.QRect(10, 150, 50, 20))
        self.toolButton_select_file_GASPS.setFixedWidth(50)
        self.toolButton_select_file_GASPS.clicked.connect(self.select_file)
        self.toolButton_select_file_GASPS.setToolTip(self.toolButton_select_file_GASPS.objectName())



        # comboBox_liter_GASPS
        self.comboBox_liter_GASPS = PyQt5.QtWidgets.QComboBox(self)
        self.comboBox_liter_GASPS.setObjectName('comboBox_liter_GASPS')
        self.comboBox_liter_GASPS.setGeometry(PyQt5.QtCore.QRect(10, 180, 70, 20))
        self.comboBox_liter_GASPS.addItem('Колонка')
        self.comboBox_liter_GASPS.setEnabled(False)
        self.comboBox_liter_GASPS.setToolTip(self.comboBox_liter_GASPS.objectName())

        # comboBox_digit_GASPS
        self.comboBox_digit_GASPS = PyQt5.QtWidgets.QComboBox(self)
        self.comboBox_digit_GASPS.setObjectName('comboBox_digit_GASPS')
        self.comboBox_digit_GASPS.setGeometry(PyQt5.QtCore.QRect(90, 180, 70, 20))
        self.comboBox_digit_GASPS.addItem('Строка')
        self.comboBox_digit_GASPS.setEnabled(False)
        self.comboBox_digit_GASPS.setToolTip(self.comboBox_digit_GASPS.objectName())

        # pushButton_do_fill_data
        self.pushButton_do_fill_data = PyQt5.QtWidgets.QPushButton(self)
        self.pushButton_do_fill_data.setObjectName('pushButton_do_fill_data')
        self.pushButton_do_fill_data.setEnabled(False)
        self.pushButton_do_fill_data.setText('Произвести заполнение')
        self.pushButton_do_fill_data.setGeometry(PyQt5.QtCore.QRect(10, 225, 180, 25))
        self.pushButton_do_fill_data.setFixedWidth(130)
        self.pushButton_do_fill_data.clicked.connect(self.do_fill_data)
        self.pushButton_do_fill_data.setToolTip(self.pushButton_do_fill_data.objectName())

        # кнопка button_exit
        self.button_exit = PyQt5.QtWidgets.QPushButton(self)
        self.button_exit.setObjectName('button_exit')
        self.button_exit.setText('Выход')
        self.button_exit.setGeometry(PyQt5.QtCore.QRect(10, 260, 180, 25))
        self.button_exit.setFixedWidth(50)
        self.button_exit.clicked.connect(self.click_on_btn_exit)
        self.button_exit.setToolTip(self.button_exit.objectName())

    # событие - нажатие на кнопку выбора файла
    def select_file(self):
        # запоминание старого значения пути выбора файлов
        old_path_of_selected_file_IC = self.label_path_file_IC.text()
        old_path_of_selected_file_GASPS = self.label_path_file_GASPS.text()

        # определение какая кнопка выбора файла нажата
        # если ИЦ, то выдать в окно про ИЦ
        if self.sender().objectName() == self.toolButton_select_file_IC.objectName():
            self.info_for_open_file = 'Выберите файл ИЦ формата Excel, версии старше 2007 года (.XLSX)'
        # если ГАСПС, то выдать в окно про ГАСПС
        elif self.sender().objectName() == self.toolButton_select_file_GASPS.objectName():
            self.info_for_open_file = 'Выберите файл ГАС ПС формата Excel, версии старше 2007 года (.XLSX)'

        # непосредственное окно выбора файла и переменная для хранения пути файла
        data_of_open_file_name = PyQt5.QtWidgets.QFileDialog.getOpenFileName(self,
                                                                             self.info_for_open_file,
                                                                             self.info_path_open_file,
                                                                             self.info_extention_open_file)
        # вычленение пути файла из data_of_open_file_name
        file_name = data_of_open_file_name[0]

        # выбор где и что менять исходя из выбора пользователя
        # нажата кнопка выбора ИЦ
        if self.sender().objectName() == self.toolButton_select_file_IC.objectName():
            if file_name == '':
                self.label_path_file_IC.setText(old_path_of_selected_file_IC)
                self.label_path_file_IC.adjustSize()
            else:
                old_path_of_selected_file_IC = self.label_path_file_IC.text()

                self.label_path_file_IC.setText(file_name)
                self.label_path_file_IC.adjustSize()

        # нажата кнопка выбора ГАСПС
        if self.sender().objectName() == self.toolButton_select_file_GASPS.objectName():
            if file_name == '':
                self.label_path_file_GASPS.setText(old_path_of_selected_file_GASPS)
                self.label_path_file_GASPS.adjustSize()
            else:
                old_path_of_selected_file_GASPS = self.label_path_file_GASPS.text()

                self.label_path_file_GASPS.setText(file_name)
                self.label_path_file_GASPS.adjustSize()

        if self.label_path_file_IC.text() != self.label_path_file_GASPS.text():
            if self.text_empty_path_file not in (self.label_path_file_IC.text(), self.label_path_file_GASPS.text()):
                self.pushButton_do_fill_data.setEnabled(True)
                self.comboBox_liter_IC.setEnabled(True)
                self.comboBox_digit_IC.setEnabled(True)
                self.comboBox_liter_GASPS.setEnabled(True)
                self.comboBox_digit_GASPS.setEnabled(True)
        else:
            self.pushButton_do_fill_data.setEnabled(False)
            self.comboBox_liter_IC.setEnabled(False)
            self.comboBox_digit_IC.setEnabled(False)
            self.comboBox_liter_GASPS.setEnabled(False)
            self.comboBox_digit_GASPS.setEnabled(False)

    # событие - нажатие на кнопку заполнения файла
    def do_fill_data(self):
        self.file_IC = self.label_path_file_IC.text()
        self.file_GASPS = self.label_path_file_GASPS.text()

        # TODO
        # открывается файл "приёмник", назначается активный лист, выбирается диапазон ячеек
        # wb_narush = openpyxl.load_workbook(file_template_xl)
        # wb_narush_s = wb_narush.active
        # wb_narush_cells_range = wb_narush_s[template_cells_range]
        #
        # print()
        # print(f'файл {self.file_IC    = }')
        # print(f'файл {self.file_GASPS = }')
        # print(f'нажата кнопка {self.sender().objectName() = }')

    # событие - нажатие на кнопку Выход
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

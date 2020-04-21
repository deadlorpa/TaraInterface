from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QWidget, QApplication, QDesktopWidget, QTextEdit, QGridLayout, QTableWidgetItem
from PyQt5.QtWidgets import QPushButton, QHBoxLayout, QVBoxLayout, QScrollArea, QLabel
from PyQt5.QtWidgets import QMenu, QMenuBar, QFileDialog, QToolBar
from app import Ui_MainWindow  # импорт нашего сгенерированного файла
import sys
from openpyxl import load_workbook


class mywindow(QtWidgets.QMainWindow):
    okyd = '0330230'  # код по окуд
    key_no_update_lines = False
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.current_state_period = self.ui.comboBox.currentText()
        self.current_state_case = self.ui.comboBox_2.currentText()
        self.ui.CreateRow.clicked.connect(self.CreateRowClicked)
        self.ui.CreateButton.clicked.connect(self.CreateButtonClicked)
        self.ui.textEdit.textChanged.connect(self.UpdateHead)
        self.ui.date_create.dateTimeChanged.connect(self.UpdateHead)
        self.ui.comboBox.currentTextChanged.connect(self.UpdatePeriod)
        self.ui.comboBox_2.currentTextChanged.connect(self.UpdateCase)
        self.ui.tableWidget.insertRow(self.ui.tableWidget.rowCount())
        self.UpdateHead()
        row = self.ui.tableWidget.rowCount()
        for i in range(13):
            item = QTableWidgetItem()
            item.setText(" ")
            self.ui.tableWidget.setItem(row-1, i, item)
        #  ТАРА
        self.line0 = [""]  # наименование
        self.line1 = [""]  # код
        #  ПОСТАВЩИК
        self.line2 = [""]  # наименование
        self.line3 = [""]  # код
        #  ЦЕНА
        self.line4 = [""]  # цена
        #  НАЧАЛО ПЕРИОДА
        self.line5 = [""]  # колво
        self.line6 = [""]  # сумма
        #  ПРИХОД
        self.line7 = [""]  # колво
        self.line8 = [""]  # сумма
        #  РАСХОД
        self.line9 = [""]  # колво
        self.line10 = [""]  # сумма
        #  КОНЕЦ ПЕРИОДА
        self.line11 = [""]  # колво
        self.line12 = [""]  # сумма

        self.ui.tableWidget.cellChanged.connect(self.TableRouteen)

    # обновление данных в таблицах
    def TableRouteen(self):
        if not self.key_no_update_lines:
            row = self.ui.tableWidget.currentRow()
            #  ТАРА
            self.line0[row] = self.ui.tableWidget.item(row, 0).text()
            self.line1[row] = self.ui.tableWidget.item(row, 1).text()
            #  ПОСТАВЩИК
            self.line2[row] = self.ui.tableWidget.item(row, 2).text()
            self.line3[row] = self.ui.tableWidget.item(row, 3).text()
            #  ЦЕНА
            self.line4[row] = self.ui.tableWidget.item(row, 4).text()
            if self.current_state_period == "Начало периода":
                #  НАЧАЛО ПЕРИОДА
                self.line5[row] = self.ui.tableWidget.item(row, 5).text()
                self.line6[row] = self.ui.tableWidget.item(row, 6).text()
            if self.current_state_period == "Конец периода":
                #  КОНЕЦ ПЕРИОДА
                self.line11[row] = self.ui.tableWidget.item(row, 5).text()
                self.line12[row] = self.ui.tableWidget.item(row, 6).text()
            if self.current_state_case == "Приход":
                #  ПРИХОД
                self.line7[row] = self.ui.tableWidget.item(row, 7).text()
                self.line8[row] = self.ui.tableWidget.item(row, 8).text()
            if self.current_state_case == "Расход":
                #  РАСХОД
                self.line9[row] = self.ui.tableWidget.item(row, 7).text()
                self.line10[row] = self.ui.tableWidget.item(row, 8).text()

    # таблица
    def TableRowAdd(self):
        #  ТАРА
        self.line0.append("")
        self.line1.append("")
        #  ПОСТАВЩИК
        self.line2.append("")
        self.line3.append("")
        #  ЦЕНА
        self.line4.append("")
        #  НАЧАЛО ПЕРИОДА
        self.line5.append("")
        self.line6.append("")
        #  ПРИХОД
        self.line7.append("")
        self.line8.append("")
        #  РАСХОД
        self.line9.append("")
        self.line10.append("")
        #  КОНЕЦ ПЕРИОДА
        self.line11.append("")
        self.line12.append("")

    # обновление таблицы приход-расход
    def UpdateCase(self):
        if self.current_state_case != self.ui.comboBox_2.currentText():
            self.key_no_update_lines = True
            if self.ui.comboBox_2.currentText() == "Приход":
                self.current_state_case = "Приход"
                self.ui.label_of_period_2.setText("Приход")
                for i in range(self.ui.tableWidget.rowCount()):
                    self.ui.tableWidget.setItem(i, 7, QTableWidgetItem(self.line7[i]))
                    self.ui.tableWidget.setItem(i, 8, QTableWidgetItem(self.line8[i]))
            else:
                self.current_state_case = "Расход"
                self.ui.label_of_period_2.setText("Расход")
                for i in range(self.ui.tableWidget.rowCount()):
                    self.ui.tableWidget.setItem(i, 7, QTableWidgetItem(self.line9[i]))
                    self.ui.tableWidget.setItem(i, 8, QTableWidgetItem(self.line10[i]))
            self.key_no_update_lines = False
    # обновление таблицы начало-конец периода
    def UpdatePeriod(self):
        if self.current_state_period != self.ui.comboBox.currentText():
            self.key_no_update_lines = True
            if self.ui.comboBox.currentText() == "Начало периода":
                self.current_state_period = "Начало периода"
                self.ui.label_of_period.setText("Начало периода")
                for i in range(self.ui.tableWidget.rowCount()):
                    self.ui.tableWidget.setItem(i, 5, QTableWidgetItem(self.line5[i]))
                    self.ui.tableWidget.setItem(i, 6, QTableWidgetItem(self.line6[i]))
            else:
                self.current_state_period = "Конец периода"
                self.ui.label_of_period.setText("Конец периода")
                for i in range(self.ui.tableWidget.rowCount()):
                    self.ui.tableWidget.setItem(i, 5, QTableWidgetItem(self.line11[i]))
                    self.ui.tableWidget.setItem(i, 6, QTableWidgetItem(self.line12[i]))
            self.key_no_update_lines = False
    # обновление шапки
    def UpdateHead(self):
        num = self.ui.textEdit.toPlainText()
        from_ = self.ui.date_create.text()
        self.ui.label.setText('Отчет по таре: документ ' + num + ' от ' + from_)

    # обработчик кнопки СОЗДАТЬ
    def CreateButtonClicked(self):
        self.organization = self.ui.org_text.toPlainText()
        self.podrazdelenie = self.ui.podraz_text.toPlainText()
        self.number_doc = self.ui.textEdit.toPlainText()
        self.date_start = self.ui.date_begin.text()
        self.date_end = self.ui.date_end.text()
        self.date_create = self.ui.date_create.text()
        self.okpo = self.ui.code_text.toPlainText()
        self.okdp = self.ui.okdp_text.toPlainText()
        self.type_active = self.ui.deya_text.toPlainText()
        self.otv_face = self.ui.who_text.toPlainText()
        self.dolz = self.ui.work_text.toPlainText()
        self.tabel_num = self.ui.tabel_text.toPlainText()

        workbook = load_workbook('example.xlsx')
        sheet = workbook['стр1']

        sheet['A6'] = self.organization
        sheet['A8'] = self.podrazdelenie
        sheet['G15'] = self.dolz + ' ' + self.otv_face
        sheet['AE15'] = self.tabel_num

        sheet['AD5'] = self.okyd
        sheet['AD6'] = self.okpo
        sheet['AD7'] = self.okpo
        sheet['AD9'] = self.okdp
        sheet['AD11'] = self.type_active

        sheet['P14'] = self.number_doc
        sheet['U14'] = self.date_create[:-5]
        sheet['W14'] = self.date_start
        sheet['Y14'] = self.date_end

        rows = len(self.line0)
        for i in range(rows):
            if i <= 6:
                row = str(i + 23)
                sheet['A' + row] = i+1
                sheet['B' + row] = self.line0[i]
                sheet['F' + row] = self.line1[i]
                sheet['H' + row] = self.line2[i]
                sheet['J' + row] = self.line3[i]
                sheet['K' + row] = self.line4[i]
                sheet['M' + row] = self.line5[i]
                sheet['P' + row] = self.line6[i]
                sheet['U' + row] = self.line7[i]
                sheet['V' + row] = self.line8[i]
                sheet['X' + row] = self.line9[i]
                sheet['Z' + row] = self.line10[i]
                sheet['AD' + row] = self.line11[i]
                sheet['AG' + row] = self.line12[i]

        #суммы
        self.line5 = [int(i) for i in self.line5]
        sheet['M29'] = sum(self.line5)
        self.line6 = [int(i) for i in self.line6]
        sheet['P29'] = sum(self.line6)
        self.line7 = [int(i) for i in self.line7]
        sheet['U29'] = sum(self.line7)
        self.line8 = [int(i) for i in self.line8]
        sheet['V29'] = sum(self.line8)
        self.line9 = [int(i) for i in self.line9]
        sheet['X29'] = sum(self.line9)
        self.line10 = [int(i) for i in self.line10]
        sheet['Z29'] = sum(self.line10)
        self.line11 = [int(i) for i in self.line11]
        sheet['AD29'] = sum(self.line11)
        self.line12 = [int(i) for i in self.line12]
        sheet['AG29'] = sum(self.line12)

        workbook.save('last.xlsx')
        print('saved!')


    # обработчик кнопки ДОБАВИТЬ ЗАПИСЬ
    def CreateRowClicked(self):
        self.key_no_update_lines = True
        self.ui.tableWidget.insertRow(self.ui.tableWidget.rowCount())
        row = self.ui.tableWidget.rowCount()
        for i in range(13):
            item = QTableWidgetItem()
            item.setText(" ")
            self.ui.tableWidget.setItem(row-1, i, item)
        self.TableRowAdd()
        self.key_no_update_lines = False

app = QtWidgets.QApplication([])
application = mywindow()
application.show()

sys.exit(app.exec())

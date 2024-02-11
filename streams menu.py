# libraries for creating GUI
from PyQt5 import QtWidgets, QtCore
# libraries for calculations and graphics
import numpy as np
from scipy.optimize import fsolve
import scipy.integrate as spi
import matplotlib.pyplot as plt
# libraries for other purposes
import sys
import openpyxl
import traceback

# useful constants
id_name_properties = None  # will be defined from the library file with components
conditions_names = ["Temperature [C]", "Pressure [kPa]", "Flow Rate [kg/sec]"]


class StartMenu(QtWidgets.QWidget):

    def __init__(self):
        super(StartMenu, self).__init__()  # function that passes the properties of the super class to child class?
        self.init_UI()

    def init_UI(self):
        self.setWindowTitle("Start Menu")
        self.resize(300, 300)

        self.create_btn_create()
        self.create_btn_open()

    def create_btn_create(self):
        self.btn_create = QtWidgets.QPushButton(self)
        self.btn_create.setText("Create...")
        self.btn_create.setGeometry(15, 15, 90, 30)

        self.btn_create.clicked.connect(self.create_worksheet)

    def create_worksheet(self):
        self.work_sheet = WorkSheet()
        self.work_sheet.show()
        self.close()

    def create_btn_open(self):
        self.btn_open = QtWidgets.QPushButton(self)
        self.btn_open.setText("Open...")
        self.btn_open.setGeometry(15, 55, 90, 30)


class WorkSheet(QtWidgets.QWidget):
    streams_dict = {}
    windows_dict = {}

    def __init__(self):
        super(WorkSheet, self).__init__()
        self.init_UI()

    def init_UI(self):
        self.setWindowTitle("Work Sheet")
        self.resize(1050, 700)

        self.create_labels()
        self.create_btns_add()
        self.create_streams_list()
        self.create_elements_list()

    def create_labels(self):
        self.label_streams = QtWidgets.QLabel(self)
        self.label_streams.setText("Streams")
        self.label_streams.setGeometry(600, 0, 200, 20)
        self.label_streams.setStyleSheet("background-color: rgb(186, 174, 255);")
        self.label_streams.setAlignment(QtCore.Qt.AlignCenter)

        self.label_elements = QtWidgets.QLabel(self)
        self.label_elements.setText("Elements")
        self.label_elements.setGeometry(QtCore.QRect(820, 0, 200, 20))
        self.label_elements.setStyleSheet("background-color: rgb(186, 174, 255);")
        self.label_elements.setAlignment(QtCore.Qt.AlignCenter)

    def create_btns_add(self):
        self.create_btn_add_stream()
        self.create_btn_add_element()

    def create_btn_add_stream(self):
        self.btn_add_stream = QtWidgets.QPushButton(self)
        self.btn_add_stream.setText("Add...")
        self.btn_add_stream.setGeometry(600, 580, 200, 30)

        self.btn_add_stream.clicked.connect(self.open_component_library)

    def open_component_library(self):
        self.component_library = ComponentsLibrary(start_menu.work_sheet.streams_list.count())
        self.component_library.show()

    def create_btn_add_element(self):
        self.btn_elements_add = QtWidgets.QPushButton(self)
        self.btn_elements_add.setText("Add...")
        self.btn_elements_add.setGeometry(820, 580, 200, 30)

    def create_streams_list(self):
        self.streams_list = QtWidgets.QListWidget(self)
        self.streams_list.setGeometry(600, 20, 200, 550)

        self.streams_list.itemDoubleClicked.connect(self.open_stream_properties)

    def open_stream_properties(self, item):  # method itemDoubleClicked has information about clicked item and can pass it to function
        stream_name = item.text()

        if stream_name not in self.windows_dict.keys():  # checks if a StreamProperties Window has been created for this stream
            self.windows_dict[stream_name] = StreamProperties(stream_name)
        self.windows_dict[stream_name].show()

    def create_elements_list(self):
        self.elements_list = QtWidgets.QListWidget(self)
        self.elements_list.setGeometry(820, 20, 200, 550)


class ComponentsLibrary(QtWidgets.QWidget):

    def __init__(self, num_streams=0):
        super(ComponentsLibrary, self).__init__()
        self.num_streams = num_streams
        self.component_set = set()

        self.init_UI()

    def init_UI(self):
        self.resize(800, 500)
        self.setWindowTitle("Component Library")

        self.create_menubar()
        self.create_statusbar()
        self.create_library_table()
        self.create_current_stream_table()
        self.create_btn_add_component_to_stream()
        self.create_btn_remove_component_from_stream()
        self.create_btn_add_stream_to_worksheet()

    def create_menubar(self):
        self.menubar = QtWidgets.QMenuBar(self)

        self.file_menu = QtWidgets.QMenu(self.menubar)
        self.file_menu.setTitle("File")
        self.menubar.addMenu(self.file_menu)

        self.openFile = QtWidgets.QAction('Open', self)
        self.openFile.setShortcut('Ctrl+Q')
        self.openFile.triggered.connect(self.open_library_file)

        self.file_menu.addAction(self.openFile)

    def open_library_file(self):
        # standard_folder = 'D:\Chemical programming\Analog Hysys\Start Menu + Flow\Component Library Excel'
        # self.library_name = QtWidgets.QFileDialog.getOpenFileName(self, "Open Library File", standard_folder)[0]
        self.library_name = "D:\Chemical programming\Analog Hysys\Start Menu + Flow\Component Library Excel\Hydrocarbons C1-C10.xlsx"
        self.load_data_library_table()

        self.statusbar.showMessage("Add components to stream")
        self.statusbar.setStyleSheet("QStatusBar{background:rgb(0,255,0);color:black;font-weight:bold;}")

        self.current_stream_table.setRowCount(0)
        self.current_stream_table.setColumnCount(self.cols_number)

        global id_name_properties
        self.current_stream_table.setHorizontalHeaderLabels(id_name_properties)
        self.current_stream_table.resizeColumnsToContents()
        self.current_stream_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

    def create_statusbar(self):
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.showMessage("Open Library File")
        self.statusbar.setStyleSheet("QStatusBar{background:rgb(255,0,0);color:black;font-weight:bold;}")
        self.statusbar.setGeometry(0, 480, 800, 20)

    def create_library_table(self):
        self.library_table = QtWidgets.QTableWidget(self)
        self.library_table.setGeometry(20, 40, 300, 400)

    def load_data_library_table(self):
        path = self.library_name
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        # code block below this doesn't save empty cells (I don't know how to save xlsx-file without "None" values)
        list_values_excel = list(sheet.values)
        list_values = []
        for row_index in range(len(list_values_excel)):
            i_list = []
            for cell_value in list_values_excel[row_index]:
                if cell_value is not None:
                    i_list.append(cell_value)
            list_values.append(i_list)
        # zero row - these are general table headers that don't contain important information (solely for better visualization in Excel)
        self.rows_number = len(list_values[2:])  # components number in library file
        self.cols_number = len(list_values[1])  # columns number is equal number of headers (component characteristics)
        # component characteristics will also be used further, so I save those
        global id_name_properties
        id_name_properties = list_values[1]

        self.library_table.setRowCount(self.rows_number)
        self.library_table.setColumnCount(self.cols_number)

        self.library_table.setHorizontalHeaderLabels(id_name_properties)

        row_index = 0
        for values_list in list_values[2:]:
            col_index = 0
            for value in values_list:
                value = QtWidgets.QTableWidgetItem(str(value))
                value.setTextAlignment(QtCore.Qt.AlignCenter)
                self.library_table.setItem(row_index, col_index, QtWidgets.QTableWidgetItem(value))
                col_index += 1
            row_index += 1

        self.library_table.resizeColumnsToContents()
        self.library_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)  # cells can not be changed now

    def create_current_stream_table(self):
        self.current_stream_table = QtWidgets.QTableWidget(self)
        self.current_stream_table.setGeometry(480, 40, 300, 400)

    def create_btn_add_component_to_stream(self):
        self.btn_add_component_to_stream = QtWidgets.QPushButton(self)
        self.btn_add_component_to_stream.setGeometry(360, 160, 80, 40)
        self.btn_add_component_to_stream.setText("Add")

        self.btn_add_component_to_stream.clicked.connect(self.add_component_to_stream)

    def add_component_to_stream(self):
        current_row = self.library_table.currentRow()
        if current_row != -1:  # if user don't choose the cell, current row is equal -1
            # id_name = group + id
            id_name = self.library_table.item(current_row, 0).text() + str(self.library_table.item(current_row, 1).text())
            if id_name not in self.component_set:
                num_of_rows = self.current_stream_table.rowCount() + 1
                self.current_stream_table.setRowCount(num_of_rows)
                row_for_add = num_of_rows - 1
                for col_index in range(0, self.current_stream_table.columnCount()):
                    item = self.library_table.item(current_row, col_index).text()
                    if col_index == 1:
                        value = QtWidgets.QTableWidgetItem()
                        value.setData(QtCore.Qt.EditRole, int(item))
                    else:
                        value = QtWidgets.QTableWidgetItem(str(item))
                    value.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.current_stream_table.setItem(row_for_add, col_index, QtWidgets.QTableWidgetItem(value))
                self.component_set.add(id_name)

        self.current_stream_table.sortItems(1, QtCore.Qt.AscendingOrder)
        self.current_stream_table.resizeColumnsToContents()

    def create_btn_remove_component_from_stream(self):
        self.btn_remove_component_from_stream = QtWidgets.QPushButton(self)
        self.btn_remove_component_from_stream.setGeometry(360, 240, 80, 40)
        self.btn_remove_component_from_stream.setText("Remove")

        self.btn_remove_component_from_stream.clicked.connect(self.remove_component_from_stream)

    def remove_component_from_stream(self):
        current_row = self.current_stream_table.currentRow()
        self.current_stream_table.removeRow(current_row)

    def create_btn_add_stream_to_worksheet(self):
        self.btn_add_stream = QtWidgets.QPushButton(self)
        self.btn_add_stream.setGeometry(480, 440, 300, 40)
        self.btn_add_stream.setText("Add stream")

        self.btn_add_stream.clicked.connect(self.add_stream_to_worksheet)

    def add_stream_to_worksheet(self):
        stream_name = f"Stream {str(self.num_streams + 1)}"
        # adding stream in streams list
        if start_menu.work_sheet.component_library.current_stream_table.rowCount() == 0:  # check if stream is empty
            message_empty_current_stream_table = QtWidgets.QMessageBox(self)
            message_empty_current_stream_table.setWindowTitle("Error")
            message_empty_current_stream_table.setText("You haven't added any components")
            message_empty_current_stream_table.show()
        else:
            component_list = ComponentList(stream_name)  # saving information about the stream in a dictionary for further use
            start_menu.work_sheet.streams_dict[stream_name] = component_list
            start_menu.work_sheet.streams_list.addItem(QtWidgets.QListWidgetItem(stream_name))

            self.num_streams += 1
            self.close()


class ComponentList(QtWidgets.QWidget):
    name = None

    def __init__(self, name):
        super(ComponentList, self).__init__()
        self.name = name

        self.component_dict = {}
        self.create_component_dict()

    def create_component_dict(self):
        # creating keys for conditions and populate values as "empty"
        self.component_dict["conditions"] = {}
        global conditions_names
        for condition in conditions_names:
            self.component_dict["conditions"][condition] = 'empty'
        # transferring information about components with list (1 list is equal 1 component)
        current_stream_table = start_menu.work_sheet.component_library.current_stream_table
        for row_index in range(0, current_stream_table.rowCount()):
            properties_of_current_row_component = list()
            for col_index in range(current_stream_table.columnCount()):
                properties_of_current_row_component.append(current_stream_table.item(row_index, col_index).text())
            # adding a component to component_dict under the name "component X"
            component_num = f"component {str(row_index + 1)}"
            self.component_dict[component_num] = {}
            global id_name_properties
            for property_index in range(len(id_name_properties)):
                self.component_dict[component_num][id_name_properties[property_index]] = properties_of_current_row_component[property_index]
            self.component_dict[component_num]["Molar Fraction"] = "empty"


class StreamProperties(QtWidgets.QWidget):
    name = None
    T_isDefined = False
    P_isDefined = False
    FlowRate_isDefined = False
    Composition_isDefined = False
    OK = False

    def __init__(self, name="undefined"):
        super(StreamProperties, self).__init__()
        self.name = name

        self.init_UI()

    def init_UI(self):
        self.resize(800, 520)
        self.setWindowTitle(f"Stream Properties of {self.name}")

        self.create_statusbar()
        self.create_tab_component_properties()
        self.create_tab_conditions()
        # creating tabs must be initialized after creating the objects
        self.create_tabs()

    def create_tabs(self):
        self.tabs = QtWidgets.QTabWidget(self)
        self.tabs.setGeometry(0, 0, 800, 500)

        self.tabs.addTab(self.tab_component_properties, "Base properties of components")
        self.tabs.addTab(self.tab_conditions, "Conditions")

    def create_statusbar(self):
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setGeometry(0, 500, 800, 20)

        self.statusbar.showMessage("Composition Unknown")
        self.statusbar.setStyleSheet("QStatusBar{background:rgb(255,0,0);color:black;font-weight:bold;}")

    def create_tab_component_properties(self):
        self.tab_component_properties = QtWidgets.QWidget(self)
        self.tab_component_properties.setGeometry(0, 0, 800, 500)

        self.create_properties_table()
        self.load_data_properties_table()

    def create_properties_table(self):
        self.properties_table = QtWidgets.QTableWidget(self.tab_component_properties)
        self.properties_table.setGeometry(0, 0, 800, 475)

        global id_name_properties
        self.properties_table_vertical_headers_labels = id_name_properties
        num_of_properties = len(self.properties_table_vertical_headers_labels)

        self.properties_table.setRowCount(num_of_properties)
        self.properties_table.setColumnCount(0)
        self.properties_table.setVerticalHeaderLabels(self.properties_table_vertical_headers_labels)
        self.properties_table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        self.properties_table.resizeColumnsToContents()
        self.properties_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

    def load_data_properties_table(self):
        num_of_components = len(start_menu.work_sheet.streams_dict[self.name].component_dict) - 1
        # len - 1, because dictionary has "conditions" + components

        for component_index in range(1, num_of_components + 1):
            num_of_cols = self.properties_table.columnCount() + 1
            self.properties_table.setColumnCount(num_of_cols)

            component_properties = start_menu.work_sheet.streams_dict[self.name].component_dict[f'component {component_index}']
            col_for_add = num_of_cols - 1  # index column begin from 0
            row_index = 0
            for property_name in component_properties.keys():
                item = QtWidgets.QTableWidgetItem(component_properties[property_name])
                item.setTextAlignment(QtCore.Qt.AlignCenter)
                self.properties_table.setItem(row_index, col_for_add, item)
                row_index += 1

        self.properties_table.resizeColumnsToContents()

    def create_tab_conditions(self):
        self.tab_conditions = QtWidgets.QWidget(self)
        self.tab_component_properties.setGeometry(0, 0, 800, 500)

        self.create_conditions_table()
        self.create_composition_table()
        self.create_btn_calculate()
        self.create_btn_cooler()

        self.determine_missing_for_calc()

    def create_conditions_table(self):
        self.conditions_table = QtWidgets.QTableWidget(self.tab_conditions)
        self.conditions_table.setGeometry(0, 0, 400, 415)

        global conditions_names
        self.conditions_table.setRowCount(len(conditions_names + ["SRK Molar Volume [m3/kmol]"]))
        self.conditions_table.setColumnCount(1)

        self.conditions_table.setVerticalHeaderLabels(conditions_names + ["SRK Molar Volume [m3/kmol]"])
        self.conditions_table.setHorizontalHeaderLabels(["Value"])
        self.conditions_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        for condition_index in range(len(conditions_names)):
            item = QtWidgets.QTableWidgetItem(
                start_menu.work_sheet.streams_dict[self.name].component_dict["conditions"][conditions_names[condition_index]])
            item.setTextAlignment(QtCore.Qt.AlignCenter)
            self.conditions_table.setItem(condition_index, 0, item)
        SRK_value_default = QtWidgets.QTableWidgetItem("empty")
        SRK_value_default.setTextAlignment(QtCore.Qt.AlignCenter)
        self.conditions_table.setItem(len(conditions_names), 0, SRK_value_default)

        self.conditions_table.itemChanged.connect(self.change_conditions)

    def change_conditions(self, item):  # method itemChanged has information about changed item and can give it to function
        # each command, that calls item methods, calls itemChanged again, which can cause recursion, so we used block mode
        self.conditions_table.blockSignals(True)

        item.setTextAlignment(QtCore.Qt.AlignCenter)
        condition_name = self.conditions_table.verticalHeaderItem(self.conditions_table.currentRow())
        start_menu.work_sheet.streams_dict[self.name].component_dict["conditions"][condition_name.text()] = item.text()

        match condition_name.text():
            case "Temperature [C]":
                self.T_isDefined = True
            case "Pressure [kPa]":
                self.P_isDefined = True
            case "Flow Rate [kg/sec]":
                self.FlowRate_isDefined = True
        self.determine_missing_for_calc()

        self.conditions_table.blockSignals(False)

    def create_composition_table(self):
        self.composition_table = QtWidgets.QTableWidget(self.tab_conditions)
        self.composition_table.setGeometry(400, 0, 400, 415)

        self.composition_table.setColumnCount(2)
        row_number = len(start_menu.work_sheet.streams_dict[self.name].component_dict)  # the number of rows is greater by 1 than the
        # actual number of components, because the last row for the sum
        self.composition_table.setRowCount(row_number)

        self.composition_table.setHorizontalHeaderLabels(["Component", "Molar Fraction"])
        self.composition_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)

        for index in range(1, row_number):
            component_num = f"component {index}"
            mol_fr_value = QtWidgets.QTableWidgetItem(
                start_menu.work_sheet.streams_dict[self.name].component_dict[component_num]["Molar Fraction"])
            mol_fr_value.setTextAlignment(QtCore.Qt.AlignCenter)
            component_name = QtWidgets.QTableWidgetItem(start_menu.work_sheet.streams_dict[self.name].component_dict[component_num]["Name"])
            row_for_add = index - 1
            self.composition_table.setItem(row_for_add, 0, component_name)
            self.composition_table.setItem(row_for_add, 1, mol_fr_value)
        self.composition_table.setItem(row_number - 1, 0, QtWidgets.QTableWidgetItem("Total"))
        self.calculate_mol_fr_total()

        self.composition_table.itemChanged.connect(self.change_component_mol_fr)

    def change_component_mol_fr(self, item):
        self.composition_table.blockSignals(True)

        item.setTextAlignment(QtCore.Qt.AlignCenter)
        component_index = self.composition_table.currentRow() + 1
        start_menu.work_sheet.streams_dict[self.name].component_dict[f"component {component_index}"]["Molar Fraction"] = item.text()

        self.calculate_mol_fr_total()
        self.determine_missing_for_calc()

        self.composition_table.blockSignals(False)

    def calculate_mol_fr_total(self):
        mol_fr_total = 0
        for row_index in range(0, self.composition_table.rowCount() - 1):  # last row should not be calculated (it is "Total")
            if (self.composition_table.item(row_index, 1).text()) != "empty":
                mol_fr_total += float(self.composition_table.item(row_index, 1).text())
        mol_fr_total = round(mol_fr_total, 2)
        if 0.9999 <= mol_fr_total <= 1.0001:
            self.Composition_isDefined = True
        else:
            self.Composition_isDefined = False

        if mol_fr_total == 0:
            mol_fr_total = "empty"
        else:
            mol_fr_total = str(mol_fr_total)
        mol_fr_total = QtWidgets.QTableWidgetItem(mol_fr_total)
        mol_fr_total.setTextAlignment(QtCore.Qt.AlignCenter)
        self.composition_table.setItem(self.composition_table.rowCount() - 1, 1, mol_fr_total)

    def create_btn_calculate(self):
        self.btn_calculate = QtWidgets.QPushButton(self.tab_conditions)
        self.btn_calculate.setGeometry(0, 415, 195, 60)
        self.btn_calculate.setText("Calculate" + "\n" + "SRK Molar Volume!")

        self.btn_calculate.clicked.connect(self.calculate_SRK)

    def calculate_SRK(self):
        if self.OK:
            a_list = []
            b_list = []
            mol_fr_list = []

            for index in range(1, len(start_menu.work_sheet.streams_dict[self.name].component_dict)):
                mol_fr_comp = float(start_menu.work_sheet.streams_dict[self.name].component_dict[f"component {index}"]["Molar Fraction"])
                mol_fr_list.append(mol_fr_comp)

            T = float(start_menu.work_sheet.streams_dict[self.name].component_dict["conditions"]["Temperature [C]"]) + 273.15
            # added 273.15 because the temperature is in [C], and in the equation the temperature is in [K]
            P = float(start_menu.work_sheet.streams_dict[self.name].component_dict["conditions"]["Pressure [kPa]"]) * 1000
            # multiplied by 1000, because the pressure is in kPa, and in the equation the pressure is in Pa

            R = 8.31441
            for index in range(1, len(start_menu.work_sheet.streams_dict[self.name].component_dict)):
                component_num = f"component {index}"
                T_c = float(
                    start_menu.work_sheet.streams_dict[self.name].component_dict[component_num]["Critical Temperature [C]"]) + 273.15
                P_c = float(start_menu.work_sheet.streams_dict[self.name].component_dict[component_num]["Critical Pressure [kPa]"]) * 1000
                w = float(start_menu.work_sheet.streams_dict[self.name].component_dict[component_num]["Acentricity"])

                m = 0.480 + 1.574 * w - 0.176 * (w ** 2)
                alpha = (1 + m * (1 - np.sqrt(T / T_c))) ** 2

                a = 0.42748 * (((R ** 2) * (T_c ** 2)) / P_c) * alpha
                b = 0.08664 * ((R * T_c) / P_c)

                a_list.append(a)
                b_list.append(b)

            a_mix = 0
            for index in range(len(a_list)):
                a_mix += mol_fr_list[index] * np.sqrt(a_list[index])
            a_mix = a_mix ** 2

            b_mix = 0
            for index in range(len(b_list)):
                b_mix += mol_fr_list[index] * b_list[index]

            def SRV_function(v):
                func = (v ** 3) * P - (v ** 2) * R * T + v * (a_mix - P * (b_mix ** 2) - R * T * b_mix) - a_mix * b_mix
                return func

            molar_volume = fsolve(SRV_function, [10000])
            molar_volume = QtWidgets.QTableWidgetItem(str(molar_volume[0] * 1000))  # multiply by 1000 to convert [m3/mol] to [m3/kmol]

            self.conditions_table.setItem(len(conditions_names), 0, molar_volume)

    def create_btn_cooler(self):
        self.btn_cooler = QtWidgets.QPushButton(self.tab_conditions)
        self.btn_cooler.setGeometry(195, 415, 195, 60)
        self.btn_cooler.setText("Cooler")

        self.btn_cooler.clicked.connect(self.calculate_cooler)

    def calculate_cooler(self):
        # ДАННЫЕ ДЛЯ ТРУБ
        n = 110  # количество трубок
        l = 12  # длина, м
        d_out = 0.025  # наружный диаметр трубок, мм
        d_in = 0.02  # внутренний диаметр трубок, мм
        v_pipe = n * np.pi / 4 * l * (d_out ** 2 - d_in ** 2)  # объем трубок
        rho_pipe = 7850  # плотность материала труб, кг/м3
        m_pipe = v_pipe * rho_pipe  # масса трубок, кг
        c_pipe = 490  # теплоемкость материала труб Дж/кг/К
        lam_p = 48  # коэффициент теплопроводности материала труб, Вт/м/К

        # ДАННЫЕ ДЛЯ НК-70 (межтрубное пр-во)
        rho_1 = 720  # плотность НК-70, кг/м3
        v_1 = 0.79 * 2  # объем межтрубного пространства, м3
        m_1 = rho_1 * v_1  # масса НК-70 в аппарате, кг
        c_1 = 2662  # теплоемкость НК-70, Дж/кг/К
        l_1 = 2 * (l + 21 * 0.5)  # длина пути межтрубного пространства, м
        S_1 = np.pi * d_out * l * n  # площадь пов-ти теплообмена со стороны НК-70, м2
        Vf_1 = 27.57 / 3600  # объемный расход НК-70, м3/с
        Mf_1 = Vf_1 * rho_1  # массовый расход НК-70, кг/с
        a1 = 500  # коэффициент теплоотдачи НК-70 - трубка, Вт/м2/К
        T1_1 = 47.58 + 273  # температура НК-70 на входе в аппарат, К
        T2_1 = 38.46 + 273  # температура НК-70 на выходе из аппарата, К
        lam_1 = 0.14  # коэффициент теплопроводности НК-70, Вт/м/К
        nu1 = 6 * 10 ** -7 * rho_1

        # ДАННЫЕ ДЛЯ ВОДЫ (трубное пр-во)
        rho_2 = 994  # плотность воды, кг/м3
        v_2 = 0.2 * 2  # объем трубного пространства, м3
        m_2 = rho_2 * v_2  # масса воды в аппарате, кг
        c_2 = 4180  # теплоемкость воды, Дж/кг/К
        S_2 = np.pi * d_in * l * n  # площадь пов-ти теплообмена со стороны воды, м2
        a2 = 2000  # коэффициент теплоотдачи трубка - вода, Вт/м2/К
        T1_2 = 20 + 273  # температура воды на входе в аппарат, К
        T2_2 = 25.33 + 273  # температура воды на выходе из аппарата, К
        lam_2 = 0.6  # коэффициент теплопроводности воды, Вт/м/К
        Mf_2 = Mf_1 * c_1 * (T1_1 - T2_1) / (c_2 * (T2_2 - T1_2))  # массовый расход воды, кг/с
        nu2 = 1000 * 10 ** (-6)

        # ТЕМПЕРАТУРНЫЙ НАПОР
        diff1 = (T1_1 - T1_2)
        diff2 = (T2_1 - T2_2)
        delta_T = (diff1 - diff2) / np.log(diff1 / diff2)  # средняя разница температур

        Re1 = d_out * Mf_1 / (0.785 * 0.5 ** 2 - n * 0.785 * d_out ** 2) / nu1 * 4
        Pr1 = c_1 * nu1 / lam_1
        a1 = 0.24 * (Re1 ** 0.6) * (Pr1 ** 0.36) * lam_1 / d_out

        Re2 = 4 * Mf_2 / np.pi / d_in / 28 / nu2
        Pr2 = c_2 * nu2 / lam_2
        a2 = 0.021 * (Re2 ** 0.8) * (Pr2 ** 0.43) * lam_2 / d_in
        print('коэффициенты теплоотдачи:', round(a1, 2), round(a2, 2))

        def h_e(t, T):
            dTdt = np.zeros(3)
            dTdt[0] = (-(a1 * S_1 + Mf_1 * c_1) / m_1 / c_1) * T[0] + a1 * S_1 * T[1] / m_1 / c_1 + T1_1 * Mf_1 / m_1
            dTdt[1] = a1 * S_1 * T[0] / m_pipe / c_pipe - ((a1 * S_1 + a2 * S_2) / m_pipe / c_pipe) * T[1] + a2 * S_2 * T[
                2] / m_pipe / c_pipe
            dTdt[2] = a2 * S_2 * T[1] / m_2 / c_2 - ((a2 * S_2 + Mf_2 * c_2) / m_2 / c_2) * T[2] + T1_2 * Mf_2 / m_2
            return dTdt

        time = np.linspace(0, 1000, 100)

        solution = spi.solve_ivp(h_e, [0, 1000], [T1_1, 273 + 25, T1_2], method='Radau', t_eval=time)
        T1 = solution.y[0]
        T2 = solution.y[1]
        T3 = solution.y[2]

        print('Температура НК-70 на выходе из аппарата', round(T1[-1], 2) - 273.15, 'K')

        fig, (ax1, ax2) = plt.subplots(2, 1)
        ax1.plot(time, T1, label='Температура НК-70')
        ax1.plot(time, T2, label='Температура трубок')
        ax1.plot(time, T3, label='Температура воды')
        ax1.legend()
        ax1.set_xlabel('Время, c', )
        ax1.set_ylabel('Температура, K')
        # ax1.set_xticks(fontsize=12)
        # ax1.set_yticks(fontsize=12)
        ax1.set_xlim(0, )
        ax1.set_ylim(250, 350)
        ax1.grid()

        for Mf_2 in range(5, 30, 5):
            def h_e(t, T):
                dTdt = np.zeros(3)
                dTdt[0] = (-(a1 * S_1 + Mf_1 * c_1) / m_1 / c_1) * T[0] + a1 * S_1 * T[1] / m_1 / c_1 + T1_1 * Mf_1 / m_1
                dTdt[1] = a1 * S_1 * T[0] / m_pipe / c_pipe - ((a1 * S_1 + a2 * S_2) / m_pipe / c_pipe) * T[1] + a2 * S_2 * T[
                    2] / m_pipe / c_pipe
                dTdt[2] = a2 * S_2 * T[1] / m_2 / c_2 - ((a2 * S_2 + Mf_2 * c_2) / m_2 / c_2) * T[2] + T1_2 * Mf_2 / m_2
                return dTdt

            time = np.linspace(0, 1000, 100)

            solution = spi.solve_ivp(h_e, [0, 1000], [T1_1, 273 + 25, T1_2], method='LSODA', t_eval=time)
            T1 = solution.y[0]
            T2 = solution.y[1]
            T3 = solution.y[2]
            ax2.plot(time, T1, label=Mf_2)
        # plt.plot(time, T1, label = 'Температура НК-70')
        # plt.plot(time, T2, label = 'Температура трубок')
        # plt.plot(time, T3, label = 'Температура воды')
        ax2.legend()
        ax2.set_xlabel('Время, c', )
        ax2.set_ylabel('Температура, K')
        # ax2.set_xticks(fontsize=12)
        # ax2.set_yticks(fontsize=12)
        ax2.set_xlim(0, )
        ax2.set_ylim(250, 350)
        ax2.grid(label='расход воды')
        ax2.legend()

        plt.show()


    def determine_missing_for_calc(self):
        conditions_areDefined = [self.Composition_isDefined, self.T_isDefined, self.P_isDefined, self.FlowRate_isDefined]
        match conditions_areDefined:
            case [True, True, True, True]:
                self.OK = True
                self.statusbar.showMessage("OK")
                self.statusbar.setStyleSheet("QStatusBar{background:rgb(0,255,0);color:black;font-weight:bold;}")
            case [True, False, False, False] | [True, False, True, False] | [True, False, True, True]:
                self.statusbar.showMessage("Temperature Unknown")
                self.statusbar.setStyleSheet("QStatusBar{background:rgb(255,0,0);color:black;font-weight:bold;}")
            case [True, True, False, False] | [True, True, False, True]:
                self.statusbar.showMessage("Pressure Unknown")
                self.statusbar.setStyleSheet("QStatusBar{background:rgb(255,0,0);color:black;font-weight:bold;}")
            case [True, True, True, False]:
                self.statusbar.showMessage("Flow Rate Unknown")
                self.statusbar.setStyleSheet("QStatusBar{background:rgb(255,0,0);color:black;font-weight:bold;}")
            case _:
                self.statusbar.showMessage("Composition Unknown")
                self.statusbar.setStyleSheet("QStatusBar{background:rgb(255,0,0);color:black;font-weight:bold;}")


app = QtWidgets.QApplication(sys.argv)
start_menu = StartMenu()
start_menu.show()

sys.exit(app.exec_())

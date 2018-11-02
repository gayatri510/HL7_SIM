import logging
from PyQt4 import QtGui, QtCore
from hl7menu import Hl7_menu
from copy import deepcopy
from collections import OrderedDict

logging.basicConfig(level=logging.CRITICAL,
                    format='[%(threadName)s] %(message)s',
                    )

COMPANY_NAME = "CapsuleTech"


class Configuration_Window(QtGui.QDialog):

    def __init__(self, settings_to_load, parent=None):
        super(Configuration_Window, self).__init__()

        self.cfg_window_settings = QtCore.QSettings(COMPANY_NAME, "Configuration_Window")

        # Obtain the settings to load from the Main Window class and load the default settings
        self.default_hl7_segments = settings_to_load.value('HL7_segments', type=str)
        self.default_calculated_variables = settings_to_load.value('Calculated Variables', type=str)
        default_hl7_segment_setting_values = settings_to_load.value('Configurationbox_segments').toPyObject()
        self.default_hl7_segment_setting_values_dict = OrderedDict()
        for key, value in default_hl7_segment_setting_values.items():
            self.default_hl7_segment_setting_values_dict[str(key)] = value
        default_header_variable_values = settings_to_load.value('Header Variables').toPyObject()
        self.default_header_variable_values_dict = OrderedDict()
        for key, value in default_header_variable_values.items():
            self.default_header_variable_values_dict[key] = str(value)
        default_obr7_timestamp_state = settings_to_load.value('OBR 7 timestamp').toBool()
        default_obx14_timestamp_state = settings_to_load.value('OBX 14 timestamp').toBool()
        # Create actions and widgets
        msg_label = QtGui.QLabel("HL7 Segments")
        self.msg_linedit = QtGui.QLineEdit()
        self.msg_linedit.setText(self.default_hl7_segments)
        msg_label.setBuddy(self.msg_linedit)
        self.table = QtGui.QTableWidget()
        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)
        self.table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        # This will populate the table with the contents in the xcolumns list
        # and the number of columns are always set to 2
        self.table.setRowCount(len(self.default_hl7_segment_setting_values_dict))
        self.table.setColumnCount(2)
        index = 0
        for key, value in self.default_hl7_segment_setting_values_dict.iteritems():
            message_label = QtGui.QLabel(key)
            message_box_information = QtGui.QComboBox()
            message_box_information.addItems([str(i) for i in xrange(1, 31)])
            message_box_information.setCurrentIndex(int(value) - 1)

            self.table.setCellWidget(index, 0, message_label)
            self.table.setCellWidget(index, 1, message_box_information)
            index = index + 1

        self.table.horizontalHeader().setResizeMode(0, QtGui.QHeaderView.Stretch)
        self.table.horizontalHeader().setResizeMode(1, QtGui.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)

        self.tableGroupBox = QtGui.QGroupBox("HL7 Segment Default Settings")
        tableLayout = QtGui.QVBoxLayout()
        tableLayout.addWidget(self.table)
        self.tableGroupBox.setLayout(tableLayout)
        # Header Variables table
        self.headervariablesGroupBox = QtGui.QGroupBox("Header Variables")
        self.headervariablesTable = QtGui.QTableWidget()
        self.headervariablesTable.horizontalHeader().setVisible(False)
        self.headervariablesTable.verticalHeader().setVisible(False)
        self.headervariablesTable.setColumnCount(2)
        self.headervariablesTable.setRowCount(len(self.default_header_variable_values_dict) + 10)
        for row_item, key in enumerate(self.default_header_variable_values_dict.keys()):
            self.headervariablesTable.setItem(row_item, 0, QtGui.QTableWidgetItem(key))
            self.headervariablesTable.setItem(row_item, 1, QtGui.QTableWidgetItem(self.default_header_variable_values_dict[key]))
        self.headervariablesTable.setHorizontalHeaderLabels(("Capsule Variable ID", "Header Description"))
        self.headervariablesTable.horizontalHeader().setResizeMode(0, QtGui.QHeaderView.Stretch)
        self.headervariablesTable.horizontalHeader().setResizeMode(1, QtGui.QHeaderView.Stretch)
        self.headervariablesTable.horizontalHeader().resizeSection(1, 180)
        headervariablesLayout = QtGui.QVBoxLayout()
        headervariablesLayout.addWidget(self.headervariablesTable)
        self.headervariablesGroupBox.setLayout(headervariablesLayout)
        # Calculated Variables
        calculated_variables_text = QtGui.QLabel("Calculated Variables")
        self.calculated_variables = QtGui.QLineEdit()
        self.calculated_variables.setText(self.default_calculated_variables)
        calculated_variables_text.setBuddy(self.calculated_variables)
        # OBR and OBX Timestamps
        self.obr7_timestamp_checkbox = QtGui.QCheckBox("Generate Timestamp In OBR-7 field", self)
        self.obr7_timestamp_checkbox.setChecked(default_obr7_timestamp_state)
        self.cfg_window_settings.setValue('OBR 7 timestamp', default_obr7_timestamp_state)
        self.obr7_timestamp_checkbox.stateChanged.connect(self.obr7_timestamp_statechange)

        self.obx14_timestamp_checkbox = QtGui.QCheckBox("Generate Timestamp In OBX-14 field", self)
        self.obx14_timestamp_checkbox.setChecked(default_obx14_timestamp_state)
        self.cfg_window_settings.setValue('OBX 14 timestamp', default_obx14_timestamp_state)
        self.obx14_timestamp_checkbox.stateChanged.connect(self.obx14_timestamp_statechange)

        self.buttonBox = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok | QtGui.QDialogButtonBox.Cancel)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        mainLayout = QtGui.QGridLayout()
        mainLayout.addWidget(msg_label, 0, 0)
        mainLayout.addWidget(self.msg_linedit, 0, 1)
        mainLayout.addWidget(self.tableGroupBox, 1, 0, 1, 2)
        mainLayout.addWidget(self.headervariablesGroupBox, 2, 0, 1, 2)
        mainLayout.addWidget(calculated_variables_text, 3, 0)
        mainLayout.addWidget(self.calculated_variables, 3, 1)
        mainLayout.addWidget(self.obr7_timestamp_checkbox, 4, 0, 1, 2)
        mainLayout.addWidget(self.obx14_timestamp_checkbox, 5, 0, 1, 2)
        mainLayout.addWidget(self.buttonBox, 6, 0, 1, 2)
        self.setLayout(mainLayout)
        self.setWindowTitle("Configuration Settings")
        self.resize(900, 600)

    def obr7_timestamp_statechange(self):
        if self.obr7_timestamp_checkbox.isChecked() is True:
            self.obr7_timestamp_checkbox.setChecked(True)
            self.cfg_window_settings.setValue('OBR 7 timestamp', True)
        else:
            self.obr7_timestamp_checkbox.setChecked(False)
            self.cfg_window_settings.setValue('OBR 7 timestamp', False)

    def obx14_timestamp_statechange(self):
        if self.obx14_timestamp_checkbox.isChecked() is True:
            self.obx14_timestamp_checkbox.setChecked(True)
            self.cfg_window_settings.setValue('OBX 14 timestamp', True)
        else:
            self.obx14_timestamp_checkbox.setChecked(False)
            self.cfg_window_settings.setValue('OBX 14 timestamp', False)

    def get_current_settings(self):
        ''' Compares the original information with the applied changes if any and ...'''
        self.cfg_window_settings = QtCore.QSettings(COMPANY_NAME, "Configuration_Window")
        # Grab HL7 segment changes and compare to see if anything changed
        current_hl7_segments = self.msg_linedit.text()
        self.cfg_window_settings.setValue('HL7_segments', current_hl7_segments)

        # Grab calculated variables list and compare to see if anything changed
        current_calculated_variables = self.calculated_variables.text()
        self.cfg_window_settings.setValue('Calculated Variables', current_calculated_variables)
        # Grab the HL7 configurationbox segment dictionary values and compare to see if anything changed
        current_configurationbox_segments = deepcopy(self.default_hl7_segment_setting_values_dict)

        for row in range(0, self.table.rowCount()):
            column0_text = str(self.table.cellWidget(row, 0).text())
            column1_text = str(self.table.cellWidget(row, 1).currentText())

        # Modify the global configurationbox segment dictionary with the new values
            if column0_text in current_configurationbox_segments.keys():
                current_configurationbox_segments[column0_text] = column1_text

        self.cfg_window_settings.setValue('Configurationbox_segments', current_configurationbox_segments)

        # Grab the Header variables dictionary from the header variables list
        current_header_variable_dict = OrderedDict()
        for hd_row in range(0, self.headervariablesTable.rowCount()):
            try:
                hd_key = self.headervariablesTable.item(hd_row, 0).text()
                hd_value = self.headervariablesTable.item(hd_row, 1).text()
                current_header_variable_dict[hd_key] = hd_value
            except AttributeError:       # If the table contents in that row are empty, then ignore
                pass
        self.cfg_window_settings.setValue('Header Variables', current_header_variable_dict)

        return self.cfg_window_settings

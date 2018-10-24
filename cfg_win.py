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

        self.cfg_window_settings = None


        #Obtain the settings to load from the Main Window class and load the default settings
        self.default_hl7_segments = settings_to_load.value('HL7_segments', type = str)
   

        default_hl7_segment_setting_values = settings_to_load.value('Configurationbox_segments').toPyObject()
        self.default_hl7_segment_setting_values_dict = OrderedDict()
        for key, value in default_hl7_segment_setting_values.items():
            self.default_hl7_segment_setting_values_dict[str(key)] = value
 


        msg_label = QtGui.QLabel("HL7 Segments")
        self.msg_linedit = QtGui.QLineEdit()
        self.msg_linedit.setText(self.default_hl7_segments)
        msg_label.setBuddy(self.msg_linedit)

        

        self.table = QtGui.QTableWidget()
        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setResizeMode(0, QtGui.QHeaderView.Stretch)
        self.table.horizontalHeader().setResizeMode(1, QtGui.QHeaderView.Stretch)
        self.table.horizontalHeader().resizeSection(1, 180)            
        self.table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        
        # This will populate the table with the contents in the xcolumns list and the number of columns are always set to 2 
        self.table.setRowCount(len(self.default_hl7_segment_setting_values_dict))
        self.table.setColumnCount(2)


        index = 0
        for key, value in self.default_hl7_segment_setting_values_dict.iteritems():
            message_label    = QtGui.QLabel(key)
            message_box_information = QtGui.QComboBox()
            message_box_information.addItems([str(i) for i in xrange(1,31)])
            message_box_information.setCurrentIndex(int(value) - 1)

            self.table.setCellWidget(index, 0, message_label)
            self.table.setCellWidget(index, 1, message_box_information)
            index = index + 1

        self.tableGroupBox = QtGui.QGroupBox("HL7 Segment Default Settings")
        tableLayout = QtGui.QVBoxLayout()
        tableLayout.addWidget(self.table)
        self.tableGroupBox.setLayout(tableLayout)



        self.obr7_timestamp_checkbox = QtGui.QCheckBox("Generate Timestamp In OBR-7 field",self)
        self.obr_timestamp_state = "UnChecked"
        # #self.obr7_timestamp_checkbox.stateChanged.connect(self.click_obr_timestamp_checkBox)

        self.obx14_timestamp_checkbox = QtGui.QCheckBox("Generate Timestamp In OBX-14 field",self)
        self.obx_timestamp_state = "UnChecked"
        # #self.obx14_timestamp_checkbox.stateChanged.connect(self.click_obx_timestamp_checkBox)

        self.buttonBox = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok | QtGui.QDialogButtonBox.Cancel)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
 
        mainLayout = QtGui.QGridLayout()
        mainLayout.addWidget(msg_label, 0, 0)
        mainLayout.addWidget(self.msg_linedit, 0, 1)
        mainLayout.addWidget(self.tableGroupBox, 1, 0, 1, 2)
        mainLayout.addWidget(self.obr7_timestamp_checkbox, 2, 0, 1, 2)
        mainLayout.addWidget(self.obx14_timestamp_checkbox, 3, 0, 1, 2)
        mainLayout.addWidget(self.buttonBox, 4, 0, 1, 2)
        self.setLayout(mainLayout)
 
 
        self.setWindowTitle("Configuration Settings")
        self.resize(650, 400)


    #     self.obr7_timestamp_checkbox = QtGui.QCheckBox("Generate Timestamp In OBR-7 field",self)
    #     self.obr_timestamp_state = "UnChecked"
    #     self.obr7_timestamp_checkbox.stateChanged.connect(self.click_obr_timestamp_checkBox)

    #     self.obx14_timestamp_checkbox = QtGui.QCheckBox("Generate Timestamp In OBX-14 field",self)
    #     self.obx_timestamp_state = "UnChecked"
    #     self.obx14_timestamp_checkbox.stateChanged.connect(self.click_obx_timestamp_checkBox)

    #     '''
    #     action pane
    #     '''  
  
 

    def compare_hl7_segments(self, original, current):
        ''' Compares the original and current HL7 segments'''
        original_hl7_segment_list = original.split(',')
        current_hl7_segment_list = map(str, current.split(','))
        return bool(set(current_hl7_segment_list).difference(original_hl7_segment_list))


    def compare_hl7_configurationbox_segment_values(self, original_dict, current_dict):
        '''Compares the original and current configuration box segment values'''
        return bool(set(current_dict.items()).difference(original_dict.items()))



    def get_current_settings(self):
        ''' Compares the original information with the applied changes if any and ...'''
        self.cfg_window_settings = QtCore.QSettings(COMPANY_NAME, "Configuration_Window") 

        # Grab HL7 segment changes and compare to see if anything changed
        current_hl7_segments = self.msg_linedit.text()
        self.cfg_window_settings.setValue('HL7_segments', current_hl7_segments)


        # Grab the HL7 configurationbox segment dictionary values and compare to see if anything changed
        current_configurationbox_segments = OrderedDict(self.default_hl7_segment_setting_values_dict)

        for row in range(0,self.table.rowCount()):
            column0_text = str(self.table.cellWidget(row,0).text())
            column1_text = str(self.table.cellWidget(row,1).currentText())

        # Modify the global configurationbox segment dictionary with the new values
            if column0_text in current_configurationbox_segments.keys():
                current_configurationbox_segments[column0_text] = column1_text

        self.cfg_window_settings.setValue('Configurationbox_segments', current_configurationbox_segments)


        return self.cfg_window_settings

        # # Grab the checkbox state for the OBR-7 timestamp checkbox
        # if self.obr_timestamp_state == "Checked":
        #     add_obr_timestamp = True


        # # Grab the checkbox state for the OBR-7 timestamp checkbox
        # if self.obx_timestamp_state == "Checked":
        #     add_obx_timestamp = True


        # return (update_hl7_segment_boxes, update_hl7_menu, add_obr_timestamp, add_obx_timestamp)









    #     # # Will have to modify the dropdown hl7 menu list once the configurationbox_segments dictionary is updated
    #     # hl7_dropdown_menu_items = Hl7_menu()
    #     # hl7_dropdown_menu_items.create_dropdown_items_from_dict(current_configurationbox_segments)

        #self.close()

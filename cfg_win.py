import logging
from PyQt4 import QtGui, QtCore
from hl7menu import Hl7_menu
from config import configurationbox_segments, default_hl7_segments, obr_7_timestamp, obx_14_timestamp
from copy import deepcopy
from collections import OrderedDict

logging.basicConfig(level=logging.CRITICAL,
                    format='[%(threadName)s] %(message)s',
)


class Configuration_Window(QtGui.QDialog):

    def __init__(self, parent=None):
        super(Configuration_Window, self).__init__()

        self.initWindow()
        self.initPage()                   
        self.runWindow()

        # adjust dialogs height and width based
        # on the items added inside
        self.height = 0
        self.width  = 0        

        
    def initWindow(self):    
        QtGui.QToolTip.setFont(QtGui.QFont('SansSerif', 10))
        self.wid = QtGui.QWidget(self)


    def initPage(self):
        logging.debug("Configuration Page")
        vbox = QtGui.QVBoxLayout()

        '''
        Populating the default HL7 segments widget'''

        self.hbox = QtGui.QHBoxLayout()

        msg_label = QtGui.QLabel("HL7 Segments")
        msg_linedit = QtGui.QLineEdit()
        msg_linedit.setText(default_hl7_segments)

        self.hbox.addWidget(msg_label)
        self.hbox.addWidget(msg_linedit) 

        ''' Populating table for HL7 segments default values'''
          
        self.popTable()   

        ''' Populating multiple checkboxes for timestamp informations''' 

        self.obr7_timestamp_checkbox = QtGui.QCheckBox("Generate Timestamp In OBR-7 field",self)
        self.obr_timestamp_state = "UnChecked"
        self.obr7_timestamp_checkbox.stateChanged.connect(self.click_obr_timestamp_checkBox)

        self.obx14_timestamp_checkbox = QtGui.QCheckBox("Generate Timestamp In OBX-14 field",self)
        self.obx_timestamp_state = "UnChecked"
        self.obx14_timestamp_checkbox.stateChanged.connect(self.click_obx_timestamp_checkBox)

        '''
        action pane
        '''  
        # This button will apply any configuration changes applied by the user     
        btn_apply = QtGui.QPushButton('Apply', self)
        btn_apply.setToolTip('Applies the Configuration set by the user')
        btn_plapy.clicked.connect(self.apply_configuration)
       
        apane = QtGui.QHBoxLayout()
        apane.addStretch(1)
        apane.addWidget(btn_apply)
        # apane.addWidget(btn_exit)

        vbox.addLayout(self.hbox)
        vbox.addWidget(self.table)
        vbox.addWidget(self.obr7_timestamp_checkbox) 
        vbox.addWidget(self.obx14_timestamp_checkbox)  
        vbox.addLayout(apane) 
        vbox.addLayout(apane)

        self.wid.setLayout(vbox)

    def popTable(self):
        self.table = QtGui.QTableWidget()
        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)            
        self.table.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        
        # This will populate the table with the contents in the xcolumns list and the number of columns are always set to 2 
        self.table.setRowCount(len(configurationbox_segments))
        self.table.setColumnCount(2)

        items_list = [str(i) for i in xrange(1,31)]

        # This will generate the table contents that will next be used to fill the table
        table_contents = []


        for column in configurationbox_segments.keys():
            message_label    = QtGui.QLabel(column)
            message_box_information = QtGui.QComboBox()
            message_box_information.addItems(items_list)
            self.set_default_configuration_value(message_box_information, message_label)

            table_contents.append((message_label,message_box_information))

             
        # This populates the table with table_contents's contents in the given row and cell.
        # Each row will have two columns with the first column populated with column_1 which is the label and column_2 which is the Dropdown 
        index = 0
        for column_1, column_2 in table_contents:
            self.table.setCellWidget(index, 0, column_1)
            self.table.setCellWidget(index, 1, column_2)
            index = index + 1

        self.table.resizeColumnsToContents()
        self.table.setFixedWidth(self.table.columnWidth(0)+self.table.columnWidth(1)+25)

        self.height = self.table.height()
        # i dont know why i have to do this, how is the table's vertical scroll limit set
        self.height = self.height - 50
        
        # length() includes the width of all its sections + scroll bar width
        self.width = self.table.horizontalHeader().length()
        self.width += self.table.verticalHeader().width()
        self.width += self.table.verticalScrollBar().width()
        
           
    def runWindow(self):
        self.resize(self.width,self.height)
        self.setWindowTitle('Configuration')

    def click_obr_timestamp_checkBox(self, state):
        if state == QtCore.Qt.Checked:
            self.obr_timestamp_state = "Checked"
        elif state == QtCore.Qt.UnChecked:
            self.obr_timestamp_state = "UnChecked"


    def click_obx_timestamp_checkBox(self, state):
        if state == QtCore.Qt.Checked:
            self.obx_timestamp_state = "Checked"
        elif state == QtCore.Qt.UnChecked:
            self.obx_timestamp_state = "UnChecked"

    def set_default_configuration_value(self, message_box_information, message_label):
        ''' This method will set the default values for the different segments in the dropdown box'''
    
        if message_label.text() in configurationbox_segments.keys():
            default_value = configurationbox_segments[str(message_label.text())]
            index = message_box_information.findText(default_value)
            message_box_information.setCurrentIndex(index)

    def make_hl7Menu(self,dropdown,menu,hl7_dictionary):      
        for key in hl7_dictionary:
            sub_menu = menu.addMenu(key)
            action = sub_menu.addAction(key)
            action.triggered.connect(self.updateTable(dropdown,key))
            if isinstance(hl7_dictionary[key], dict):
                self.make_hl7Menu(dropdown,sub_menu,hl7_dictionary[key])
            else:
                for dict_value in hl7_dictionary[key]:
                    action = sub_menu.addAction(dict_value)
                    action.triggered.connect(self.updateTable(dropdown,dict_value))


    def updateTable(self,dropdown,dict_value):
        return lambda : dropdown.setText(dict_value)


    def compare_hl7_segments(self, original, current):
        ''' Compares the original and current HL7 segments'''
        original_hl7_segment_list = original.split(',')
        current_hl7_segment_list = map(str, current.split(','))
        return bool(set(current_hl7_segment_list).difference(original_hl7_segment_list))


    def compare_hl7_configurationbox_segment_values(self, original_dict, current_dict):
        '''Compares the original and current configuration box segment values'''
        return bool(set(current_dict.items()).difference(original_dict.items()))



    def apply_configuration(self):
        ''' Compares the original information with the applied changes if any and ...'''
        global configurationbox_segments

        update_hl7_segment_boxes = False
        update_hl7_menu = False
        add_obr_timestamp = False
        add_obx_timestamp = False

        # Grab HL7 segment changes and compare to see if anything changed
        current_hl7_segments = self.hbox.itemAt(1).widget().text()
        if self.compare_hl7_segments(default_hl7_segments, current_hl7_segments):
            update_hl7_segment_boxes =  True




        # Grab the HL7 configurationbox segment dictionary values and compare to see if anything changed
        original_configurationbox_segments = deepcopy(configurationbox_segments)
        current_configurationbox_segments = OrderedDict(configurationbox_segments)

        for row in range(0,self.table.rowCount()):
            column0_text = str(self.table.cellWidget(row,0).text())
            column1_text = str(self.table.cellWidget(row,1).currentText())

        # Modify the global configurationbox segment dictionary with the new values
            if column0_text in current_configurationbox_segments.keys():
                current_configurationbox_segments[column0_text] = column1_text

        if self.compare_hl7_configurationbox_segment_values(original_configurationbox_segments, current_configurationbox_segments):
            update_hl7_menu =  True


        # Grab the checkbox state for the OBR-7 timestamp checkbox
        if self.obr_timestamp_state == "Checked":
            add_obr_timestamp = True


        # Grab the checkbox state for the OBR-7 timestamp checkbox
        if self.obx_timestamp_state == "Checked":
            add_obx_timestamp = True


        return (update_hl7_segment_boxes, update_hl7_menu, add_obr_timestamp, add_obx_timestamp)









        # # Will have to modify the dropdown hl7 menu list once the configurationbox_segments dictionary is updated
        # hl7_dropdown_menu_items = Hl7_menu()
        # hl7_dropdown_menu_items.create_dropdown_items_from_dict(current_configurationbox_segments)

        self.close()

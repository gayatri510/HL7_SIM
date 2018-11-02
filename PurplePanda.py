#!/usr/bin/python2.7
# encoding=utf8
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import os
import logging
from PyQt4 import QtGui, QtCore
from PyQt4.QtCore import pyqtSlot
import pandas as pd
import xlrd
import img_qr
from cfg_win import Configuration_Window
import re
from copy import deepcopy
from hl7menu import Hl7_menu
from collections import OrderedDict
from create_hl7_data import Create_Hl7_Data
from itertools import cycle
import traceback

logging.basicConfig(level=logging.DEBUG,
                    format='[%(threadName)s] %(message)s',
)


COMPANY_NAME = "CapsuleTech"
APPLICATION_NAME = "MainWindow"
QE_INPUT_SHEETS_TO_DISCARD = ["Template Cover Sheet", "Summary", "General", "Report", "Comparison_Report"]
DEFAULT_HL7_SEGMENTS = "MSH, PID, PV1, OBR"
DEFAULT_CALCULATED_VARIABLES = "1190, 2583, 5815, 5816"
DEFAULT_CONFIGURATIONBOX_SEGMENTS_DATA = OrderedDict([('MSH Field Count', '21'),
                                                      ('MSH Component Count', '10'),
                                                      ('MSH Subcomponent Count', '5'),
                                                      ('PID Field Count', '17'),
                                                      ('PID Component Count', '10'),
                                                      ('PID Subcomponent Count', '5'),                                      
                                                      ('PV1 Field Count', '17'),
                                                      ('PV1 Component Count', '10'),
                                                      ('PV1 Subcomponent Count', '5'),                                        
                                                      ('OBR Field Count', '17'),
                                                      ('OBR Component Count', '10'),
                                                      ('OBR Subcomponent Count', '5'),
                                                      ('OBX Field Count', '17'),
                                                      ('OBX Component Count', '10'),
                                                      ('OBX Subcomponent Count', '5')
])
HEADER_VARIABLES = OrderedDict([('170',  'PID-8'),
                                ('1220', 'PID-7'),
                                ('1911', 'MSH-3-2'),
                                ('1929', 'PID-3-1'),                                        
                                ('1930', 'PID-5'),
                                ('1931', 'PV1-3'),
                                ('2255', 'MSH-7'),
                                ('2901', 'PID-5-2'),
                                ('6510', 'MSH-3'),                                                              
                                ('6511', 'MSH-4'),
                                ('6512', 'MSH-5'),
                                ('6513', 'MSH-6'), 
                                ('6521', 'PV1-2'),
                                ('7914', 'MSH-3-1'),
                                ('8102', 'OBR-3'),
                                ('8036', 'PID-18'),     
                                ('8340', 'PID-5-1'),
                                ('9593', 'OBR-2'),
                                ('9569', 'PID-3'), 
                                ])

OBR_7_TIMESTAMP_DEFAULT_STATE  = False
OBX_14_TIMESTAMP_DEFAULT_STATE = False

class Gui(QtGui.QMainWindow):
    
    def __init__(self,app):
        super(Gui, self).__init__()

        self.app       = app
        self.wid       = None
        self.xfile     = None
        self.headers   = QtGui.QVBoxLayout()
        self.tabs      = QtGui.QTabWidget()
        self.vbox      = QtGui.QVBoxLayout()
        self.fname     = ''
        self.configuration_window = None      
        self.hl7_dropdown_menu_items = None
        sys.excepthook = self._display_error_message

        self.initWindow()
        self.initMenu() 
        self.load_settings()       
        self.initPage()                
        self.runWindow()

    # PyQt Error Handling
    def _display_error_message(self, error_type, error_value, error_traceback):
    	trb = ''.join(traceback.format_exception(error_type, error_value, error_traceback))
    	QtGui.QMessageBox.critical(self, "FATAL ERROR", "An unexpected error occurred:\n%s\n\n%s" % (error_value, trb))


    def load_settings(self):
        '''Loads the default settings'''
        # Creates a Qsetting object with default settings
        self.default_settings = QtCore.QSettings(COMPANY_NAME, APPLICATION_NAME) 
        self.default_settings.setValue('HL7_segments', DEFAULT_HL7_SEGMENTS)
        self.default_settings.setValue('Configurationbox_segments', DEFAULT_CONFIGURATIONBOX_SEGMENTS_DATA)
        self.default_settings.setValue('Header Variables', HEADER_VARIABLES)
        self.default_settings.setValue('Calculated Variables', DEFAULT_CALCULATED_VARIABLES)
        self.default_settings.setValue('OBR 7 timestamp', OBR_7_TIMESTAMP_DEFAULT_STATE)
        self.default_settings.setValue('OBX 14 timestamp', OBX_14_TIMESTAMP_DEFAULT_STATE)

        self.default_settings.sync()
        
        # also load gui object with default settings
        # This is a list of HL7 segments, fields and the default number of values for each field
        self.current_configurationbox_segments_data = DEFAULT_CONFIGURATIONBOX_SEGMENTS_DATA
        self.current_hl7_segment_data = DEFAULT_HL7_SEGMENTS
        self.current_calculated_variables_data = DEFAULT_CALCULATED_VARIABLES
        self.current_obr_7_timestamp_state = OBR_7_TIMESTAMP_DEFAULT_STATE
        self.current_obx_14_timestamp_state = OBX_14_TIMESTAMP_DEFAULT_STATE                                                      
        self.current_header_variables_data_dict = HEADER_VARIABLES
        
    def initWindow(self):    
        logging.debug("initializing window") 

        QtGui.QToolTip.setFont(QtGui.QFont('SansSerif', 10))

        self.wid = QtGui.QWidget(self)
        self.setCentralWidget(self.wid)

        
    def initMenu(self):
        logging.debug("initializing Menu/Tool/Status Bars") 

        openFile = QtGui.QAction(QtGui.QIcon(':/excel.png'), "&Open Excel File", self)
        openFile.setShortcut("Ctrl+F")
        openFile.triggered.connect(self.fileOpen)

        # Add an action to open a configuration window that can be used to change some configuration/settings
        configwindow = QtGui.QAction("&Configuration", self)
        configwindow.setShortcut("F8")
        configwindow.triggered.connect(self.open_configuration_window)
       
        help_usermanual_Action = QtGui.QAction(QtGui.QIcon(':/usermanual.png'), "&User Manual", self)
        help_usermanual_Action.setShortcut("Alt+H")
        help_usermanual_Action.triggered.connect(self.helpMessage)       

        help_about_Action = QtGui.QAction("&About", self)
        help_about_Action.setShortcut("Alt+F12")
        help_about_Action.triggered.connect(self.helpMessage)  
        
        menubar = self.menuBar()
        
        # Add the different action menus on the File menubar
        f_menu = menubar.addMenu('&File')
        f_menu.addAction(openFile)        

        # Add the different action menus on the View menubar
        v_menu = menubar.addMenu('&View')
        v_menu.addAction(configwindow)
           
        # Add different action menus on the Help menubar
        h_menu = menubar.addMenu('&Help')
        h_menu.addAction(help_usermanual_Action)
        h_menu.addAction(help_about_Action)        
                
        self.toolbar = self.addToolBar('Open Excel File')
        self.toolbar.addAction(openFile)
        
        self.statusBar().showMessage('Ready')

        
    def center(self):
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        logging.debug("frame size is " + str(qr) + "center point is " + str(cp)) 
        qr.moveCenter(cp)
        self.move(qr.topLeft())

        
    def runWindow(self):

        self.resize(1000, 1000)
        self.center()
        self.setWindowTitle('Purple Panda')
        self.setWindowIcon(QtGui.QIcon(':/tool.png'))
        self.show()


    def update_main_window(self, update_hl7segments_tabs = False, update_hl7_configurationbox_menu = False, update_obr_timestamp = False, update_obx_timestamp = False, update_header_variables_list  = False, update_calculated_variables = False):
        ''' Applies the changed settings to the main window'''
        self.statusBar().showMessage('Wait! Applying config changes')

        if update_hl7_configurationbox_menu == True:
            if self.update_Menu_Table() == True:
                self.statusBar().showMessage('Done! The Table has been updated with the applied configuration settings')
            else:
                pass
                # Write code for an exception that might have occurred
        if update_hl7segments_tabs == True:
            if self.update_hl7_segment_boxes() == True:
                self.statusBar().showMessage('Done! HL7 segment boxes have been updated with the applied configuration settings')
        else:
            pass


    def compare_list_values(self, original, current):
        ''' Compares the original and current HL7 segments'''
        original_hl7_segment_list = [x.strip() for x in map(str, original.split(','))]
        current_hl7_segment_list = [y.strip() for y in map(str, current.split(','))]
        return bool(set(current_hl7_segment_list).difference(original_hl7_segment_list))


    def compare_dict_values(self, original_dict, current_dict):
        '''Compares the original and current configuration box segment values'''
        return bool(set(current_dict.items()).difference(original_dict.items()))



    def update_settings(self, cfg_window_settings):
        '''Will look at the differences between the current settings configured by the user and the default settings and apply changes wherever necessary'''          
        update_hl7segments_tabs = False
        update_hl7_configurationbox_menu = False
        update_obr_timestamp = False
        update_obx_timestamp = False
        update_header_variables_list  = False
        update_calculated_variables = False

        # Grabs the HL7 Segments data and compares with the defaults to see if it changed
        new_hl7_segment_data = cfg_window_settings.value('HL7_segments', type = str)
        if self.compare_list_values(self.current_hl7_segment_data, new_hl7_segment_data):
            print "Yes HL7 segments changed"
            update_hl7segments_tabs = True
            self.current_hl7_segment_data = new_hl7_segment_data

        # Grabs the Calculated Variables data and compares with the defaults to see if it changed
        new_calculated_variables_data = cfg_window_settings.value('Calculated Variables', type = str)
        if self.compare_list_values(self.current_calculated_variables_data, new_calculated_variables_data):
            print "Yes Calculated Variables have changed"
            update_calculated_variables = True
            self.current_calculated_variables_data = new_calculated_variables_data
 
        # Grabs the HL7 segment configuration box data and compares with the defaults to see if it changed
        new_configurationbox_data = cfg_window_settings.value('Configurationbox_segments').toPyObject()
        new_configurationbox_segments_data = OrderedDict()
        for key, value in new_configurationbox_data.items():
             new_configurationbox_segments_data[str(key)] = value

        if  self.compare_dict_values(self.current_configurationbox_segments_data, new_configurationbox_segments_data):
            print "Yes HL7 Menubox changed"
            update_hl7_configurationbox_menu = True
            self.current_configurationbox_segments_data = new_configurationbox_segments_data


        # Grabs the OBR7 and OBX14 checkbox states and compares with the defaults to see if they changed
        new_obr_7_timestamp_state = cfg_window_settings.value('OBR 7 timestamp').toBool()
        if (new_obr_7_timestamp_state ^ self.current_obr_7_timestamp_state):
            print "OBR 7 changed"
            update_obr_timestamp = True
            self.current_obr_7_timestamp_state = new_obr_7_timestamp_state

        new_obx_14_timestamp_state = cfg_window_settings.value('OBX 14 timestamp').toBool()
        if (new_obx_14_timestamp_state ^ self.current_obx_14_timestamp_state):
            print "OBX 14 changed"
            update_obx_timestamp = True
            self.current_obx_14_timestamp_state = new_obx_14_timestamp_state

        # Grabs the header variables list and compares with the defaults to see if they changed
        new_header_variables_data = cfg_window_settings.value('Header Variables').toPyObject()

        new_header_variables_data_dict = OrderedDict()
        for key, value in new_header_variables_data.items():
             new_header_variables_data_dict[str(key)] = str(value)

        if  self.compare_dict_values(new_header_variables_data_dict, self.current_header_variables_data_dict):
            print HEADER_VARIABLES
            print self.current_header_variables_data_dict
            print "Yes Header variables changed"
            update_header_variables_list  = True
            self.current_header_variables_data_dict = new_header_variables_data_dict


        # Once it determines what settings have been changed, the appropriate settings will only be applied to the mainwindow
        self.update_main_window(update_hl7segments_tabs, update_hl7_configurationbox_menu, update_obr_timestamp, update_obx_timestamp, update_header_variables_list)


    def open_configuration_window(self):
        ''' Opens the configuration window to make the updates'''

        if self.configuration_window is None:
            self.configuration_window = Configuration_Window(self.default_settings)
 
        if self.configuration_window.exec_():
            cfg_window_settings = self.configuration_window.get_current_settings()
            self.update_settings(cfg_window_settings)
    

    def update_Menu_Table(self):
        #Need to add code to update the table even when no tabs exist
        if not self.hl7_dropdown_menu_items:
            self.hl7_dropdown_menu_items = Hl7_menu()

        self.hl7_dropdown_menu_items.create_dropdown_items_from_dict(self.current_configurationbox_segments_data)

        for i in xrange(len(self.tabs)):            
            widg  = self.tabs.widget(i).layout()
            table = widg.itemAt(0).widget()
            for row in xrange(table.rowCount()):
                dropdown = table.cellWidget(row,1)
                menu     = QtGui.QMenu()
            
                self.make_hl7Menu(dropdown,menu,self.hl7_dropdown_menu_items.hl7_dropdown_dict)
                dropdown.setMenu(menu)

        return True

        
        
        
    def set_message_boxes(self):

        self.list_of_message_information = []
        self.messagebox_segments = self.current_hl7_segment_data.split(",")


        # self.messagebox.segments is statically defined in the init function
        # as self.messagebox_segments = ['MSH - Message Header Segment', 'PID - Patient Identification Segment', 'PV1 Segment - Patient Visit Information Segment', 'OBR - Observation Request Segment'] 

        for each_message in self.messagebox_segments:
            message_label = QtGui.QLabel(each_message)
            message_box_information = QtGui.QLineEdit()
            self.headers.addWidget(message_label)
            self.headers.addWidget(message_box_information)  
            self.list_of_message_information.append(message_box_information)


    def initPage(self):        
        logging.debug("Setting Page")  

        '''
        Message boxes
        '''
        # This function will generate, initialize mulitple message/vbox widgets with the appropriate labels and return the list of the boxes to be laid out
        self.set_message_boxes()

        '''
        action pane
        '''  
        # This will populate the "Map in HL7 format"     
        btn_map_hl7 = QtGui.QPushButton('Map in HL7 format', self)
        btn_map_hl7.setToolTip('Maps the selected Excel files columns to the various HL7 segments and generates an HL7 file\n' \
                        'Click on the excel icon on the top left to select and excel file')
        btn_map_hl7.clicked.connect(self.setMapping_hl7)

        # This will populate the "Map in CLBS format" 
        btn_map_clbs = QtGui.QPushButton('Map in CLBS format', self)
        btn_map_clbs.setToolTip('Maps the selected Excel files columns to the various HL7 segments and generates a CLBS file\n' \
                        'Click on the excel icon on the top left to select and excel file')
        btn_map_clbs.clicked.connect(self.setMapping_clbs)

        # This will populate the "Exit" button
        qbtn = QtGui.QPushButton('Exit', self)
        qbtn.setToolTip('Exit applicaton')        
        qbtn.clicked.connect(self.close)        
        self.apane = QtGui.QHBoxLayout()
        self.apane.addStretch(1)
        self.apane.addWidget(btn_map_hl7)
        self.apane.addWidget(btn_map_clbs)
        self.apane.addWidget(qbtn)   

        '''
        Page Layout
        '''        
        self.vbox.addLayout(self.headers)
        self.vbox.addWidget(self.tabs)     
        self.vbox.addLayout(self.apane)
        
        self.wid.setLayout(self.vbox)

    def make_hl7Menu(self,dropdown,menu,hl7_dictionary): 
        for key in hl7_dictionary:
            if key == " ":
                action = menu.addAction(key)
                action.triggered.connect(self.updateTable(dropdown,key))
                continue
            sub_menu = menu.addMenu(key)
            action = sub_menu.addAction(key)
            action.triggered.connect(self.updateTable(dropdown,key))
            if isinstance(hl7_dictionary[key], dict):
                self.make_hl7Menu(dropdown,sub_menu,hl7_dictionary[key])
            else:
                for dict_value in hl7_dictionary[key]:
                    action = sub_menu.addAction(dict_value)
                    action.triggered.connect(self.updateTable(dropdown,dict_value))

                
    def closeEvent(self, event):
        reply = QtGui.QMessageBox.question(self, 'Message',
                                           "Are you sure to quit?", QtGui.QMessageBox.Yes |
                                           QtGui.QMessageBox.No, QtGui.QMessageBox.No)        
        if reply == QtGui.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()                                            

    def helpMessage(self):
        ''' Opens the user manual for Purple Panda'''
        # should be more robust and needs to be changed 
        file = "help.doc"
        os.system(file)

    def updateTable(self,dropdown,dict_value):
        return lambda : dropdown.setText(dict_value)
        
    def fileOpen(self):
        self.statusBar().showMessage('Opening File wait')


        _file = QtGui.QFileDialog.getOpenFileName(self, 'Open File')
        fname = str(_file)
        logging.debug("Opened " + fname)

        if not fname:
            pass
        else:
            self.filename, file_extension = os.path.splitext(os.path.basename(fname))
            if ('xlsx' in file_extension) or ('xlsm' in file_extension):
                self.fname = fname
                self.xfile = pd.ExcelFile(self.fname)
                self.popTable()
            else:
            #add try catch and create a message box with file open error
                self.helpMessage()

        self.statusBar().showMessage('Ready')       


    def sheets_to_import(self, sheetnames):
        '''Will return a list of sheets to be considered from the excel sheet'''

        sheets_not_to_import = deepcopy(QE_INPUT_SHEETS_TO_DISCARD)

        for sheet in sheetnames:
            matchObj = re.match( r'OUTPUT_.+', sheet, re.I)
            if matchObj:
                sheets_not_to_import.append(matchObj.group())


        sheets_to_import = [sheet for sheet in sheetnames if sheet not in sheets_not_to_import]

        return sheets_to_import
 


    def popTable(self):
        '''populates the table with tabs etc..'''

        # Check if there are any existing tabs and clear them
        if self.tabs.count() > 0:
            self.tabs.clear()


        if not self.hl7_dropdown_menu_items:
            self.hl7_dropdown_menu_items = Hl7_menu()
        self.hl7_dropdown_menu_items.create_dropdown_items_from_dict(self.current_configurationbox_segments_data)

        # for each sheet create a tab and populate with a table
        tables  = []
        xtabs   = []
        vboxs   = []

        # Obtain the main sheets which have information for creating the HL7 files and exclude all the miscellaneous files of the QE sheet
        self.sheet_names = self.sheets_to_import(self.xfile.sheet_names)

        for sht_idx, sheet in enumerate(self.sheet_names):
            logging.debug("sheet found : " + sheet)            

            # This will populate the table with the contents in the xcolumns list
            tab   = QtGui.QWidget()
            table = QtGui.QTableWidget()
            vbox  = QtGui.QVBoxLayout()

            tables.append(table)
            xtabs.append(tab)
            vboxs.append(vbox)
            
            # This will generate the table contents that will next be used to fill the table
            table_contents = []
            
            #This will populate a list of tuples with the items in the tuple being a column (ex Column 1) and a dropdown_box widget for each which
            #will be populated with the items from the hl7_segments list
        
            # configure the table with a few extra settings
            tables[sht_idx].horizontalHeader().setVisible(False)
            tables[sht_idx].verticalHeader().setVisible(False)
            tables[sht_idx].setRowCount(len(self.xfile.parse(sheet).columns))
            tables[sht_idx].setColumnCount(2)            
                
            for column in self.xfile.parse(sheet).columns:
                logging.debug("column found : " + column)

                label    = QtGui.QLabel(column)
                dropdown = QtGui.QPushButton()
                menu     = QtGui.QMenu()
            
                self.make_hl7Menu(dropdown,menu,self.hl7_dropdown_menu_items.hl7_dropdown_dict)
                dropdown.setMenu(menu)
                
                # If the NodeID exists in the columns of the excel sheet, then statically set it to PV1-3
                if label.text() == "NodeID":
                    dropdown.setText("PV1-3")
                else:
                    pass
                    # Make it like a pop up that either aborts the tool or allows the user to say okay and continue to add one manually
                    #self.load_popup("No NodeID column found in the QE Input Sheet")
                
                table_contents.append((label,dropdown))

            # This populates the table with table_contents's contents in the given row and cell.o
            # row Each will have two columns with the first column populated with column_1 which is the label and column_2 which is the Dropdown 
            ix = 0
            for column_1, column_2 in table_contents:
                tables[sht_idx].setCellWidget(ix, 0, column_1)
                tables[sht_idx].setCellWidget(ix, 1, column_2)
                ix = ix + 1

            header = tables[sht_idx].horizontalHeader()
            header.setResizeMode(0, QtGui.QHeaderView.Stretch)
            header.setResizeMode(1, QtGui.QHeaderView.Stretch)

            # Add table to the tab
            vboxs[sht_idx].addWidget(tables[sht_idx])        
            xtabs[sht_idx].setLayout(vboxs[sht_idx])
            self.tabs.addTab(xtabs[sht_idx],sheet)

                    
    def load_popup(self,text):
        popup = QtGui.QMessageBox( QtGui.QMessageBox.NoIcon, "Success!!! ", text,  QtGui.QMessageBox.NoButton)
        # Get the layout
        l = popup.layout()
        # Hide the default button
        l.itemAtPosition( l.rowCount() - 1, 0 ).widget().hide()
        popup.exec_()


    def generic_set_mapping(self):
        ''' This runs all the setup necessary to build the sim files'''

        self.create_hl7_dict_values = Create_Hl7_Data()
        self.create_hl7_dict_values.set_hl7segments_count(self.current_configurationbox_segments_data)

        # Obtain the texts and map them to the respective HL7 Segmenclsts
        self.create_hl7_dict_values.split_message_box_segments(self.list_of_message_information)

        # update status bar
        self.statusBar().showMessage('Wait!...Mapping...')

        # Reads data from the current tab's sheet and table   
        self.sheetab_string = str(self.tabs.tabText (self.tabs.currentIndex()))
        sheettab_table = self.tabs.currentWidget().layout().itemAt(0).widget()

        calculated_var_list = map(int, self.current_calculated_variables_data.split(','))

        self.create_hl7_dict_values.read_excel_data(self.xfile, self.sheetab_string, sheettab_table, calculated_var_list, self.current_header_variables_data_dict)

        # Generates a dictionary with all the data that can now be written to a file
        self.final_hl7_message = self.create_hl7_dict_values.generate_hl7_message_data()

        self.additional_seg_results = self.get_static_header_segments()

    def get_static_header_segments(self):
    	'''Gets any static header segment info to be added for every HL7 message'''
    	message_box_dict = OrderedDict()
    	additional_seg_results = []
        for messagebox_info in self.list_of_message_information:
        	message_box_dict[(str(messagebox_info.text()).split('|'))[0]] = str(messagebox_info.text())

        if self.current_hl7_segment_data != DEFAULT_HL7_SEGMENTS:
        	diff_msg_boxes_list = set([x.strip() for x in map(str, self.current_hl7_segment_data.split(','))]).difference([y.strip() for y in map(str, DEFAULT_HL7_SEGMENTS.split(','))])
        	additional_seg_results = [message_box_dict[diff_segment] for diff_segment in diff_msg_boxes_list if diff_segment in message_box_dict.keys()]
        else:
        	additional_seg_results = None

        return additional_seg_results

    def update_hl7_segment_boxes(self):
        while (self.headers.count() != 0):
            widg = self.headers.itemAt(0).widget()
            self.headers.removeWidget(widg)
            widg.setParent(None)
        
        self.set_message_boxes()

        return True


    def data_to_write_to_file(self, data_dict=None):
    	''' Gets data from the data_dict and outputs data to be written to text file'''

    	obx_fields_list = []

    	first_dict_item_contents = next(iter(data_dict))
    	msh_result = self.create_hl7_dict_values.get_key_val_data(first_dict_item_contents, key_val='MSH', timestamp_state=False, timestamp_index=None, index_to_write=None)
    	pid_result = self.create_hl7_dict_values.get_key_val_data(first_dict_item_contents, key_val='PID', timestamp_state=False, timestamp_index=None, index_to_write=None)
    	pv1_result = self.create_hl7_dict_values.get_key_val_data(first_dict_item_contents, key_val='PV1', timestamp_state=False, timestamp_index=None, index_to_write=None)
    	obr_result = self.create_hl7_dict_values.get_key_val_data(first_dict_item_contents, key_val='OBR', timestamp_state = self.current_obr_7_timestamp_state, timestamp_index=7, index_to_write=None)


        # Write the generated OBX timestamps back to the QE sheet
        if self.current_obx_14_timestamp_state is True:
            # Obtain the NodeID to be able to copy the OBX timestamp to the respective row in the QE sheet
            match_found = re.match(r"PV1[|](.*)[|](.*)[|](.*)", pv1_result, re.I)
            if match_found:
            	current_node_id = match_found.group(3) 
            	list_of_indexes = self.create_hl7_dict_values.df.index[self.create_hl7_dict_values.df['PV1-3'] == current_node_id].tolist()
            	index_cycle = cycle(list_of_indexes)
            else:
            	print "No Node ID found"
        else:
            pass # Nothing needs to be done

        
        for each_dict_item in data_dict:
        	if self.current_obx_14_timestamp_state is True:
        		df_index_to_write = next(index_cycle)
        	else:
        		df_index_to_write = None
        	obx_result = self.create_hl7_dict_values.get_key_val_data(each_dict_item, key_val='OBX', timestamp_state = self.current_obx_14_timestamp_state, timestamp_index=14, index_to_write=df_index_to_write)
        	obx_fields_list.append(obx_result)

        return msh_result, pid_result, pv1_result, obr_result, obx_fields_list


    def setMapping_hl7(self):
        '''This will perform the Mapping of the Columns to the different HL7 segments and Generate the HL7 file'''

        # # Makes sure that when the Map button is clicked, a filename has been selected and if not, it will generate a popup message
        if not self.xfile:
            self.helpMessage()

        else:
            #create popup
            self.statusBar().showMessage('Wait!...Data being written...')
            self.generic_set_mapping()
            
            # This is the part where the data is going be written with proper format to the HL7 file
            with open("AllVariables " + self.sheetab_string + ".hl7", "w") as text_file:
                # Obtains all the data that is grouped under each unique NodeID
                for each_unique_id, each_unique_id_info in self.final_hl7_message.iteritems():
                    obx_fields_list = []

                    if any(each_unique_id_info):
                        msh_result, pid_result, pv1_result, obr_result, obx_fields_list = self.data_to_write_to_file(data_dict=each_unique_id_info)
                        text_file.write("")
                        text_file.write(msh_result + "\n")
                        text_file.write(pid_result + "\n")
                        text_file.write(pv1_result + "\n")
                        text_file.write(obr_result + "\n")

                        if self.additional_seg_results is not None:
                        	for additional_seg_result in self.additional_seg_results:
                        		text_file.write(additional_seg_result + "\n")

                        for each_obx_string in obx_fields_list:
                            text_file.write(each_obx_string + "\n")
                        text_file.write("\n") 


            text_to_print = "Mapping is Done!!!\n" + "The file is stored in the following location:\n" + os.getcwd() + "\AllVariables " + self.sheetab_string + ".hl7"
            self.create_hl7_dict_values.df.to_excel("AllVariables " + self.sheetab_string + ".xlsx", index=False)             
            self.load_popup(text_to_print)
            self.statusBar().showMessage('Ready')


    def setMapping_clbs(self):
        '''This will perform the Mapping of the Columns to the different HL7 segments and Generate the CLBS file'''

        # # Makes sure that when the Map button is clicked, a filename has been selected and if not, it will generate a popup message
        if not self.xfile:
            self.helpMessage()

        else:
            #create popup
            self.statusBar().showMessage('Wait!...Mapping...')
            self.generic_set_mapping()
            
            # This is the part where the data is going be written with proper format to the HL7 file
            with open("AllVariables " + self.sheetab_string + ".clbs", "w") as text_file:

                # Obtains all the data that is grouped under each unique NodeID
                for each_unique_id, each_unique_id_info in self.final_hl7_message.iteritems():
                    obx_fields_list = []

                    if any(each_unique_id_info):
                    	msh_result, pid_result, pv1_result, obr_result, obx_fields_list = self.data_to_write_to_file(data_dict=each_unique_id_info)

                        text_file.write("[BEGIN DEVICE]\n")
                        text_file.write("<VT>" + msh_result + "|<CR>" + "\n")
                        text_file.write(pid_result + "|<CR>" + "\n")
                        text_file.write(pv1_result + "|<CR>" + "\n")
                        text_file.write(obr_result + "|<CR>" + "\n")

                        if self.additional_seg_results is not None:
                        	for additional_seg_result in self.additional_seg_results:
                        		text_file.write(additional_seg_result + "\n")


                        for each_obx_string in obx_fields_list:
                            text_file.write(each_obx_string + "|<CR>" + "\n")
                            text_file.write("<FS><CR>\n")
                            text_file.write("[END DEVICE]\n")
                            text_file.write("\n")

            text_to_print = "Mapping is Done!!!\n" + "The file is stored in the following location\n" + os.getcwd() + "\AllVariables " + self.sheetab_string + ".clbs" 
            self.create_hl7_dict_values.df.to_excel("AllVariables " + self.sheetab_string + ".xlsx", index=False)            
            self.load_popup(text_to_print)
            self.statusBar().showMessage('Ready')


if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    ex = Gui(app)
    sys.exit(app.exec_())



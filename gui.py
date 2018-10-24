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
from config import configurationbox_segments, default_hl7_segments, obr_7_timestamp_default_state, obx_14_timestamp_default_state

logging.basicConfig(level=logging.DEBUG,
                    format='[%(threadName)s] %(message)s',
)


COMPANY_NAME = "CapsuleTech"
APPLICATION_NAME = "MainWindow"
QE_INPUT_SHEETS_TO_DISCARD = ["Template Cover Sheet", "Summary", "General", "Report", "Comparison_Report"]
self.msg_header_variables = [758, 1914, 1929, 1930, 1931, 2255, 2949, 3097, 6510, 6511, 8036]
self.non_obx_message_variables = [1190, 2583, 5815, 5816]

class Gui(QtGui.QMainWindow):
    
    def __init__(self,app):
        super(Gui, self).__init__()

        self.app       = app
        self.wid       = None
        self.xfile     = None
        self.tabs      = QtGui.QTabWidget()
        self.fname     = ''
        self.configuration_window = None      
        self.hl7_dropdown_menu_items = None
        
        self.initWindow()
        self.initMenu() 
        self.load_settings()       
        self.initPage()                
        self.runWindow()


    def load_settings(self):
        '''Loads the default settings'''
        # Creates a Qsetting object with default settings
        self.default_settings = QtCore.QSettings(COMPANY_NAME, APPLICATION_NAME) 
        self.default_settings.setValue('HL7_segments', default_hl7_segments)
        self.default_settings.setValue('Configurationbox_segments', configurationbox_segments)
        self.default_settings.setValue('OBR 7 timestamp', obr_7_timestamp_default_state)
        self.default_settings.setValue('OBX 14 timestamp', obx_14_timestamp_default_state)

        self.default_settings.sync()


        # state = bool(self.default_settings.value('OBR 7 timestamp'))
        # print state

        
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


    def update_main_window(self, update_hl7segments_tabs = False, update_hl7_configurationbox_menu = False, update_obr_timestamp = False, update_obx_timestamp = False):
        ''' Applies the changed settings to the main window'''
        self.statusBar().showMessage('Wait! Applying config changes')
        if update_hl7_configurationbox_menu == True:
            if self.update_Menu_Table() == True:
                self.statusBar().showMessage('Done! The Table has been updated with the applied configuration settings')
            else:
                pass
                # Write code for an exception that might have occurred
        else:
            pass


    def compare_hl7_segments(self, original, current):
        ''' Compares the original and current HL7 segments'''
        original_hl7_segment_list = original.split(',')
        current_hl7_segment_list = map(str, current.split(','))
        return bool(set(current_hl7_segment_list).difference(original_hl7_segment_list))


    def compare_hl7_configurationbox_segment_values(self, original_dict, current_dict):
        '''Compares the original and current configuration box segment values'''
        return bool(set(current_dict.items()).difference(original_dict.items()))



    def update_settings(self, cfg_window_settings):
        '''Will look at the differences between the current settings configured by the user and the default settings and apply changes wherever necessary'''          
        update_hl7segments_tabs = False
        update_hl7_configurationbox_menu = False
        update_obr_timestamp = False
        update_obx_timestamp = False

        # Grabs the HL7 Segments data and compares with the defaults to see if it changed
        self.current_hl7_segment_data = cfg_window_settings.value('HL7_segments', type = str)
        if self.compare_hl7_segments(default_hl7_segments, self.current_hl7_segment_data):
            print "Yes HL7 segments changed"
            update_hl7segments_tabs = True
            #self.update_hl7_segments()

        # Grabs the HL7 segment configuration box data and compares with the defaults to see if it changed
        current_configurationbox_data = cfg_window_settings.value('Configurationbox_segments').toPyObject()
        self.current_configurationbox_segments_data = OrderedDict()
        for key, value in current_configurationbox_data.items():
             self.current_configurationbox_segments_data[str(key)] = value

        if  self.compare_hl7_configurationbox_segment_values(configurationbox_segments, self.current_configurationbox_segments_data):
            print "Yes HL7 Menubox changed"
            update_hl7_configurationbox_menu = True


        # Grabs the OBR7 and OBX14 checkbox states and compares with the defaults to see if they changed
        self.obr_7_timestamp_state = cfg_window_settings.value('OBR 7 timestamp').toBool()
        if (obr_7_timestamp_default_state ^ self.obr_7_timestamp_state):
            print "OBR 7 changed"
            update_obr_timestamp = True

        self.obx_14_timestamp_state = cfg_window_settings.value('OBX 14 timestamp').toBool()
        if (obx_14_timestamp_default_state ^ self.obx_14_timestamp_state):
            print "OBX 14 changed"
            update_obx_timestamp = True


        # Once it determines what settings have been changed, the appropriate settings will only be applied to the mainwindow
        self.update_main_window(update_hl7segments_tabs, update_hl7_configurationbox_menu, update_obr_timestamp, update_obx_timestamp)


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

        self.list_of_message_boxes = []
        self.list_of_message_information = []

        self.messagebox_segments = default_hl7_segments.split(",")


        # self.messagebox.segments is statically defined in the init function
        # as self.messagebox_segments = ['MSH - Message Header Segment', 'PID - Patient Identification Segment', 'PV1 Segment - Patient Visit Information Segment', 'OBR - Observation Request Segment'] 

        for each_message in self.messagebox_segments:
            message_box = QtGui.QVBoxLayout()
            message_label = QtGui.QLabel(each_message)
            self.message_box_information = QtGui.QLineEdit()
            message_box.addWidget(message_label)
            message_box.addWidget(self.message_box_information)  
            self.list_of_message_boxes.append(message_box)
            self.list_of_message_information.append(self.message_box_information)


        return self.list_of_message_boxes


    def initPage(self):        
        logging.debug("Setting Page") 


        self.vbox = QtGui.QVBoxLayout()
 

        '''
        Message boxes
        '''
        # This function will generate, initialize mulitple message/vbox widgets with the appropriate labels and return the list of the boxes to be laid out
        self.list_of_message_boxes = self.set_message_boxes()

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
        # The following lists the order in which the page will be laid out. First the message boxes are placed, following the table followed by the buttons
        for each_vbox in self.list_of_message_boxes:
            self.vbox.addLayout(each_vbox)

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
        help_msg = QtGui.QMessageBox.question(self, 'Select a valid file',
                                              'Please select an xlsx file by clicking on the excel icon, ' \
                                              'Map the columns to respective HL7 segments and '\
                                              'click on the Map icon to generate the HL7 file', QtGui.QMessageBox.Ok)

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
            matchObj = re.match( r'OUTPUT_\w+', sheet, re.I)
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
        self.hl7_dropdown_menu_items.create_dropdown_items_from_dict(configurationbox_segments)

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
        self.create_hl7_dict_values.set_hl7segments_count(configurationbox_segments)

        # Obtain the texts and map them to the respective HL7 Segmenclsts
        self.create_hl7_dict_values.split_message_box_segments(self.list_of_message_information)

        # update status bar
        self.statusBar().showMessage('Wait!...Mapping...')

        # Reads data from the current tab's sheet and table   
        self.sheetab_string = str(self.tabs.tabText (self.tabs.currentIndex()))
        sheettab_table = self.tabs.currentWidget().layout().itemAt(0).widget()

        self.create_hl7_dict_values.read_excel_data(self.xfile, self.sheetab_string, sheettab_table)

        # # This will remove any variables that dont have to be part of the HL7 message like the I:E ratio variables etc
        # # The variables are listed in two lists one is the self.msg_header_variables and self.non_obx_message_variables
        # self.create_hl7_dict_values.remove_variables_from_list(self.non_obx_message_variables)
        #self.create_hl7_dict_values.remove_variables_from_list(self.msg_header_variables)

        # Generates a dictionary with all the data that can now be written to a file
        self.final_hl7_message = self.create_hl7_dict_values.generate_hl7_message_data()


    def setMapping_hl7(self):
        '''This will perform the Mapping of the Columns to the different HL7 segments and Generate the HL7 file'''


        # -----------------------TO DO : CREATE A DICTIONARY WITH EACH KEY AS THE DIFFERENT FIELDS AND VALUES AS EMPTY ---------------------------------


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

                    # For each grouped data, it goes through each dictionary and gets all the respective data segments
                    # To note for that each group, the MSH/PID/PV1 and OBR segments should be the same and only the OBX segments would change
                    for each_dict_item in each_unique_id_info:
                                

                        msh_result = self.create_hl7_dict_values.get_msh_data(each_dict_item)
                        pid_result = self.create_hl7_dict_values.get_pid_data(each_dict_item)
                        pv1_result = self.create_hl7_dict_values.get_pv1_data(each_dict_item)
                        obr_result = self.create_hl7_dict_values.get_obr_data(each_dict_item)
                        obx_result = self.create_hl7_dict_values.get_obx_data(each_dict_item)
                        obx_fields_list.append(obx_result)

                    text_file.write("")
                    text_file.write(msh_result + "\n")
                    text_file.write(pid_result + "\n")
                    text_file.write(pv1_result + "\n")
                    text_file.write(obr_result + "\n")
                    for each_obx_string in obx_fields_list:
                        text_file.write(each_obx_string + "\n")
                    text_file.write("\n")                                           

            text_to_print = "Mapping is Done!!!\n" + "The file is stored as " + "AllVariables " + self.sheetab_string + ".hl7"             
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

                    # For each grouped data, it goes through each dictionary and gets all the respective data segments
                    # To note for that each group, the MSH/PID/PV1 and OBR segments should be the same and only the OBX segments would change
                    for each_dict_item in each_unique_id_info:              

                        msh_result = self.create_hl7_dict_values.get_msh_data(each_dict_item)
                        pid_result = self.create_hl7_dict_values.get_pid_data(each_dict_item)
                        pv1_result = self.create_hl7_dict_values.get_pv1_data(each_dict_item)
                        obr_result = self.create_hl7_dict_values.get_obr_data(each_dict_item)
                        obx_result = self.create_hl7_dict_values.get_obx_data(each_dict_item)
                        obx_fields_list.append(obx_result)



                    text_file.write("[BEGIN DEVICE]\n")
                    text_file.write("<VT>" + msh_result + "|<CR>" + "\n")
                    text_file.write(pid_result + "|<CR>" + "\n")
                    text_file.write(pv1_result + "|<CR>" + "\n")
                    text_file.write(obr_result + "|<CR>" + "\n")
                    for each_obx_string in obx_fields_list:
                        text_file.write(each_obx_string + "|<CR>" + "\n")
                    text_file.write("<FS><CR>\n")
                    text_file.write("[END DEVICE]\n")
                    text_file.write("\n")

            text_to_print = "Mapping is Done!!!\n" + "The file is stored as " + "AllVariables " + self.sheetab_string + ".clbs"            
            self.load_popup(text_to_print)
            self.statusBar().showMessage('Ready')



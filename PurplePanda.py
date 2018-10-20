#!/usr/bin/python2.7
# encoding=utf8
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import os
import logging
from PyQt4 import QtGui, QtCore
import pandas as pd
import xlrd
import img_qr
from collections import defaultdict
from operator import attrgetter
import re
import copy
from collections import OrderedDict
import time
from hl7menu import Hl7_menu
from create_hl7_data import Create_Hl7_Data



logging.basicConfig(level=logging.CRITICAL,
                    format='[%(threadName)s] %(message)s',
)

# This is a list of HL7 segments, fields and the default number of values for each field
configurationbox_segments = OrderedDict([('MSH Field Count', '17'),
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
                                         ('OBX Subcomponent Count', '5'),                                        
                                         ])

default_hl7_segments = "MSH, PID, PV1, OBR"




class Gui(QtGui.QMainWindow):
    
    def __init__(self,app):
        super(Gui, self).__init__()



        self.app       = app
        self.wid       = None
        self.popup     = None
        self.popup_lbl = None
        self.prog      = None
        self.table    = None
        self.fname    = ''
        
       
        self.xcolumns  = ['Column1','Column2','Column3','...','...','ColumnN']

        self.msg_header_variables = [758, 1914, 1929, 1930, 1931, 2255, 2949, 3097, 6510, 6511, 8036]

        self.non_obx_message_variables = [1190, 2583, 5815, 5816]


        self.initWindow()
        self.initMenu()        
        self.initPage()                
        self.runWindow()
        
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


    def open_configuration_window(self):
        self.dialog = Configuration_Window(self)
        self.dialog.show()
        print configurationbox_segments



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
        logging.debug("setting Page") 

        vbox = QtGui.QVBoxLayout()

        '''
        Message boxes
        '''
        # This function will generate, initialize mulitple message/vbox widgets with the appropriate labels and return the list of the boxes to be laid out
        self.list_of_message_boxes = self.set_message_boxes()

 
        '''
        table
        '''
        # This function will populate a table and its contents to be laid out 
        self.popTable()
      
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
        apane = QtGui.QHBoxLayout()
        apane.addStretch(1)
        apane.addWidget(btn_map_hl7)
        apane.addWidget(btn_map_clbs)
        apane.addWidget(qbtn)   

        '''
        Page Layout
        '''        
        # The following lists the order in which the page will be laid out. First the message boxes are placed, following the table followed by the buttons
        for each_vbox in self.list_of_message_boxes:
            vbox.addLayout(each_vbox)

        vbox.addWidget(self.table)
     
        vbox.addLayout(apane)
        
        self.wid.setLayout(vbox)

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
            if 'xlsx' in file_extension:
                self.fname = fname
                _xfile = pd.read_excel(self.fname)
                self.xcolumns = _xfile.columns
                logging.debug("Columns found : " + str(self.xcolumns))            
                self.popTable()
            else:
            #add try catch and create a message box with file open error
                self.helpMessage()

        self.statusBar().showMessage('Ready')        


    def popTable(self):

        # This will populate a table if no table exists. It will hide the row and column attributes (i.e.1 , 2 etc.)
        if not self.table:
            self.table = QtGui.QTableWidget()
            self.table.horizontalHeader().setVisible(False)
            self.table.verticalHeader().setVisible(False)            
           
        # This will populate the table with the contents in the xcolumns list and the number of columns are always set to 2 
        self.table.setRowCount(len(self.xcolumns))
        self.table.setColumnCount(2)


        # This will generate the table contents that will next be used to fill the table
        table_contents = []
        
        #This will populate a list of tuples with the items in the tuple being a column (ex Column 1) and a dropdown_box widget for each which
        #will be populated with the items from the hl7_segments list

        hl7_dropdown_menu_items = Hl7_menu()
        hl7_dropdown_menu_items.create_dropdown_items_from_dict(configurationbox_segments)



        for column in self.xcolumns:
            label    = QtGui.QLabel(column)
            dropdown = QtGui.QPushButton()
            menu     = QtGui.QMenu()
            
            self.make_hl7Menu(dropdown,menu,hl7_dropdown_menu_items.hl7_dropdown_dict)
            dropdown.setMenu(menu)

            table_contents.append((label,dropdown))

             
        # This populates the table with table_contents's contents in the given row and cell.
        # Each row will have two columns with the first column populated with column_1 which is the label and column_2 which is the Dropdown 
        index = 0
        for column_1, column_2 in table_contents:
            self.table.setCellWidget(index, 0, column_1)
            self.table.setCellWidget(index, 1, column_2)
            index = index + 1

        header = self.table.horizontalHeader()
        header.setResizeMode(0, QtGui.QHeaderView.Stretch)
        header.setResizeMode(1, QtGui.QHeaderView.Stretch)



    def load_popup(self):
        self.popup = QtGui.QMessageBox( QtGui.QMessageBox.NoIcon, "Wait", "Mapping in Progress !!!", QtGui.QMessageBox.NoButton)
        # Get the layout
        l = self.popup.layout()
        # Hide the default button
        l.itemAtPosition( l.rowCount() - 1, 0 ).widget().hide()

        self.prog      = QtGui.QProgressBar()
        self.popup_lbl = QtGui.QLabel()
        
        # Add the progress bar at the bottom (last row + 1) and first column with column span
        l.addWidget(self.prog,l.rowCount(), 0, 1, l.columnCount(), QtCore.Qt.AlignCenter )
        l.addWidget(self.popup_lbl,l.rowCount(), 1, 1, l.columnCount(), QtCore.Qt.AlignCenter )        

        self.popup.show()


    def generic_set_mapping(self):
        ''' This runs all the setup necessary to build the sim files'''

        self.create_hl7_dict_values = Create_Hl7_Data()
        self.create_hl7_dict_values.set_hl7segments_count(configurationbox_segments)

        # Obtain the texts and map them to the respective HL7 Segmenclsts
        self.create_hl7_dict_values.split_message_box_segments(self.list_of_message_information)

        # update progress bar
        self.prog.setValue(20)

        time.sleep(30)

        # Reads data from excel sheet and formats it 
        self.create_hl7_dict_values.read_excel_data(self.fname, self.table)

        # This will remove any variables that dont have to be part of the HL7 message like the I:E ratio variables etc
        # The variables are listed in two lists one is the self.msg_header_variables and self.non_obx_message_variables
        self.create_hl7_dict_values.remove_variables_from_list(self.non_obx_message_variables)
        self.create_hl7_dict_values.remove_variables_from_list(self.msg_header_variables)


        # update progress bar
        self.prog.setValue(40)

        time.sleep(30)

        # Generates a dictionary with all the data that can now be written to a file
        self.final_hl7_message = self.create_hl7_dict_values.generate_hl7_message_data()

        # update progress bar
        self.prog.setValue(60)
   

    def setMapping_hl7(self):
        '''This will perform the Mapping of the Columns to the different HL7 segments and Generate the HL7 file'''


        # -----------------------TO DO : CREATE A DICTIONARY WITH EACH KEY AS THE DIFFERENT FIELDS AND VALUES AS EMPTY ---------------------------------


        # # Makes sure that when the Map button is clicked, a filename has been selected and if not, it will generate a popup message
        if not self.fname:
            self.helpMessage()

        else:
            #create popup
            self.load_popup()
            self.generic_set_mapping()

            # This is the part where the data is going be written with proper format to the HL7 file
            with open("AllVariables purplepanda " + self.filename + ".hl7", "w") as text_file:

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


            self.prog.setValue(100)
            text_to_print = "Mapping is Done!!!\n" + "The file is stored as " + "AllVariables PurplePanda " + self.filename + ".hl7"             
            self.popup_lbl.setText(text_to_print)



    def setMapping_clbs(self):
        '''This will perform the Mapping of the Columns to the different HL7 segments and Generate the CLBS file'''

        # # Makes sure that when the Map button is clicked, a filename has been selected and if not, it will generate a popup message
        if not self.fname:
            self.helpMessage()

        else:
            #create popup
            self.load_popup()
            self.generic_set_mapping()

            # This is the part where the data is going be written with proper format to the HL7 file
            with open("AllVariables purplepanda " + self.filename + ".clbs", "w") as text_file:

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

            self.prog.setValue(100)
            text_to_print = "Mapping is Done!!!\n" + "The file is stored as " + "AllVariables PurplePanda " + self.filename + ".clbs"             
            self.popup_lbl.setText(text_to_print)




class Configuration_Window(QtGui.QMainWindow):

    def __init__(self, parent=None):
        super(Configuration_Window, self).__init__()

        self.initWindow()
        self.initPage()                   
        self.runWindow()


    def initWindow(self):    
        QtGui.QToolTip.setFont(QtGui.QFont('SansSerif', 10))
        self.wid = QtGui.QWidget(self)
        self.setCentralWidget(self.wid)


    def center(self):
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()   
        self.move(qr.topLeft())

    def initPage(self):
        logging.debug("setting Page") 


        hbox = QtGui.QHBoxLayout()

        msg_label = QtGui.QLabel("HL7 Segments")
        msg_linedit = QtGui.QLineEdit()
        msg_linedit.setText(default_hl7_segments)


        hbox.addWidget(msg_label)
        hbox.addWidget(msg_linedit) 



        vbox = QtGui.QVBoxLayout()
          
        self.popTable()    

        '''
        action pane
        '''  
        # This button will apply any configuration changes applied by the user     
        btn_apply = QtGui.QPushButton('Apply', self)
        btn_apply.setToolTip('Applies the Configuration set by the user')
        btn_apply.clicked.connect(self.apply_configuration)
       
        apane = QtGui.QHBoxLayout()
        apane.addStretch(1)
        apane.addWidget(btn_apply)
        # apane.addWidget(btn_exit)

        vbox.addLayout(hbox)

        vbox.addWidget(self.table)

        vbox.addLayout(hbox)   
    
        vbox.addLayout(apane)

        self.wid.setLayout(vbox)

    def popTable(self):
        self.table = QtGui.QTableWidget()
        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)            
           
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

        header = self.table.horizontalHeader()
        header.setResizeMode(0, QtGui.QHeaderView.Stretch)
        header.setResizeMode(1, QtGui.QHeaderView.Stretch)


           
    def runWindow(self):
        self.resize(1000, 1000)
        self.center()
        self.setWindowTitle('Configuration')
        self.setWindowIcon(QtGui.QIcon(':/tool.png'))


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

    def apply_configuration(self):
        ''' This will apply the new values to the configurationbox_segments dictionary'''
        global configurationbox_segments

        # First read all the values from the configuration window
        for row in range(0,self.table.rowCount()):
            column0_text = str(self.table.cellWidget(row,0).text())
            column1_text = str(self.table.cellWidget(row,1).currentText())


        # Modify the global configurationbox segment dictionary with the new values
            if column0_text in configurationbox_segments.keys():
                configurationbox_segments[column0_text] = column1_text

        # Will have to modify the dropdown hl7 menu list once the configurationbox_segments dictionary is updated
        hl7_dropdown_menu_items = Hl7_menu()
        hl7_dropdown_menu_items.create_dropdown_items_from_dict(configurationbox_segments)



def main():
    app = QtGui.QApplication(sys.argv)
    ex = Gui(app)
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()

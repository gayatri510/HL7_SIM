#!/usr/bin/python2.7
import sys
from collections import defaultdict, OrderedDict
import pandas as pd
import xlrd
import copy
from datetime import datetime, timedelta
from itertools import cycle
from PyQt4 import QtGui, QtCore



SEGMENT_TERMINATOR = "0x0D"   # Segment Terminator : Terminates a segment record
FIELD_SEPERATOR= "|"          # Field Separator : Separates two adjacent data fields within a segment
COMPONENT_SEPERATOR = "^"     # Component Separator : Separates adjacent components of data fields
SUBCOMPONENT_SEPERATOR = "&"  # Subcomponent Separator : Separates adjacent subcomponents of data fields
REPITITION_SEPERATOR = "~"    # Repitition Separator : Seperates multiple occurrences of a field


class Create_Hl7_Data():

    def __init__(self):


        self.count = 0
        self.hl7_dict_values = OrderedDict()
        self.current_time = datetime.now()


    def set_hl7segments_count(self, segments_and_count_dict_values):
        ''' This method will get the segments' fields, components and subcomponents values and set it for other methods to use'''

        segments_and_counts_dict = OrderedDict(defaultdict(list))

        for segment in segments_and_count_dict_values.keys():
            value = segments_and_count_dict_values[segment]

            # Segment will have the value 'MSH Field Count' and in segment name, we are splitting the word and the first word in the list 
            # will be the segment name in this case 'MSH' and segment_message will have "Field"
            segment_name = (segment.split())[0]
            segment_message = (segment.split())[1]

            
            #  We need to check if this segment exists in the segments_and_counts_dict and if not add it to the dictionary as a key
            if segment_name in segments_and_counts_dict.keys():
                segments_counts_list.append(value)
                segments_and_counts_dict[segment_name] = segments_counts_list
            else:
                segments_counts_list = []
                segments_counts_list.append(value)
                segments_and_counts_dict[segment_name] = segments_counts_list

        # Create empty dictionaries for each segment. The following will generate something similar to {MSH:{1:{1:{1 : ''}}}} and the blanks will hold the value for
        # MSH-1-1-1 

        for every_key in segments_and_counts_dict.keys():
            self.hl7_dict_values[every_key] = OrderedDict()
            for field_count in range(1, int(segments_and_counts_dict[every_key][0]) + 1):
                self.hl7_dict_values[every_key][field_count] = OrderedDict()
                for component_count in range(1, int(segments_and_counts_dict[every_key][1]) + 1):
                    self.hl7_dict_values[every_key][field_count][component_count] = OrderedDict()
                    for subcomponent_count in range(1, int(segments_and_counts_dict[every_key][2]) + 1):
                        self.hl7_dict_values[every_key][field_count][component_count][subcomponent_count] = ''
 

    
    def populate_each_seg_dict(self, segment_box_data=None):
        '''Populates the dictionary values for each segment sent in the segment_name'''
        field_values = str(segment_box_data).strip().split(FIELD_SEPERATOR)
        key_val = field_values[0]

        if key_val not in self.hl7_dict_values.keys():
            return

        if field_values[0] == "MSH":
            pass
        else:
            field_values = field_values[1:]

        for index, field_value in enumerate(field_values):
            component_values = field_value.split(COMPONENT_SEPERATOR)
            if len(component_values) == 1:
                self.hl7_dict_values[key_val][index + 1][1][1] = field_value
            elif len(component_values) == 0:
                self.hl7_dict_values[key_val][index + 1][1][1] = ''

            # If the field has multiple componetns like "000000 & 111111^^^^AN"
            elif len(component_values) > 1:
                for comp_index, component_value in enumerate(component_values):
                    # Need to check if the components have multiple subcomponents like "000000 & 111111"
                    subcomponent_values = component_value.split(SUBCOMPONENT_SEPERATOR)
                    # If this list is empty, it means this component doesn't have multiple subcomponents
                    if len(subcomponent_values) == 1:
                        self.hl7_dict_values[key_val][index + 1][comp_index + 1][1] = component_value
                    else:
                        for s_index, subcompvalues in enumerate(subcomponent_values):
                            self.hl7_dict_values[key_val][index + 1][comp_index + 1][s_index + 1] = subcompvalues


    def split_message_box_segments(self, list_of_message_information):
        # Obtain the texts and map them to the respective HL7 Segments
            for each_info in list_of_message_information:
                # This will split the information on '|' and generate a list of the different field values          
                self.populate_each_seg_dict(segment_box_data = each_info.text())


    def remove_variables_from_list(self, var_list):
        ''' This will remove the variables from the variables_list list from the pandas dataframe'''

        for var_id in var_list:
            self.df = self.df[self.df['Capsule Variable ID'] != str(var_id)]


    def read_excel_data(self, excel_sheet, sheet_to_read_from, data_table, calculated_var_list = None, non_obx_variables_dict = None):
        ''' This method will read data from the excel sheet into a pandas data frame and also read the data from the table on the tool'''


        # This part of the code will work with the QE input sheet, read the excel into a pandas data frame
        # Rename the columns to the appropriate HL7 segments and then generate the HL7 file
        # The Code below ensures that each column data is taken as string instead of doing any conversions to float, int etc.

        df_actual = excel_sheet.parse(sheet_to_read_from)
        self.df = excel_sheet.parse(sheet_to_read_from, converters= {col: str for col in df_actual.columns.values.tolist()})

        #Remove the calculated and non_obx Capsule Variable IDs that will not be part of the OBX messages
        self.remove_variables_from_list(calculated_var_list)
        self.non_obx_variables_dict = non_obx_variables_dict
        # Traverse through each row of the table, get the column 0 and 1 strings and if an HL7 segment mapping exists, then
        # rename the excel sheet with the name present in column 0 with that present in column 1
        self.mapping_list = []
        for row in range(0,data_table.rowCount()):
            column0_text = str(data_table.cellWidget(row,0).text())
            column1_text = str(data_table.cellWidget(row,1).text())
            # If the column1 is not empty and has a valid string, then replace string in column 0 with column 1
            if column1_text:
                self.mapping_list.append(column1_text)
                self.df.rename(columns={column0_text : column1_text},inplace=True)


        #self.df.to_excel("Renamed.xlsx", index=False)    


    def set_data(self, data=None, data_dict=None, item_to_split=None):
        '''Sets the data in the data_dict with the args specifiying the key, field, comp and subcomp locations'''

        # Check for the length of the list. If it is 4, it implies it is in the format of ex: OBX-3-1-1
        if len(item_to_split.split('-')) == 4:
            header, field, comp, subcomp = item_to_split.split('-')
            data_dict[header][int(field)][int(comp)][int(subcomp)] = data

        # Check for the length of the list. If it is 3, it implies it is in the format of ex: OBX-3-1
        if len(item_to_split.split('-')) == 3:
            header, field, comp = item_to_split.split('-')
            subcomponents_list = data.split(SUBCOMPONENT_SEPERATOR)
            # Case where OBX-3-1 is in format of 1&2
            if len(subcomponents_list) > 1:
                for s_index, subcomponent_value in enumerate(subcomponents_list):
                    data_dict[header][int(field)][int(comp)][s_index + 1] = subcomponent_value
            # Case where OBX-3-1 is in format of 1
            elif len(subcomponents_list) == 1:
                data_dict[header][int(field)][int(comp)][1] = data
                

        # Check for the length of the list. If it is 2, it implies it is in the format of ex: OBX-3
        if len(item_to_split.split('-')) == 2:
            header, field = item_to_split.split('-')
            # It can be in the format of 1^2 or 1&2^3&4 or just 1, all three cases have to be covered     
            components_list = data.split(COMPONENT_SEPERATOR)
            # Case 1: OBX-3 is in format of 1^2 or 1&2^3 or 1&2^3&4
            if (len(components_list) > 1):
                for each_comp_index, each_comp in enumerate(components_list):
                    subcomponents_list = each_comp.split(SUBCOMPONENT_SEPERATOR)
                    if len(subcomponents_list) == 1:
                        data_dict[header][int(field)][each_comp_index + 1][1] = each_comp
                    elif len(subcomponents_list) > 1:
                        for s_index, subcomponent_value in enumerate(subcomponents_list):
                            data_dict[header][int(field)][each_comp_index + 1][s_index + 1] = subcomponent_value   
            # Case 2: OBX-3 is in format of 1
            elif len(components_list) == 1:
                data_dict[header][int(field)][1][1] = data
                            

    def generate_hl7_message_data(self):
        ''' This method will create a dictionary which will have all the data necessary to generate the simulation file'''
        self.non_obx_variables_list = []

        # Once the sheet is ready with the respective columns and the header segment information is also obtained,
        # traverse through each row of the excel sheet and update the self.hl7_dict_values dictionary with each OBX/PID/PV1/OBR messages etc.
        # To note that a new dict will be created with each row of the excel sheet

        # The following will try to get all the unique NodeIDs in the excel sheet
        try:
            unique_message_ids = [x for x in map(str, self.df['NodeID'].unique()) if x != 'nan']
        except KeyError:
            unique_message_ids = [x for x in map(str, self.df['PV1-3'].unique()) if x != 'nan']

        self.num_of_unique_ids = len(unique_message_ids)

        # Create a big message_dictionary whose key will be the unique Node ID 
        self.hl7_file_message = defaultdict(list)
        for unique_msg_id in unique_message_ids:
            self.hl7_file_message[unique_msg_id] = []


        for row_index, row_entry in self.df.iterrows():
            if row_entry['Capsule Variable ID'] in self.non_obx_variables_dict.keys():
                try:
                    self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], row_entry['Value HL7']))
                except KeyError:
                    try:
                        self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], row_entry['Value CLBS']))
                    except KeyError:
                        try:
                            self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], row_entry['Value']))
                        except KeyError:
                            try:
                                self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], row_entry['OBX-5']))
                            except KeyError:
                                self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], row_entry['OBX-5-1']))
                
            else:
                self.hl7_dict_values_each_msg = {}
                self.hl7_dict_values_each_msg = copy.deepcopy(self.hl7_dict_values)
                for item in self.mapping_list:
                    # Grab the value from that particular column except if it is empty, an Attribute Error is thrown, then set it to ""
                    try:
                        row_value = row_entry[item].strip()
                    except AttributeError:
                        row_value = ""

                    try:
                        dt = datetime.strptime(row_value, "%m/%d/%Y - %H:%M:%S.%f")
                        row_value = dt.strftime("%Y%m%d%H%M%S")
                    except ValueError:
                        pass

                    self.set_data(data=row_value, data_dict=self.hl7_dict_values_each_msg, item_to_split=item)

                # Traverse through each row of the sheet, obtain the Unique ID (in this case Node ID is used) and append that
                # hl7_dict_values dictionary to the appropriate list associated with that Unique ID
                try:
                    message_id = str(row_entry['NodeID'])
                except KeyError:
                    message_id = str(row_entry['PV1-3'])

                try:
                    self.hl7_file_message[message_id].append(self.hl7_dict_values_each_msg)
                except KeyError:
                    QtGui.QMessageBox.critical(self, "FATAL ERROR", "No NodeID was found at row_index %s in the excel sheet" % (str(row_index)))
  


                self.hl7_file_message = OrderedDict(sorted(self.hl7_file_message.items(), key=lambda t: t[0])) 



        # Remove the non_obx_variables from the list once their values are obtained for further computation
        self.remove_variables_from_list([tupleelement[0] for tupleelement in self.non_obx_variables_list])
        # Remove PV1-3 / NodeID/ 1931 variable from the self.non_obx_variables_list if any else it will incorrectly override any existing NodeIDs and mess up the messages
        self.non_obx_variables_list = [i for i in self.non_obx_variables_list if i[0] != '1931']
        myIterator = cycle(range(self.num_of_unique_ids))
        
        # Insert the non-obx variables in the self.hl7_file_message before returning it 
        for caps_id, var_value in self.non_obx_variables_list:
            nodeid_curr = list(self.hl7_file_message.keys())[myIterator.next()]
            self.set_data(data=var_value, data_dict=self.hl7_file_message[nodeid_curr][0], item_to_split=self.non_obx_variables_dict[caps_id])

        return self.hl7_file_message

        
    def remove_trailing_delimiters(self, string_to_trim, string_delimiter):
        ''' Will trim any string based off of the given delimiters'''
        # This part of the code will remove any trailing '|'s
        string_temp =  string_to_trim.split(string_delimiter)
        while string_temp:
            if string_temp[-1] == '':
                string_temp.pop()
            else:
                break

        return string_delimiter.join(string_temp)

    def get_key_val_data(self, dictionary_item, key_val=None, timestamp_state=False, timestamp_index=None, index_to_write=None):
        ''' Gets the key_val segment data'''
        component_part = []        
        for f_key, f_value in dictionary_item[key_val].iteritems():
            subcomponent_part = []
            for comp_key, comp_value in f_value.iteritems(): 
                subcomp_list = []
                for subcomp_key, subcomp_value in comp_value.iteritems():
                    subcomp_list.append(subcomp_value)
                subcomponent_part.append(self.remove_trailing_delimiters(SUBCOMPONENT_SEPERATOR.join(subcomp_list), SUBCOMPONENT_SEPERATOR))  # 1&&&&
            component_part.append(self.remove_trailing_delimiters(COMPONENT_SEPERATOR.join(subcomponent_part), COMPONENT_SEPERATOR))  # 1&&&&^2&&&&&
        

        if key_val == 'OBX' and 'OBX-1' not in self.mapping_list:
            self.count = self.count + 1
            component_part[0] = str(self.count)
        else:
            pass

        
        # Setting timestamp from the current timestamp + 1s delay for every segment that has timestamp = True
        if timestamp_state is True and timestamp_index is not None:
            self.current_time = self.current_time + timedelta(seconds = 1)         
            component_part[timestamp_index - 1] = "%d%d%d%d%d%d" % (self.current_time.year, self.current_time.month, self.current_time.day, self.current_time.hour, self.current_time.minute, self.current_time.second)
            
            if index_to_write is not None:
                try:
                    self.df.loc[index_to_write, "MeasurementTime"] = self.current_time.strftime("%m/%d/%Y - %H:%M:%S")
                except AttributeError:
                    pass # Add error handling
       
        if key_val == 'MSH':
            result = self.remove_trailing_delimiters(FIELD_SEPERATOR.join(component_part), FIELD_SEPERATOR)
        else:
            result = self.remove_trailing_delimiters((key_val + FIELD_SEPERATOR + FIELD_SEPERATOR.join(component_part)), FIELD_SEPERATOR)


        return result

#!/usr/bin/env python2.7
# Imports ######################################################################
import sys
from collections import defaultdict, OrderedDict
import pandas as pd
import xlrd
import copy
from datetime import datetime, timedelta
from itertools import cycle
from PyQt4 import QtGui, QtCore

FIELD_SEPERATOR = "|"         # Field Separator : Separates two adjacent data fields within a segment
COMPONENT_SEPERATOR = "^"     # Component Separator : Separates adjacent components of data fields
SUBCOMPONENT_SEPERATOR = "&"  # Subcomponent Separator : Separates adjacent subcomponents of data fields
REPITITION_SEPERATOR = "~"    # Repitition Separator : Seperates multiple occurrences of a field


class Create_Hl7_Data():
    def __init__(self):
        self.count = 0
        self.hl7_dropdown_dict = OrderedDict()
        self.hl7_dict_values = OrderedDict()
        self.current_time = datetime.now()

    def get_segments_and_count_dict(self, hl7_menu_dict):
        '''Create's a dictionary with the HL7 segments and their corresponding default fields, components and subcomponent values'''
        # Iterate through the dictionary hl7_menu_dict which is in the following form Ex: ('MSH Field Count', '17'),
        # obtain the counts for the fields, components and subcomponents and create a dictionary to populate the
        # dropdown menu appropriately
        segments_and_counts_dict = OrderedDict(defaultdict(list))
        for key, value in hl7_menu_dict.iteritems():
            segment_name = (key.split())[0]
            segments_and_counts_dict.setdefault(segment_name, []).append(value)

        return segments_and_counts_dict

    def create_dropdown_items_from_dict(self, hl7_menu_dict):
        ''' Populate hl7_dropdown_dict dictionary with the dropdown items that have to be populated in a table'''
        segments_and_counts_dict = self.get_segments_and_count_dict(hl7_menu_dict)
        for key in segments_and_counts_dict.keys():
            self.hl7_seg_field_dict = OrderedDict()
            for seq_count in range(1, int(segments_and_counts_dict[key][0]) + 1):
                self.hl7_dropdown_dict[key] = OrderedDict()
                seg_seq = key + "-" + str(seq_count)
                self.hl7_seg_field_dict[seg_seq] = OrderedDict()
                self.hl7_seg_components_dict = OrderedDict(defaultdict(list))
                for component_count in range(1, int(segments_and_counts_dict[key][1]) + 1):
                    seg_component_seq = seg_seq + "-" + str(component_count)
                    self.hl7_seg_subcomponents = []
                    for subcomponent_count in range(1, int(segments_and_counts_dict[key][2]) + 1):
                        self.hl7_seg_subcomponents.append(seg_component_seq + "-" + str(subcomponent_count))
                        self.hl7_seg_components_dict[seg_component_seq] = self.hl7_seg_subcomponents
                    self.hl7_seg_field_dict[seg_seq] = self.hl7_seg_components_dict

                self.hl7_dropdown_dict[key] = self.hl7_seg_field_dict

    def set_hl7segments_count(self, hl7_menu_dict):
        ''' Obtain the segments' fields, components and subcomponents values and initialize the hl7_dict_values dictionary'''
        segments_and_counts_dict = self.get_segments_and_count_dict(hl7_menu_dict)
        # Create empty dictionaries for each segment. The following will generate something similar to
        # {MSH:{1:{1:{1 : ''}}}} and the blanks will hold the value for MSH-1-1-1
        for every_key in segments_and_counts_dict.keys():
            self.hl7_dict_values[every_key] = OrderedDict()
            for field_count in range(1, int(segments_and_counts_dict[every_key][0]) + 1):
                self.hl7_dict_values[every_key][field_count] = OrderedDict()
                for component_count in range(1, int(segments_and_counts_dict[every_key][1]) + 1):
                    self.hl7_dict_values[every_key][field_count][component_count] = OrderedDict()
                    for subcomponent_count in range(1, int(segments_and_counts_dict[every_key][2]) + 1):
                        self.hl7_dict_values[every_key][field_count][component_count][subcomponent_count] = ''
    
    def get_message_boxes_data(self, message_information_dict):
        ''' Parse through data in each message box text entered by the user in the main window'''
        for key_val, value in message_information_dict.iteritems():
            if not value.text():  # If no data is entered in the text box
                raise Exception("The string entered in the %s segment is expected to start with %s but instead is blank" % (key_val, key_val))
            else:   # Check if the first field value matches the name of the textbox the data is entered in
                field_values = str(value.text()).strip().split(FIELD_SEPERATOR)
                if key_val == field_values[0]:
                    self.populate_each_seg_dict(segment_field_values=field_values)
                else:
                    raise Exception("The string entered in the %s segment is expected to start with %s but instead is starting with %s" % (key_val, key_val, field_values[0]))

    def populate_each_seg_dict(self, segment_field_values=None):
        '''Get data from each segment's field values and populate the hl7_dict_values dictionary'''
        key_val = segment_field_values[0]
        if key_val not in self.hl7_dict_values.keys():
            return
        if segment_field_values[0] == "MSH":
            pass
        else:
            segment_field_values = segment_field_values[1:]

        for index, field_value in enumerate(segment_field_values):
            component_values = field_value.split(COMPONENT_SEPERATOR)
            if len(component_values) == 1:   # string in the field_value is in the form of "A"
                try:
                    self.hl7_dict_values[key_val][index + 1][1][1] = field_value
                except KeyError:
                    raise Exception("The %s segment entered by the user has more fields than the default %s Field Count"
                                    " set in the Configuration Page. Please change the default settings to a higher number in the Configuration Page" % (key_val, key_val))
            elif len(component_values) == 0:   # string in the field_value is blank
                try:
                    self.hl7_dict_values[key_val][index + 1][1][1] = ''
                except KeyError:
                    raise Exception("The %s segment entered by the user has more fields than the default %s Field Count"
                                    " set in the Configuration Page. Please change the default settings to a higher number in the Configuration Page and rerun the Mapping" % (key_val, key_val))

            # If the field has multiple components like "000000 & 111111^^^^AN"
            elif len(component_values) > 1:     # string in the field_value is in the form of "A^B"
                for comp_index, component_value in enumerate(component_values):
                    # Need to check if the components have multiple subcomponents like "000000 & 111111"
                    subcomponent_values = component_value.split(SUBCOMPONENT_SEPERATOR)
                    # If this list has only one value, component doesn't have multiple subcomponents
                    if len(subcomponent_values) == 1:
                        try:
                            self.hl7_dict_values[key_val][index + 1][comp_index + 1][1] = component_value
                        except KeyError:
                            raise Exception("The %s segment entered by the user has more components than the default %s Component Count"
                                            " set in the Configuration Page. Please change the default settings to a higher number in the Configuration Page and rerun the Mapping" % (key_val, key_val))
                    elif len(subcomponent_values) > 1:  # The component has multiple subcomponents
                        for s_index, subcompvalues in enumerate(subcomponent_values):
                            try:
                                self.hl7_dict_values[key_val][index + 1][comp_index + 1][s_index + 1] = subcompvalues
                            except KeyError:
                                raise Exception("The %s segment entered by the user has more subcomponents than the default %s Subcomponent Count"
                                                " set in the Configuration Page. Please change the default settings to a higher number in the Configuration Page and rerun the Mapping" % (key_val, key_val))
                                                
    def read_excel_data(self, excel_sheet=None, sheet_to_read_from=None, data_table=None, calculated_var_list=None, non_obx_variables_dict=None):
        '''
        Read data from the sheet_to_read_from sheet of the excel_sheet into a pandas data frame
        Read the data from the table on the tools' main window and
        rename the columns of the sheet_to_read_from to the appropriate HL7 segment text
        which will be used later to generate the HL7 file
        '''
        # Ensure each column data is taken as string instead of doing any conversions to float, int etc.
        self.sheet = sheet_to_read_from
        df_actual = excel_sheet.parse(self.sheet)
        self.df = excel_sheet.parse(self.sheet, converters={col: str for col in df_actual.columns.values.tolist()})
        # Remove the calculated variables from the dataframe that will not be part of the OBX messages
        if calculated_var_list is not None:
            self.remove_variables_from_list(calculated_var_list)
        # Create a dictionary to store the non obx variables' related data later in the code
        self.non_obx_variables_dict = non_obx_variables_dict
        # Traverse through each row of the table, get the column 0 and 1 strings and if an HL7 segment mapping exists, then
        # rename the sheet's columns
        self.mapping_list = []
        for row in range(0, data_table.rowCount()):
            column0_text = str(data_table.cellWidget(row, 0).text())
            column1_text = str(data_table.cellWidget(row, 1).text())
            # If the column1 is not empty and has a valid string, then replace string in column 0 with column 1
            if column1_text:
                self.mapping_list.append(column1_text)
                self.df.rename(columns={column0_text: column1_text}, inplace=True)

    def remove_variables_from_list(self, var_list):
        ''' Remove the var_list variables from the pandas dataframe'''
        for var_id in var_list:
            self.df = self.df[self.df['Capsule Variable ID'] != str(var_id)]

    def generate_hl7_message_data(self):
        ''' Create a dictionary which will have all the data necessary to generate the simulation file'''
        self.non_obx_variables_list = []
        # Get all the unique NodeIDs from the sheet that is being parsed
        try:
            unique_message_ids = [x for x in map(str, self.df['NodeID'].unique()) if x != 'nan']
        except KeyError:
            unique_message_ids = [x for x in map(str, self.df['PV1-3'].unique()) if x != 'nan']

        self.num_of_unique_ids = len(unique_message_ids)

        # Create a hl7_file_message dictionary with the unique Node IDs as keys
        self.hl7_file_message = defaultdict(list)
        for unique_msg_id in unique_message_ids:
            self.hl7_file_message[unique_msg_id] = []
        for row_index, row_entry in self.df.iterrows():  # Parse through each sheet's rows
            # Append any Capsule variable IDs present in the sheet that shouldn't be mapped to OBX segments to non_obx_variables_list
            if row_entry['Capsule Variable ID'] in self.non_obx_variables_dict.keys():
                try:
                    self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], self.get_row_value(row_entry['Value HL7'])))
                except KeyError:
                    try:
                        self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], self.get_row_value(row_entry['Value CLBS'])))
                    except KeyError:
                        try:
                            self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], self.get_row_value(row_entry['Value'])))
                        except KeyError:
                            try:
                                self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], self.get_row_value(row_entry['OBX-5'])))
                            except KeyError:
                                self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], self.get_row_value(row_entry['OBX-5-1'])))
            else:
                self.hl7_dict_values_each_msg = {}
                self.hl7_dict_values_each_msg = copy.deepcopy(self.hl7_dict_values)
                for item in self.mapping_list:
                    # Grab the value from that particular column except if it is empty set it to ""
                    row_value = self.get_row_value(row_entry[item])
                    if (len(row_value.split(FIELD_SEPERATOR))) > 1:
                        raise Exception("The data found at row %i and column mapped to %s in the %s sheet has the | character. This value is invalid. Please check the data!!!" % (row_index, item, self.sheet))
                    # If the field is in the format of a timestamp, parse it and format it appropriately to add to the simfile
                    try:
                        dt = datetime.strptime(row_value, "%m/%d/%Y - %H:%M:%S.%f")
                        row_value = dt.strftime("%Y%m%d%H%M%S")
                    except ValueError:
                        pass   # If the ValueError exception occurs, then the value is not a timestamp entry and will be ignored
                    self.set_data(data=row_value, data_dict=self.hl7_dict_values_each_msg, item_to_split=item)

                # Traverse through each row of the sheet, obtain the Unique ID (in this case Node ID is used) and append that
                # hl7_dict_values dictionary to the appropriate list associated with that Unique ID
                try:
                    message_id = str(row_entry['PV1-3'])
                except KeyError:
                    message_id = str(row_entry['NodeID'])

                try:
                    self.hl7_file_message[message_id].append(self.hl7_dict_values_each_msg)
                except KeyError:
                    pass  # If no NodeID is found, that data will be ignored and won't be in the simulation file
                    #raise Exception("No NodeID was found at row_index %i in the %s sheet" % (row_index, self.sheet))
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

    def get_row_value(self, row_entry_value=None):
        '''Determine if data is valid else set it to an empty string'''
        try:
            row_value = row_entry_value.strip()
        except AttributeError:
            row_value = ""
        return row_value

    def set_data(self, data=None, data_dict=None, item_to_split=None):
        '''Populate the data_dict dictionary with the data string appropriately'''
        # Check for the length of the list. If it is 4, it implies it is in the format of ex: OBX-3-1-1
        if len(item_to_split.split('-')) == 4:
            if data:  # If data is valid and not empty
                header, field, comp, subcomp = item_to_split.split('-')
                data_dict[header][int(field)][int(comp)][int(subcomp)] = data

        # Check for the length of the list. If it is 3, it implies it is in the format of ex: OBX-3-1
        if len(item_to_split.split('-')) == 3:
            if data:  # If data is valid and not empty
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
            if data:  # If data is valid and not empty
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

    def get_key_val_data(self, dictionary_item, key_val=None, timestamp_state=False, timestamp_index=None, index_to_write=None):
        ''' Get the data from the dictionary_item's key = key_val'''
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
            self.current_time = self.current_time + timedelta(seconds=1)
            try:
                component_part[timestamp_index - 1] = "%d%d%d%d%d%d" % (self.current_time.year, self.current_time.month, self.current_time.day, self.current_time.hour, self.current_time.minute, self.current_time.second)
            except IndexError:
                raise Exception("Please check the %s Field Count value in the Configuration Page."
                                " It has to be atleast %i in order for the tool to generate the %s-%i timestamp" % (key_val, timestamp_index, key_val, timestamp_index))
            if index_to_write is not None:
                self.df.loc[index_to_write, "MeasurementTime"] = self.current_time.strftime("%m/%d/%Y - %H:%M:%S")
        if key_val == 'MSH':
            result = self.remove_trailing_delimiters(FIELD_SEPERATOR.join(component_part), FIELD_SEPERATOR)
        else:
            result = self.remove_trailing_delimiters((key_val + FIELD_SEPERATOR + FIELD_SEPERATOR.join(component_part)), FIELD_SEPERATOR)
        return result

    def remove_trailing_delimiters(self, string_to_trim, string_delimiter):
        ''' Trim string_to_trim based off of the given string_delimiter'''
        # This part of the code will remove any trailing '|'s
        string_temp = string_to_trim.split(string_delimiter)
        while string_temp:
            if string_temp[-1] == '':
                string_temp.pop()
            else:
                break

        return string_delimiter.join(string_temp)

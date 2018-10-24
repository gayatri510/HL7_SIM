#!/usr/bin/python2.7

import sys
from collections import defaultdict, OrderedDict
import pandas as pd
import xlrd
import copy
from datetime import datetime, timedelta


class Create_Hl7_Data():

    def __init__(self):

        # Segment Terminator : Terminates a segment record
        self.segment_terminator = "0x0D"
        # Field Separator : Separates two adjacent data fields within a segment
        self.field_separator = "|"
        # Component Separator : Separates adjacent components of data fields
        self.component_separator = "^"
        # Subcomponent Separator : Separates adjacent subcomponents of data fields
        self.subcomponent_separator = "&"
        # Repitition Separator : Seperates multiple occurrences of a field
        self.repitition_separator = "~"

        self.hl7_dict_values = OrderedDict()
        self.count = 0
        self.current_time = datetime.now()

        # --------------- WILL HAVE TO MOVE TO THE CONFIG SECTION ONLY AND NOT BE PRESENT HERE
        self.header_variables = OrderedDict([('6510', 'MSH-3'),
                                ('7914', 'MSH-3-1'),
                                ('1911', 'MSH-3-2'),
                                ('6511', 'MSH-4'),
                                ('6512', 'MSH-5'),
                                ('6513', 'MSH-6'),                                      
                                ('2255', 'MSH-7'),
                                ('9569', 'PID-3'),
                                ('1929', 'PID-3-1'),                                        
                                ('1930', 'PID-5'),
                                ('8340', 'PID-5-1'),
                                ('2901', 'PID-5-2'),
                                ('1220', 'PID-7'),
                                ('170',  'PID-8'),
                                ('8036', 'PID-18'),   
                                ('6521', 'PV1-2'),
                                ('9593', 'OBR-2'),
                                ('8102', 'OBR-3')                                   
                                ])


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
                    self.hl7_dict_values[every_key][field_count][component_count] = defaultdict(list)
 

    def split_message_box_segments(self, list_of_message_information):
        # Obtain the texts and map them to the respective HL7 Segments
            for each_info in list_of_message_information:
                messagebox_text = each_info.text()

                # This will split the information on '|' and generate a list of the different field values
                field_values = str(messagebox_text).strip().split(self.field_separator)


                #Based on the field values, place the different fields in to the appropriate fields of the dictionary

                #This checks if the field belongs to the MSH Segment
                
                if field_values[0] == "MSH":
                    #field_values = field_values[1:]
                    for index, field_value in enumerate(field_values):

                        # Need to check if the fields have multiple components
                        component_values = field_value.split(self.component_separator)

                        # If this list is empty, it means this field doesn't have multiple components
                        if len(component_values) < 1:
                            self.hl7_dict_values['MSH'][index + 1][1] = field_value
                        
                        # If the field has multiple componetns like "000000 & 111111^^^^AN"
                        else:
                            for comp_index, component_value in enumerate(component_values):
                                
                                # Need to check if the components have multiple subcomponents like "000000 & 111111"
                                subcomponent_values = component_value.split(self.subcomponent_separator)

                                # If this list is empty, it means this component doesn't have multiple subcomponents
                                if len(subcomponent_values) < 1:
                                    self.hl7_dict_values['MSH'][index + 1][comp_index + 1] = component_value
                                else:
                                    self.hl7_dict_values['MSH'][index + 1][comp_index + 1] = subcomponent_values




                # This checks if the field belongs to the PID Segment
                # Ex: PID|||000000^^^^AN||UNKNOWN|||||||||||||000000
                           
                if field_values[0] == "PID":
         
                    field_values = field_values[1:]
                    for index, field_value in enumerate(field_values):

                        # Need to check if the fields have multiple components
                        component_values = field_value.split(self.component_separator)

                        # If this list is empty, it means this field doesn't have multiple components
                        if len(component_values) < 1:
                            self.hl7_dict_values['PID'][index + 1][1] = field_value
                        
                        # If the field has multiple componetns like "000000 & 111111^^^^AN"
                        else:
                            for comp_index, component_value in enumerate(component_values):
                                
                                # Need to check if the components have multiple subcomponents like "000000 & 111111"
                                subcomponent_values = component_value.split(self.subcomponent_separator)

                                # If this list is empty, it means this component doesn't have multiple subcomponents
                                if len(subcomponent_values) < 1:
                                    self.hl7_dict_values['PID'][index + 1][comp_index + 1] = component_value
                                else:
                                    self.hl7_dict_values['PID'][index + 1][comp_index + 1] = subcomponent_values



                # This checks if the field belongs to the PV1 Segment
                # Ex: PV1|||SURG^^OR4
                
                if field_values[0] == "PV1":

                    field_values = field_values[1:]
                    for index, field_value in enumerate(field_values):

                        # Need to check if the fields have multiple components
                        component_values = field_value.split(self.component_separator)
 

                        # If this list is empty, it means this field doesn't have multiple components
                        if len(component_values) < 1:
                            self.hl7_dict_values['PV1'][index + 1][1] = field_value
                        
                        # If the field has multiple componetns like "000000 & 111111^^^^AN"
                        else:
                            for comp_index, component_value in enumerate(component_values):
                                
                                # Need to check if the components have multiple subcomponents like "000000 & 111111"
                                subcomponent_values = component_value.split(self.subcomponent_separator)

                                # If this list is empty, it means this component doesn't have multiple subcomponents
                                if len(subcomponent_values) < 1:
                                    self.hl7_dict_values['PV1'][index + 1][comp_index + 1] = component_value
                                else:
                                    self.hl7_dict_values['PV1'][index + 1][comp_index + 1] = subcomponent_values

 


                # This checks if the field belongs to the PV1 Segment
                # Ex: OBR|||||||20150217131700
                
                if field_values[0] == "OBR":

                    field_values = field_values[1:]
                    for index, field_value in enumerate(field_values):

                        # Need to check if the fields have multiple components
                        component_values = field_value.split(self.component_separator)
 

                        # If this list is empty, it means this field doesn't have multiple components
                        if len(component_values) < 1:
                            self.hl7_dict_values['OBR'][index + 1][1] = field_value
                        
                        # If the field has multiple componetns like "000000 & 111111^^^^AN"
                        else:
                            for comp_index, component_value in enumerate(component_values):
                                
                                # Need to check if the components have multiple subcomponents like "000000 & 111111"
                                subcomponent_values = component_value.split(self.subcomponent_separator)

                                # If this list is empty, it means this component doesn't have multiple subcomponents
                                if len(subcomponent_values) < 1:
                                    self.hl7_dict_values['OBR'][index + 1][comp_index + 1] = component_value
                                else:
                                    self.hl7_dict_values['OBR'][index + 1][comp_index + 1] = subcomponent_values




    def read_excel_data(self, excel_sheet, sheet_to_read_from, data_table, var_list = None):
        ''' This method will read data from the excel sheet into a pandas data frame and also read the data from the table on the tool'''


        # This part of the code will work with the QE input sheet, read the excel into a pandas data frame
        # Rename the columns to the appropriate HL7 segments and then generate the HL7 file
        # The Code below ensures that each column data is taken as string instead of doing any conversions to float, int etc.

        df_actual = excel_sheet.parse(sheet_to_read_from)
        self.df = excel_sheet.parse(sheet_to_read_from, converters= {col: str for col in df_actual.columns.values.tolist()})

        self.remove_variables_from_list(var_list)


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

            # # Optionally rename this into a different sheet for debugging purposes later
            self.df.to_excel("Renamed_QE_sheet.xlsx", index=False)


    def remove_variables_from_list(self, var_list):
        ''' This will remove the variables from the variables_list list from the pandas dataframe'''

        for var_id in var_list:
            self.df = self.df[self.df['Capsule Variable ID'] != str(var_id)]

    def generate_hl7_message_data(self):
        ''' This method will create a dictionary which will have all the data necessary to generate the simulation file'''
        self.non_obx_variables_list = []

        # Once the sheet is ready with the respective columns and the header segment information is also obtained,
        # traverse through each row of the excel sheet and update the self.hl7_dict_values dictionary with each OBX/PID/PV1/OBR messages etc.
        # To note that a new dict will be created with each row of the excel sheet

        # The following will try to get all the unique NodeIDs in the excel sheet
        try:
            unique_message_ids = map(str, self.df['NodeID'].unique())
        except KeyError:
            unique_message_ids = map(str, self.df['PV1-3'].unique())

        #print unique_message_ids


        # Create a big message_dictionary whose key will be the unique Node ID 
        self.hl7_file_message = defaultdict(list)
        for unique_msg_id in unique_message_ids:
            self.hl7_file_message[unique_msg_id] = []



        for _, row_entry in self.df.iterrows():
            if row_entry['Capsule Variable ID'] in self.header_variables.keys():
                try:
                    self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], row_entry['Value'], row_entry['NodeID']))
                except KeyError:
                    self.non_obx_variables_list.append((row_entry['Capsule Variable ID'], row_entry['Value'], row_entry['PV1-3']))
                
                # do something different
            else:
                self.hl7_dict_values_each_msg = {}
                self.hl7_dict_values_each_msg = copy.deepcopy(self.hl7_dict_values)
                for item in self.mapping_list:
                    row_value = row_entry[item].strip()
                    if row_value != row_value:
                        row_value = ""

                    mapping_list_components =  item.split('-')


                    # Check for the length of the list. If it is 4, it implies it is in the format of ex: OBX-3-1-1
                    if len(mapping_list_components) == 4:
                        self.hl7_dict_values_each_msg[mapping_list_components[0]][int(mapping_list_components[1])][int(mapping_list_components[2])].append(row_value)

                    # Check for the length of the list. If it is 3, it implies it is in the format of ex: OBX-3-1
                    if len(mapping_list_components) == 3:
                        subcomponents_list = row_value.split('&')
                        # Case where OBX-3-1 is in format of 1&2
                        if len(subcomponents_list) > 0:
                            self.hl7_dict_values_each_msg[mapping_list_components[0]][int(mapping_list_components[1])][int(mapping_list_components[2])] = subcomponents_list

                        # Case where OBX-3-1 is in format of 1
                        else:
                            self.hl7_dict_values_each_msg[mapping_list_components[0]][int(mapping_list_components[1])][int(mapping_list_components[2])] = row_value
                          
                
                    # Check for the length of the list. If it is 2, it implies it is in the format of ex: OBX-3
                    # It can be in the format of 1^2 or 1&2^3&4 or just 1, all three cases have to be covered
                    if len(mapping_list_components) == 2:
                        components_list = row_value.split('^')


                        # Case 1: OBX-3 is in format of 1
                        if len(components_list) == 0:
                            self.hl7_dict_values_each_msg[mapping_list_components[0]][int(mapping_list_components[1])][1] = row_value

                        # Case 2: OBX-3 is in format of 1^2 or 1&2^3 or 1&2^3&4
                        if (len(components_list) > 0):
                            for each_comp_index, each_comp in enumerate(components_list):
                                subcomponents = each_comp.split('&')

                                if len(subcomponents) == 0:
                                    self.hl7_dict_values_each_msg[mapping_list_components[0]][int(mapping_list_components[1])][each_comp_index + 1] = each_comp
                                else:
                                    self.hl7_dict_values_each_msg[mapping_list_components[0]][int(mapping_list_components[1])][each_comp_index + 1] = subcomponents

                

                # Traverse through each row of the sheet, obtain the Unique ID (in this case Node ID is used) and append that
                # hl7_dict_values dictionary to the appropriate list associated with that Unique ID
                try:
                    message_id = str(row_entry['NodeID'])
                except KeyError:
                    message_id = str(row_entry['PV1-3'])

                self.hl7_file_message[message_id].append(self.hl7_dict_values_each_msg)

                self.hl7_file_message = OrderedDict(sorted(self.hl7_file_message.items(), key=lambda t: t[0])) 

        # Insert the non-obx variables in the self.hl7_file_message before returning it 
        for caps_var_id, value_to_set, nodeid_loc in self.non_obx_variables_list:
            # Handles the case where the header is in the format of MSH-3-1
            if len(self.header_variables[caps_var_id].split('-')) == 3:
                header, field, comp = self.header_variables[caps_var_id].split('-')
                subcomp_list = self.hl7_file_message[nodeid_loc][0][header][int(field)][int(comp)]
                del subcomp_list[:]
                subcomp_list.append(value_to_set)
                self.hl7_file_message[nodeid_loc][0][header][int(field)][int(comp)] = subcomp_list
            # Handles the case where the header is in the format of PID-8
            else:
                header, field = self.header_variables[caps_var_id].split('-')
                subcomp_list = self.hl7_file_message[nodeid_loc][0][header][int(field)][1]
                print subcomp_list
                if subcomp_list is None:
                    subcomp_list.append(value_to_set)
                else:
                    del subcomp_list[:]
                    subcomp_list.append(value_to_set)
                self.hl7_file_message[nodeid_loc][0][header][int(field)][1] = subcomp_list


        return self.hl7_file_message

        
    def remove_trailing_field_seperators(self, string_to_trim):
        ''' Will trim'''
        # This part of the code will remove any trailing '|'s
        string_temp =  string_to_trim.split('|')
        while string_temp:
            if string_temp[-1] == '':
                string_temp.pop()
            else:
                break

        return self.field_separator.join(string_temp)

    def get_msh_data(self, dictionary_item):
        ''' Gets the MSH segment '''
        msh_comp_list = []
        
        # Key value will be in the format of  ex: 1    OrderedDict([(1, ['1']), (2, defaultdict(<type 'list'>, {}))])
        for key, value in dictionary_item['MSH'].iteritems():
            msh_subcomp_list = []
            #field_keys and field_values will be in the format of 
            # 1 <type 'list'>  and 2 <type 'collections.defaultdict'>
            for field_keys, field_values in value.iteritems():
                if isinstance(field_values, (list,)):
                    if len(field_values) == 1:
                        msh_subcomp_list.extend(field_values)
                    elif len(field_values) > 1:
                        msh_subcomp_list.append(self.subcomponent_separator.join(field_values))
            msh_comp_list.append(self.component_separator.join(msh_subcomp_list))

        msh_result = self.remove_trailing_field_seperators(self.field_separator.join(msh_comp_list))

        return msh_result

        
    def get_pid_data(self, dictionary_item):
        ''' Gets the PID segment '''
        pid_comp_list = []        
        # Key value will be in the format of  ex: 1    OrderedDict([(1, ['1']), (2, defaultdict(<type 'list'>, {}))])
        for key, value in dictionary_item['PID'].iteritems():
            pid_subcomp_list = []
            #field_keys and field_values will be in the format of 
            # 1 <type 'list'>  and 2 <type 'collections.defaultdict'>
            for field_keys, field_values in value.iteritems():
                if isinstance(field_values, (list,)):
                    if len(field_values) == 1:       
                        pid_subcomp_list.extend(field_values)
                    elif len(field_values) > 1:
                        pid_subcomp_list.append(self.subcomponent_separator.join(field_values))
            pid_comp_list.append(self.component_separator.join(pid_subcomp_list))

        pid_result = self.remove_trailing_field_seperators("PID|" + self.field_separator.join(pid_comp_list))

        return pid_result


    def get_pv1_data(self, dictionary_item):
        ''' Gets the PV1 segment '''
        pv1_comp_list = []       
        # Key value will be in the format of  ex: 1    OrderedDict([(1, ['1']), (2, defaultdict(<type 'list'>, {}))])
        for key, value in dictionary_item['PV1'].iteritems():
            pv1_subcomp_list = []
            #field_keys and field_values will be in the format of 
            # 1 <type 'list'>  and 2 <type 'collections.defaultdict'>
            for field_keys, field_values in value.iteritems():
                if isinstance(field_values, (list,)):
                    if len(field_values) == 1:       
                        pv1_subcomp_list.extend(field_values)
                    elif len(field_values) > 1:
                        pv1_subcomp_list.append(self.subcomponent_separator.join(field_values))
            pv1_comp_list.append(self.component_separator.join(pv1_subcomp_list))

        pv1_result = self.remove_trailing_field_seperators("PV1|" + self.field_separator.join(pv1_comp_list))


        return pv1_result



    def get_obr_data(self, dictionary_item):
        ''' Gets the OBR segment '''
        obr_comp_list = []
        
        # Key value will be in the format of  ex: 1    OrderedDict([(1, ['1']), (2, defaultdict(<type 'list'>, {}))])
        for key, value in dictionary_item['OBR'].iteritems():
            obr_subcomp_list = []
            #field_keys and field_values will be in the format of 
            # 1 <type 'list'>  and 2 <type 'collections.defaultdict'>
            for field_keys, field_values in value.iteritems():
                if isinstance(field_values, (list,)):
                    if len(field_values) == 1:       
                        obr_subcomp_list.extend(field_values)
                    elif len(field_values) > 1:
                        obr_subcomp_list.append(self.subcomponent_separator.join(field_values))
            obr_comp_list.append(self.component_separator.join(obr_subcomp_list))
        
        # Setting timestamp from the current timestamp + 1s delay for every OBR segment (OBR-7 is the timestamp field)
        self.current_time = self.current_time + timedelta(seconds = 1)
        obr_comp_list[6] = "%d%d%d%d%d%d" % (self.current_time.year, self.current_time.month, self.current_time.day, self.current_time.hour, self.current_time.minute, self.current_time.second)

        obr_result = self.remove_trailing_field_seperators("OBR|" + self.field_separator.join(obr_comp_list))

        return obr_result

    def get_obx_data(self, dictionary_item):
        ''' Gets the OBX segment '''
        obx_comp_list = []
        
        # Key value will be in the format of  ex: 1    OrderedDict([(1, ['1']), (2, defaultdict(<type 'list'>, {}))])
        for key, value in dictionary_item['OBX'].iteritems():
            obx_subcomp_list = []
            #field_keys and field_values will be in the format of 
            # 1 <type 'list'>  and 2 <type 'collections.defaultdict'>
            for field_keys, field_values in value.iteritems():
                if isinstance(field_values, (list,)):
                    if len(field_values) == 1:       
                        obx_subcomp_list.extend(field_values)
                    elif len(field_values) > 1:
                        obx_subcomp_list.append(self.subcomponent_separator.join(field_values))
            obx_comp_list.append(self.component_separator.join(obx_subcomp_list))

        # Setting timestamp from the current timestamp + 1s delay for every OBX segment (OBX-14 is the timestamp field)
        self.current_time = self.current_time + timedelta(seconds = 1)
        obx_comp_list[13] = "%d%d%d%d%d%d" % (self.current_time.year, self.current_time.month, self.current_time.day, self.current_time.hour, self.current_time.minute, self.current_time.second)

        obx_result = self.remove_trailing_field_seperators("OBX|" + self.field_separator.join(obx_comp_list))


        return obx_result
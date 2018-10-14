#!/usr/bin/python2.7

import sys
from collections import defaultdict
from collections import OrderedDict
import pandas as pd
import xlrd
import copy


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

       


    def set_hl7segments_count(self, segments_and_count_dict_values):
        ''' This method will get the segments' fields, components and subcomponents values and set it for other methods to use'''

        segments_and_counts_dict = OrderedDict(defaultdict(list))

        for segment in segments_and_count_dict_values.keys():
            value = segments_and_count_dict_values[segment]

            # Segment will have the value 'MSH Field Count' and in segment name, we are splitting the word and the first word in the list 
            # will be the segment name in this case 'MSH' and segment_messagwe will have "Field"
            segment_name = (segment.split())[0]
            segment_message = (segment.split())[1]

            
            #  We need to check if this segment exists in the segments_and_counts_dict and if not add it to the dictionary as a key
            if segment_name in segments_and_counts_dict.keys():
                pass
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
                        self.hl7_dict_values[every_key][field_count][component_count][subcomponent_count] = ""

        



    def split_message_box_segments(self, list_of_message_information):
        # Obtain the texts and map them to the respective HL7 Segments
            for each_info in list_of_message_information:
                messagebox_text = each_info.text()

                # This will split the information on '|' and generate a list of the different field values
                field_values = str(messagebox_text).strip().split(self.field_separator)

                #Based on the field values, place the different fields in to the appropriate fields of the dictionary

                #This checks if the field belongs to the MSH Segment
                
                if field_values[0] == "MSH":
                    for index, field_value in enumerate(field_values):
                        try:
                            self.hl7_dict_values['MSH'][index + 1][1] = field_value
                        except KeyError:
                            print "The Field Count exceeds what the default is set to. The default is 17 "


                # This checks if the field belongs to the PID Segment
                # Ex: PID|||000000^^^^AN||UNKNOWN|||||||||||||000000
                           
                if field_values[0] == "PID":
         
                    field_values = field_values[1:]
                    for index, field_value in enumerate(field_values):

                        # Need to check if the fields have multiple components
                        component_values = field_value.split(self.component_separator)

                        # If this list is empty, it means this field doesn't have multiple components
                        if not component_values:
                            self.hl7_dict_values['PID'][index + 1][1] = field_value
                        
                        # If the field has multiple componetns like "000000 & 111111^^^^AN"
                        else:
                            for comp_index, component_value in enumerate(component_values):
                                
                                # Need to check if the components have multiple subcomponents like "000000 & 111111"
                                subcomponent_values = component_value.split(self.subcomponent_separator)

                                # If this list is empty, it means this component doesn't have multiple subcomponents
                                if not subcomponent_values:
                                    self.hl7_dict_values['PID'][index + 1][comp_index + 1] = component_value
                                else:
                                    for subcomp_index, subcomponent_value in enumerate(subcomponent_values):
                                        self.hl7_dict_values['PID'][index + 1][comp_index + 1][subcomp_index + 1] = component_value

                print self.hl7_dict_values['PID']



                # This checks if the field belongs to the PV1 Segment
                # Ex: PV1|||SURG^^OR4
                
                if field_values[0] == "PV1":

                    field_values = field_values[1:]
                    for index, field_value in enumerate(field_values):
                        component_values = field_value.split('^')
                        # If this list is empty, it means this field doesn't have multiple components
                        if not component_values:
                            self.hl7_dict_values['PV1'][index + 1][1] = field_value
                        
                        # If the field has multiple componetns like "000000^^^^AN"
                        else:
                            for comp_index, component_value in enumerate(component_values):
                                self.hl7_dict_values['PV1'][index + 1][comp_index + 1] = component_value



                # This checks if the field belongs to the PV1 Segment
                # Ex: OBR|||||||20150217131700
                
                if field_values[0] == "OBR":
                    field_values = field_values[1:]
                    for index, field_value in enumerate(field_values):
                        component_values = field_value.split('^')
                        # If this list is empty, it means this field doesn't have multiple components
                        if not component_values:
                            self.hl7_dict_values['OBR'][index + 1][1] = field_value
                        
                        # If the field has multiple componetns like "000000^^^^AN"
                        else:
                            for comp_index, component_value in enumerate(component_values):
                                self.hl7_dict_values['OBR'][index + 1][comp_index + 1] = component_value


    def read_excel_data(self, excel_sheet, data_table):
        ''' This method will read data from the excel sheet into a pandas data frame and also read the data from the table on the tool'''


        # This part of the code will work with the QE input sheet, read the excel into a pandas data frame
        # Rename the columns to the appropriate HL7 segments and then generate the HL7 file
        # The Code below ensures that each column data is taken as string instead of doing any conversions to float, int etc.

        df_actual = pd.read_excel(excel_sheet)
        column_list = df_actual.columns.values.tolist()

        converter = {col: str for col in column_list} 
        self.df = pd.read_excel(excel_sheet, converters=converter)



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
            self.df.to_excel("Renamed_QE_sheet_June11.xlsx", index=False)


    def remove_variables_from_list(self, var_list):
        ''' This will remove the variables from the variables_list list from the pandas dataframe'''

        for var_id in var_list:
            self.df = self.df[self.df['Capsule Variable ID'] != str(var_id)]

    def generate_hl7_message_data(self):
        ''' This method will create a doctionary which will have all the data necessary to generate the simulation file'''

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
            self.hl7_dict_values_each_msg = {}
            self.hl7_dict_values_each_msg = copy.deepcopy(self.hl7_dict_values)

            for item in self.mapping_list:
                row_value = row_entry[item]
                if row_value != row_value:
                    row_value = ""

                mapping_list_components =  item.split('-')
                # Check for the length of the list. If it is 3, it implies it is in the format of ex: OBX-3-1
                if len(mapping_list_components) == 3:
                    self.hl7_dict_values_each_msg[mapping_list_components[0]][int(mapping_list_components[1])][int(mapping_list_components[2])] = row_value
                          
                # Check for the length of the list. If it is 2, it implies it is in the format of ex: PV1-3
                elif len(mapping_list_components) == 2:
                    field_list = row_value.split('^')
                    for index in range(len(field_list)):
                        self.hl7_dict_values_each_msg[mapping_list_components[0]][int(mapping_list_components[1])][index +1] = field_list[index]
                        
                 
            # Traverse through each row of the sheet, obtain the Unique ID (in this case Node ID is used) and append that
            # hl7_dict_values dictionary to the appropriate list associated with that Unique ID
            try:
                message_id = str(row_entry['NodeID'])
            except KeyError:
                message_id = str(row_entry['PV1-3'])

            self.hl7_file_message[message_id].append(self.hl7_dict_values_each_msg)

            self.hl7_file_message = OrderedDict(sorted(self.hl7_file_message.items(), key=lambda t: t[0])) 

        return self.hl7_file_message

        


    def get_msh_data(self, dictionary_item):
        ''' Gets the MSH header segment '''
        msh_result_list = []

        for key in dictionary_item['MSH'].keys():
        # If the value of thay key is a dictionary, it implies that the field has multiple components that need to be concatenated with ^
            if isinstance(dictionary_item['MSH'][key], dict):
                hl7_dict_values_tolist = dictionary_item['MSH'][key].values()
                # This part of the code will remove any trailing '^'s
                while hl7_dict_values_tolist:
                    if hl7_dict_values_tolist[-1] == '':
                        hl7_dict_values_tolist.pop()
                    else:
                        break
                msh_result_list.append(self.component_separator.join(hl7_dict_values_tolist))
                    # If the value is not a dictionary, it implies that the field only has one component
            else:
                msh_result_list.append(dictionary_item['MSH'][key])
              
        # Once all the different fields are saved in the list, they are concatenated using the '|' operator
        msh_result = self.field_separator.join(msh_result_list)

        # This part of the code will remove any trailing '|'s
        msh_result_temp =  msh_result.split('|')
        while msh_result_temp:
            if msh_result_temp[-1] == '':
                msh_result_temp.pop()
            else:
                break

        msh_result = self.field_separator.join(msh_result_temp)

        return msh_result

        
    def get_pid_data(self, dictionary_item):
        ''' Gets the PID segment '''
        pid_result_list = []

        for key in dictionary_item['PID'].keys():
        # If the value of thay key is a dictionary, it implies that the field has multiple components that need to be concatenated with ^
            if isinstance(dictionary_item['PID'][key], dict):
                hl7_dict_values_tolist = dictionary_item['PID'][key].values()
                # This part of the code will remove any trailing '^'s
                while hl7_dict_values_tolist:
                    if hl7_dict_values_tolist[-1] == '':
                        hl7_dict_values_tolist.pop()
                    else:
                        break
                pid_result_list.append(self.component_separator.join(hl7_dict_values_tolist))
                    # If the value is not a dictionary, it implies that the field only has one component
            else:
                pid_result_list.append(dictionary_item['PID'][key])
              
        # Once all the different fields are saved in the list, they are concatenated using the '|' operator
        pid_result = "PID|" + self.field_separator.join(pid_result_list)

        # This part of the code will remove any trailing '|'s
        pid_result_temp =  pid_result.split('|')
        while pid_result_temp:
            if pid_result_temp[-1] == '':
                pid_result_temp.pop()
            else:
                break

        pid_result = self.field_separator.join(pid_result_temp)

        return pid_result


    def get_pv1_data(self, dictionary_item):
        ''' Gets the PV1 segment '''
        pv1_result_list = []

        for key in dictionary_item['PV1'].keys():
        # If the value of thay key is a dictionary, it implies that the field has multiple components that need to be concatenated with ^
            if isinstance(dictionary_item['PV1'][key], dict):
                hl7_dict_values_tolist = dictionary_item['PV1'][key].values()
                # This part of the code will remove any trailing '^'s
                while hl7_dict_values_tolist:
                    if hl7_dict_values_tolist[-1] == '':
                        hl7_dict_values_tolist.pop()
                    else:
                        break
                pv1_result_list.append(self.component_separator.join(hl7_dict_values_tolist))
                    # If the value is not a dictionary, it implies that the field only has one component
            else:
                pv1_result_list.append(dictionary_item['PV1'][key])
              
        # Once all the different fields are saved in the list, they are concatenated using the '|' operator
        pv1_result = "PV1|" + self.field_separator.join(pv1_result_list)

        # This part of the code will remove any trailing '|'s
        pv1_result_temp =  pv1_result.split('|')
        while pv1_result_temp:
            if pv1_result_temp[-1] == '':
                pv1_result_temp.pop()
            else:
                break

        pv1_result = self.field_separator.join(pv1_result_temp)

        return pv1_result




    def get_obr_data(self, dictionary_item):
        ''' Gets the OBR segment '''
        obr_result_list = []

        for key in dictionary_item['OBR'].keys():
        # If the value of thay key is a dictionary, it implies that the field has multiple components that need to be concatenated with ^
            if isinstance(dictionary_item['OBR'][key], dict):
                hl7_dict_values_tolist = dictionary_item['OBR'][key].values()
                # This part of the code will remove any trailing '^'s
                while hl7_dict_values_tolist:
                    if hl7_dict_values_tolist[-1] == '':
                        hl7_dict_values_tolist.pop()
                    else:
                        break
                obr_result_list.append(self.component_separator.join(hl7_dict_values_tolist))
                    # If the value is not a dictionary, it implies that the field only has one component
            else:
                obr_result_list.append(dictionary_item['OBR'][key])
              
        # Once all the different fields are saved in the list, they are concatenated using the '|' operator
        obr_result = "OBR|" + self.field_separator.join(obr_result_list)

        # This part of the code will remove any trailing '|'s
        obr_result_temp =  obr_result.split('|')
        while obr_result_temp:
            if obr_result_temp[-1] == '':
                obr_result_temp.pop()
            else:
                break

        obr_result = self.field_separator.join(obr_result_temp)

        return obr_result


    def get_obx_data(self, dictionary_item):
        ''' Gets the OBX segment '''
        obx_result_list = []
        

        for key in dictionary_item['OBX'].keys():
        # If the value of thay key is a dictionary, it implies that the field has multiple components that need to be concatenated with ^
            if isinstance(dictionary_item['OBX'][key], dict):
                hl7_dict_values_tolist = dictionary_item['OBX'][key].values()
                # This part of the code will remove any trailing '^'s
                while hl7_dict_values_tolist:
                    if hl7_dict_values_tolist[-1] == '':
                        hl7_dict_values_tolist.pop()
                    else:
                        break
                obx_result_list.append(self.component_separator.join(hl7_dict_values_tolist))
                    # If the value is not a dictionary, it implies that the field only has one component
            else:
                obx_result_list.append(dictionary_item['OBX'][key])
              


        obx_result_list[0] = str(self.count + 1)
        self.count += 1

        # Once all the different fields are saved in the list, they are concatenated using the '|' operator
        obx_result = "OBX|" + self.field_separator.join(obx_result_list)

        # This part of the code will remove any trailing '|'s
        obx_result_temp =  obx_result.split('|')
        while obx_result_temp:
            if obx_result_temp[-1] == '':
                obx_result_temp.pop()
            else:
                break

        obx_result = self.field_separator.join(obx_result_temp)

        return obx_result





















      





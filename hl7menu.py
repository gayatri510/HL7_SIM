#!/usr/bin/python2.7

import sys
from collections import defaultdict
from collections import OrderedDict




class Hl7_menu():

    def __init__(self):

        # Initializing the class

        # This will create a dictionary with the keys as the different HL7 segments and their values as the fields, components, subcomponents etc
        # To Note: The current implementation only handles fields and components and subcomponents. Field Repetitions are not supported
        self.hl7_dropdown_dict = OrderedDict()

    def create_dropdown_items_from_dict(self, hl7_menu_dict):

        # Iterate through the dictionary, obtain the counts for the fields, components and subcomponents and create the dictionary to populate the
        # dropdown menu appropriately

        # The dictionary is in the following form Ex: ('MSH Field Count', '17')

        segments_and_counts_dict = OrderedDict(defaultdict(list))
     
        for segment in hl7_menu_dict.keys():
            value = hl7_menu_dict[segment]

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

        
        self.hl7_dropdown_dict = OrderedDict()
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

        


   





from collections import OrderedDict

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
obr_7_timestamp = False
obx_14_timestamp = False

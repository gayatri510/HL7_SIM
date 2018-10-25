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


HEADER_VARIABLES = OrderedDict([('6510', 'MSH-3'),
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


default_hl7_segments = "MSH, PID, PV1, OBR"
DEFAULT_CALCULATED_VARIABLES = "1190, 2583, 5815, 5816"
obr_7_timestamp_default_state = True
obx_14_timestamp_default_state = True
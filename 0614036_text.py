# import global_var as gvar
# import Standard_Part
# import csv
#
# with open(gvar.strip_parameters_file_root) as csvFile:
#     rows = csv.reader(csvFile)
#     strip_parameter_list = tuple(tuple(rows)[0])
#     gvar.strip_parameter_list = strip_parameter_list
# gvar.StripDataList[1][1] = 390.0
# gvar.StripDataList[37][1][1] = 2
# gvar.StripDataList[38][1][2] = 4
# gvar.StripDataList[38][1][3] = 1
# gvar.StripDataList[38][1][4] = 1
# gvar.StripDataList[38][1][6] = 4
# gvar.StripDataList[38][1][7] = 1
# gvar.StripDataList[79][1][0] = 4
# gvar.StripDataList[79][1] = 40
# gvar.StripDataList[80][1] = 29.99999999999997
# gvar.StripDataList[81][1] = 420.0
# gvar.StripDataList[82][1] = 75.0
# gvar.StripDataList[83][1] = 234.39999999999998
# gvar.StripDataList[84][1] = 390.0
# gvar.StripDataList[85][1] = 159.39999999999998
# Pin_data = [[0.0, 0.0, 0.0, 0.0, 0.0], [0.0, 10, 10, 10, 0.0], [0.0, 59.400000000000006, 40, 69.96000000000001, 0.0],
#             [0.0, 0.0, 0.0, 0.0, 0.0], [0.0, 0.0, 0.0, 0.0, 0.0]]
# SBT_data = [[0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 16, 0, 0, 0, 0, 0, 0, 0], [0, 60, 0, 0, 0, 0, 0, 0, 0],
#             [0, 9, 0, 0, 0, 0, 0, 0, 0], [0, 24, 0, 0, 0, 0, 0, 0, 0], [0, 14, 0, 0, 0, 0, 0, 0, 0],
#             [0, 12, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 0, 0, 0, 0, 0, 0, 0, 0]]
# Bolt_data = [[0, 0, 0, 0, 0], [0, 13, 13, 13, 0], [0, 15, 15, 15, 0], [0, 90.0, 32.0, 106.0, 0], [0, 19, 19, 19, 0],
#              [0, 150, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0]]
# CB_data = [[0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0],
#            [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0],
#            [0, 18, 12, 13, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0],
#            [0, 0, 0, 0, 0], [0, 0, 0, 0, 0]]
# BoltQuantity = [0, 6, 4, 2, 0]
# PinQuantity = [0, 4, 4, 4, 0, 0, 0, 0, 0]
# Inner_Guiding_data = [[0, 0, 0, 0, 0], [0, 20, 0, 0, 0], [0, 100, 0, 0, 0], [0, 0, 0, 0, 0], [0, 0, 0, 0, 0]]
# InnerGuidingQuantity = [0, 4, 0, 0, 0, 0, 0, 0, 0]
# SBT_CB_data = [[0, 0, 0, 0, 0, 0, 0, 0, 0], [0, 8, 0, 0, 0, 0, 0, 0, 0], [0, 30, 0, 0, 0, 0, 0, 0, 0],
#                [0, 8, 0, 0, 0, 0, 0, 0, 0], [0, 13, 0, 0, 0, 0, 0, 0, 0], [0, 15, 0, 0, 0, 0, 0, 0, 0],
#                [0, 11, 0, 0, 0, 0, 0, 0, 0], [0, 9, 0, 0, 0, 0, 0, 0, 0], [0, 245, 0, 0, 0, 0, 0, 0, 0]]
# SBTQuantity = [0, 2, 0, 0, 0, 0, 0, 0, 0]
#
# (Pin_data, SBT_data, outer_Guiding_data, stripper_pin_data) = Standard_Part.Standard_Part(Bolt_data, CB_data,
#                                                                                           Pin_data,
#                                                                                           Inner_Guiding_data,
#                                                                                           SBT_data, SBT_CB_data)

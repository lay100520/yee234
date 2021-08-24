import win32com.client as win32
import defs
import global_var as gvar
import time

def PunchMaking(now_plate_line_number):
    total_op_number = int(gvar.strip_parameter_list[2])
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    g = now_plate_line_number
    line_name = [""] * 5
    creat_point_name = [""] * 4
    X_direction = [0] * 4
    Y_direction = [0] * 4
    Z_direction = [0] * 4
    first_direction = [0] * 4
    second_direction = [0] * 4
    cut_punch_height = float()
    for now_op_number in range(1, 1 + total_op_number):
        n = now_op_number
        op_number = 10 * n
        # --------------------------------------------------------------------------------------------補強沖頭
        if gvar.StripDataList[3][g][n] > 0:
            # for for_counter in range( 1, 1+ gvar.StripDataList[3][g][n]):
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # file_name = "QR_punch_Reinforcement"
            # punch_File_change
            # punch_Reinforcement_change(for_counter)
            pass  # 未使用
        # --------------------------------------------------------------------------------------------下料切斷沖頭_下
        if gvar.StripDataList[27][g][n] > 0:  # --------------------切斷沖頭_下
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # punch_d_cutting
            # punch_d_cutting_change
            pass  # 未使用
        # --------------------------------------------------------------------------------------------下料切斷沖頭_上
        if gvar.StripDataList[28][g][n] > 0:
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # punch_u_cutting
            # punch_u_cutting_change
            pass  # 未使用
        # --------------------------------------------------------------------------------------------沖切沖頭
        if gvar.StripDataList[38][g][n] > 0:
            for now_data_number in range(1, 1 + gvar.StripDataList[38][g][n]):
                partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
                interferance_pad_name = "_cut_punch_"
                interferance_line_name = "_cut_line_"
                open_file_name = "cut_punch"
                punch(open_file_name)
                punch_change(now_plate_line_number, now_op_number, interferance_line_name, interferance_pad_name,
                             cut_punch_height, now_data_number)
        # --------------------------------------------------------------------------------------------T形異形沖
        if gvar.StripDataList[53][g][n] > 0:
            # for now_unnomal_cut_line_T_number in range( 1, 1+ gvar.StripDataList[53][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     line_name[2] = "plate_line_" + str(g) + "_op" + str(
            #     op_number) + "_unnomal_cut_line_T_symmetric_" + str(now_unnomal_cut_line_T_number)
            #     line_name[3] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_T_" + str(now_unnomal_cut_line_T_number)
            #     # ------------------------------------------------------------↓   判斷是否對稱名稱
            #     selection1 = partDocument1.Selection
            #     selection1.Clear
            #     selection1.Search
            #     "Name=" + line_name[2]
            #     # ------------------------------------------------------------↑
            #     # ------------------------------------------------------------↓   判斷是否對稱名稱
            #     if selection1.Count > 0:
            #         line_name[1] = line_name[2]
            #     else:
            #         line_name[1] = line_name[3]
            #     # ------------------------------------------------------------↑
            #     cut_line_st = line_name[1]
            #     punch_unnomal_cut_line_T
            pass  # 未使用
        # --------------------------------------------------------------------------------------------I形異形沖
        if gvar.StripDataList[54][g][n] > 0:
            # for now_unnomal_cut_line_I_number in range( 1, 1+ gvar.StripDataList[54][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     line_name[2] = "plate_line_" + str(g) + "_op" + str(
            #     op_number) + "_unnomal_cut_line_I_symmetric_" + str(now_unnomal_cut_line_I_number)
            #     line_name[3] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_I_" + str(now_unnomal_cut_line_I_number)
            #     # ------------------------------------------------------------↓   判斷是否對稱名稱
            #     selection1 = partDocument1.Selection
            #     selection1.Clear
            #     selection1.Search
            #     "Name=" + line_name[2]
            #     # ------------------------------------------------------------↑
            #     # ------------------------------------------------------------↓   判斷是否對稱名稱
            #     if selection1.Count > 0:
            #         line_name[1] = line_name[2]
            #     else:
            #         line_name[1] = line_name[3]
            #     # ------------------------------------------------------------↑
            #     cut_line_st = line_name[1]
            #     punch_unnomal_cut_line_I
            pass  # 未使用
        # --------------------------------------------------------------------------------------------M形異形沖
        if gvar.StripDataList[55][g][n] > 0:
            # for now_unnomal_cut_line_M_number in range( 1, 1+ gvar.StripDataList[55][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     line_name[2] = "plate_line_" + str(g) + "_op" + str(
            #     op_number) + "_unnomal_cut_line_M_symmetric_" + str(now_unnomal_cut_line_M_number)
            #     line_name[3] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_M_" + str(now_unnomal_cut_line_M_number)
            #     # ------------------------------------------------------------↓   判斷是否對稱名稱
            #     selection1 = partDocument1.Selection
            #     selection1.Clear
            #     selection1.Search
            #     "Name=" + line_name[2]
            #     # ------------------------------------------------------------↑
            #     # ------------------------------------------------------------↓   判斷是否對稱名稱
            #     if selection1.Count > 0:
            #         line_name[1] = line_name[2]
            #     else:
            #         line_name[1] = line_name[3]
            #     # ------------------------------------------------------------↑
            #     cut_line_st = line_name[1]
            #     punch_unnomal_cut_line_M
            pass  # 未使用
        # --------------------------------------------------------------------------------------------成形沖頭,是沖頭
        if gvar.StripDataList[41][g][n] > 0:
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # forming_cavity
            # forming_cavity_change
            pass  # 未使用
        # --------------------------------------------------------------------------------------------成形沖頭,是模穴
        if gvar.StripDataList[40][g][n] > 0:
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # forming_punch
            # forming_punch_change
            pass  # 未使用
        # --------------------------------------------------------------------------------------------異型沖頭
        if gvar.StripDataList[39][g][n] > 0:
            # for now_data_number in range(1, 1+ gvar.StripDataList[39][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     interferance_pad_name = "_allotype_cut_punch_"
            #     interferance_line_name = "_allotype_cut_line_"
            #     open_file_name = "cut_punch"
            #     punch(open_file_name)
            #     punch_change(interferance_line_name, interferance_pad_name)
            pass  # 未使用
        # --------------------------------------------------------------------------------------------↓快拆沖頭
        if gvar.StripDataList[29][g][n] > 0:  # 沖切沖頭_右
            # for now_data_number in  range(1, 1+ gvar.StripDataList[29][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     data_type = "line"
            #     part_type = "cut_punch_right"
            #     part_name = part_type
            #     product_name = "right_quickly_remove_cut_punch"
            #     modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_right_quickly_remove_cut_line_"
            #     creat_point_name[1] = "right_quickly_remove_cut_line_Ymax_point"
            #     creat_point_name[2] = "right_quickly_remove_cut_line_Ymin_point"
            #     X_direction[1] = 0
            #     Y_direction[1] = 1
            #     Z_direction[1] = 0
            #     first_direction[1] = 1
            #     first_direction[2] = 0
            #     X_direction[2] = 1
            #     Y_direction[2] = 0
            #     Z_direction[2] = 0
            #     second_direction[1] = 1
            #     second_direction[2] = 1
            #     quickly_remove_punch
            #     quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[30][g][n] > 0:  # 沖切沖頭_左
            # for now_data_number in  range(1, 1+ gvar.StripDataList[30][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     data_type = "line"
            #     part_type = "cut_punch_left"
            #     part_name = part_type
            #     product_name = "left_quickly_remove_cut_punch"
            #     modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_left_quickly_remove_cut_line_"
            #     creat_point_name[1] = "left_quickly_remove_cut_line_Ymax_point"
            #     creat_point_name[2] = "left_quickly_remove_cut_line_Ymin_point"
            #     X_direction[1] = 0
            #     Y_direction[1] = 1
            #     Z_direction[1] = 0
            #     first_direction[1] = 1
            #     first_direction[2] = 0
            #     X_direction[2] = 1
            #     Y_direction[2] = 0
            #     Z_direction[2] = 0
            #     second_direction[1] = 1
            #     second_direction[2] = 1
            #     quickly_remove_punch
            #     quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[31][g][n] > 0:  # 沖切沖頭_上
            # for now_data_number in  range(1, 1+ gvar.StripDataList[31][g][n]):
            #         partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #         data_type = "line"
            #         part_type = "cut_punch_up"
            #         part_name = part_type
            #         product_name = "up_quickly_remove_cut_punch"
            #         modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_up_quickly_remove_cut_line_"
            #         creat_point_name[1] = "up_quickly_remove_cut_line_Xmax_point"
            #         creat_point_name[2] = "up_quickly_remove_cut_line_Xmin_point"
            #         X_direction[1] = 1
            #         Y_direction[1] = 0
            #         Z_direction[1] = 0
            #         first_direction[1] = 1
            #         first_direction[2] = 0
            #         X_direction[2] = 0
            #         Y_direction[2] = 1
            #         Z_direction[2] = 0
            #         second_direction[1] = 1
            #         second_direction[2] = 1
            #         quickly_remove_punch
            #         quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[32][g][n] > 0:  # 沖切沖頭_下
            # for now_data_number in  range(1, 1+ gvar.StripDataList[32][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     data_type = "line"
            #     part_type = "cut_punch_down"
            #     part_name = part_type
            #     product_name = "down_quickly_remove_cut_punch"
            #     modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_down_quickly_remove_cut_line_"
            #     creat_point_name[1] = "down_quickly_remove_cut_line_Xmax_point"
            #     creat_point_name[2] = "down_quickly_remove_cut_line_Xmin_point"
            #     X_direction[1] = 1
            #     Y_direction[1] = 0
            #     Z_direction[1] = 0
            #     first_direction[1] = 1
            #     first_direction[2] = 0
            #     X_direction[2] = 0
            #     Y_direction[2] = 1
            #     Z_direction[2] = 0
            #     second_direction[1] = 1
            #     second_direction[2] = 1
            #     quickly_remove_punch
            #     quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[33][g][n] > 0:  # 折彎沖頭_右
            # for now_data_number in  range(1, 1+ gvar.StripDataList[33][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     data_type = "surface"
            #     part_type = "bending_punch_right"
            #     part_name = part_type
            #     product_name = "right_quickly_remove_bending_punch"
            #     modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_right_quickly_remove_bending_surface_"
            #     item_sketch = "formula_Sketch"
            #     creat_point_name[1] = "right_quickly_remove_bending_surface_Ymax_point"
            #     creat_point_name[2] = "right_quickly_remove_bending_surface_Ymin_point"
            #     X_direction[1] = 0
            #     Y_direction[1] = 1
            #     Z_direction[1] = 0
            #     first_direction[1] = 1
            #     first_direction[2] = 0
            #     X_direction[2] = 1
            #     Y_direction[2] = 0
            #     Z_direction[2] = 0
            #     second_direction[1] = 1
            #     second_direction[2] = 1
            #     quickly_remove_punch
            #     quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[34][g][n] > 0:  # 折彎沖頭_左
            # for now_data_number in  range(1, 1+ gvar.StripDataList[34][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     data_type = "surface"
            #     part_type = "bending_punch_left"
            #     part_name = part_type
            #     product_name = "left_quickly_remove_bending_punch"
            #     modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_left_quickly_remove_bending_surface_"
            #     item_sketch = "formula_Sketch"
            #     creat_point_name[1] = "left_quickly_remove_bending_surface_Ymax_point"
            #     creat_point_name[2] = "left_quickly_remove_bending_surface_Ymin_point"
            #     X_direction[1] = 0
            #     Y_direction[1] = 1
            #     Z_direction[1] = 0
            #     first_direction[1] = 1
            #     first_direction[2] = 0
            #     X_direction[2] = 1
            #     Y_direction[2] = 0
            #     Z_direction[2] = 0
            #     second_direction[1] = 1
            #     second_direction[2] = 1
            #     quickly_remove_punch
            #     quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[35][g][n] > 0:  # 折彎沖頭_上
            # for now_data_number in  range(1, 1+ gvar.StripDataList[35][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     data_type = "surface"
            #     part_type = "bending_punch_up"
            #     part_name = part_type
            #     product_name = "up_quickly_remove_bending_punch"
            #     modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_up_quickly_remove_bending_surface_"
            #     item_sketch = "formula_Sketch"
            #     creat_point_name[1] = "up_quickly_remove_bending_surface_Xmax_point"
            #     creat_point_name[2] = "up_quickly_remove_bending_surface_Xmin_point"
            #     X_direction[1] = 1
            #     Y_direction[1] = 0
            #     Z_direction[1] = 0
            #     first_direction[1] = 1
            #     first_direction[2] = 0
            #     X_direction[2] = 0
            #     Y_direction[2] = 1
            #     Z_direction[2] = 0
            #     second_direction[1] = 1
            #     second_direction[2] = 1
            #     quickly_remove_punch
            #     quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[36][g][n] > 0:  # 折彎沖頭_下
            # for now_data_number in  range(1, 1+ gvar.StripDataList[36][g][n]):
            #     partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            #     data_type = "surface"
            #     part_type = "bending_punch_down"
            #     part_name = part_type
            #     product_name = "down_quickly_remove_bending_punch"
            #     modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_down_quickly_remove_bending_surface_"
            #     item_sketch = "formula_Sketch"
            #     creat_point_name[1] = "down_quickly_remove_bending_surface_Xmax_point"
            #     creat_point_name[2] = "down_quickly_remove_bending_surface_Xmin_point"
            #     X_direction[1] = 1
            #     Y_direction[1] = 0
            #     Z_direction[1] = 0
            #     first_direction[1] = 1
            #     first_direction[2] = 0
            #     X_direction[2] = 0
            #     Y_direction[2] = 1
            #     Z_direction[2] = 0
            #     second_direction[1] = 1
            #     second_direction[2] = 1
            #     quickly_remove_punch
            #     quickly_remove_punch_change
            pass  # 未使用
        # --------------------------------------------------------------------------------------------↑快拆沖頭
        # --------------------------------------------------------------------------------------------打凸包沖頭_左
        if gvar.StripDataList[21][g][n] > 0:
            # emboss_forming_punch_direction = "left"
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # emboss_forming_punch_left
            # emboss_forming_punch_left_change
            pass  # 未使用
        # --------------------------------------------------------------------------------------------打凸包沖頭_右
        if gvar.StripDataList[22][g][n] > 0:
            # emboss_forming_punch_direction = "right"
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # emboss_forming_punch_right
            # emboss_forming_punch_right_change
            pass  # 未使用
        # --------------------------------------------------------------------------------------------半沖切沖頭
        if gvar.StripDataList[3][g][n] > 0:
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # half_cut_punch
            # half_cut_punch_change
            pass  # 未使用
        # --------------------------------------------------------------------------------------------整形沖頭
        if gvar.StripDataList[73][g][n] > 0:
            # for now_data_number = 1, 1+ plate_line_bending_punch_surface(g, n)
            #     F_bending_punch
            pass  # 未使用


def punch(open_file_name):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    partDocument2 = documents1.Open(gvar.open_path + open_file_name + ".CATPart")
    time.sleep(1)
    # 在catapp上切換各視窗
    # ======================================
    defs.window_change(partDocument1, partDocument2)
    # ======================================


def punch_change(now_plate_line_number, now_op_number, interferance_line_name, interferance_pad_name, cut_punch_height,
                 now_data_number):
    total_op_number = int(gvar.strip_parameter_list[2])
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    # ---------使用迴圈，建立關連↓-------------
    # ---------決定沖頭高度↓-------------
    length = [None] * 2
    op_number = 10 * now_op_number
    Thickness = float(gvar.strip_parameter_list[1])
    # ======================================================================================================
    length[0] = part1.Parameters.Item("cut_punch_up_plane")
    g = now_plate_line_number
    if gvar.Mold_status == "開模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14]) + 0  # (upper_die_open_height)
    if gvar.Mold_status == "閉模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14])
    # ======================================================================================================
    length[1] = part1.Parameters.Item("cut_punch_height")
    n = now_op_number
    if cut_punch_height != 0.0:
        length[1].Value = cut_punch_height
    else:
        die_rule_file_name = "沖頭切入深度"
        Row_string_serch = "精密級"  # ---------X
        Column_string_serch = "合金工具鋼"  # ---------------Y
        if Thickness >= 0 and Thickness < 0.1:
            excel_Sheet_name = "0.1以下"
        elif Thickness >= 0.1 and Thickness < 0.25:
            excel_Sheet_name = "0.1~0.25"
        elif Thickness >= 0.25 and Thickness < 0.5:
            excel_Sheet_name = "0.25~0.5"
        elif Thickness >= 0.5 and Thickness < 0.8:
            excel_Sheet_name = "0.5~0.8"
        elif Thickness >= 0.8 and Thickness < 1.2:
            excel_Sheet_name = "0.8~1.2"
        elif Thickness >= 1.2 and Thickness < 1.6:
            excel_Sheet_name = "1.2~1.6"
        elif Thickness >= 1.6 and Thickness < 2.5:
            excel_Sheet_name = "1.6~2.5"
        elif Thickness >= 2.5 and Thickness < 3.5:
            excel_Sheet_name = "2.5~3.5"
        else:
            excel_Sheet_name = "3.5以上"
        (serch_result) = defs.ExcelSearch(die_rule_file_name, excel_Sheet_name, Row_string_serch, Column_string_serch)
        length[1].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14]) + float(serch_result)  # 沖頭高度
    # ---------決定沖頭高度↑-------------
    part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
        "die\\plate_line_" + str(g) + "_op" + str(op_number) + interferance_line_name + str(now_data_number))  # 草圖置換
    time.sleep(0.5)
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    product1 = partDocument1.getItem("Part1")
    # 數字二位數化,1~10改為01~10
    X = 0  # 名稱命名
    if now_data_number >= 10:
        X = ""  # 名稱命名
    product1.PartNumber = "op" + str(op_number) + interferance_pad_name + str(X) + str(now_data_number)  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("op" + str(op_number) + interferance_pad_name + str(X) + str(now_data_number))
    parameters1 = product1.UserRefProperties
    strParam1 = parameters1.CreateString("NO.", "")
    strParam1.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters2 = product1.UserRefProperties
    strParam2 = parameters2.CreateString("Part Name", "")
    strParam2.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters3 = product1.UserRefProperties
    strParam3 = parameters3.CreateString("Size", "")
    strParam3.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters4 = product1.UserRefProperties
    strParam4 = parameters4.CreateString("Material_Data", "")
    strParam4.ValuateFromString(gvar.strip_parameter_list[37])
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam5 = parameters5.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(gvar.strip_parameter_list[38])
    product1 = product1.ReferenceProduct
    parameters6 = product1.UserRefProperties
    strParam6 = parameters6.CreateString("Quantity", "")
    strParam6.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters7 = product1.UserRefProperties
    strParam7 = parameters7.CreateString("Page", "")
    strParam7.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters8 = product1.UserRefProperties
    strParam8 = parameters8.CreateString("L1", "")  # 形狀孔
    strParam8.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters9 = product1.UserRefProperties
    strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
    strParam9.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters14 = product1.UserRefProperties
    strParam12 = parameters14.CreateString("HP", "")  # 合銷孔
    strParam12.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters15 = product1.UserRefProperties
    strParam13 = parameters15.CreateString("B", "")  # B型引導沖孔
    strParam13.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters16 = product1.UserRefProperties
    strParam14 = parameters16.CreateString("BP", "")  # B沖沖孔
    strParam14.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters17 = product1.UserRefProperties
    strParam15 = parameters17.CreateString("TS", "")  # 浮升引導
    strParam15.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters18 = product1.UserRefProperties
    strParam16 = parameters18.CreateString("IG", "")  # 內導柱
    strParam16.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters19 = product1.UserRefProperties
    strParam17 = parameters19.CreateString("F", "")  # 外導柱
    strParam17.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters20 = product1.UserRefProperties
    strParam18 = parameters20.CreateString("CS", "")  # 等高套筒
    strParam18.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters21 = product1.UserRefProperties
    strParam19 = parameters21.CreateString("AP", "")  # A沖沖孔
    strParam19.ValuateFromString("")
    product1 = product1.ReferenceProduct
    # ====↑設定性質↑=====================================
    part1.Update()
    # --------------↓刪除不需要的Data↓--------------
    selection1 = partDocument1.Selection
    if now_plate_line_number == 1:  # --------模板1
        selection1.Clear()
        selection1.Search("Name=plate_line_2*,all")
        if selection1.Count != 0:
            selection1.Delete()
            selection1.Clear()
        for o in range(1, 1 + (now_op_number - 1)):
            selection1.Clear()
            selection1.Search("Name=*_op" + str(o) + "0_*,all")
            if selection1.Count != 0:
                selection1.Delete()
            selection1.Clear()
        for o in range((now_op_number + 1), 1 + total_op_number):
            selection1.Clear()
            selection1.Search("Name=*_op" + str(o) + "0_*,all")
            if selection1.Count != 0:
                selection1.Delete()
            selection1.Clear()
    if now_plate_line_number == 2:  # --------模板2
        selection1.Clear()
        selection1 = partDocument1.Selection
        selection1.Search("Name=plate_line_1*,all")
        if selection1.Count != 0:
            selection1.Delete()
        selection1.Clear()
        for o in range(1, 1 + (now_op_number - 1)):
            selection1.Clear()
        selection1.Search("Name=*_op" + str(o) + "0_*,all")
        if selection1.Count != 0:
            selection1.Delete()
        selection1.Clear()
        for o in range((now_op_number + 1), 1 + total_op_number):
            selection1.Clear()
        selection1.Search("Name=*_op" + str(o) + "0_*,all")
        if selection1.Count != 0:
            selection1.Delete()
        selection1.Clear()
    # --------------↑刪除不需要的Data↑--------------
    part1.Update()
    time.sleep(2)
    partDocument1.SaveAs(gvar.save_path + "op" + str(op_number) + interferance_pad_name + str(X) + str(
        now_data_number) + ".CATPart")  # 存檔的檔案名稱
    time.sleep(1)
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    # ---------使用迴圈，建立關連↑-------------
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    partDocument1.Close()

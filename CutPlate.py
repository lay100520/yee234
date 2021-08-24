import win32com.client as win32
import defs
import global_var as gvar
import PunchDef
import time
import math


def CutPlate(now_plate_line_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    partDocument2 = documents1.Open(gvar.open_path + "cut_plate.CATPart")
    # ======================================
    defs.window_change(partDocument1, partDocument2)  # 在CATIA上切換各視窗
    # ======================================
    documents1 = catapp.Documents
    partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert.CATPart")
    # ======================================
    defs.window_change(partDocument1, partDocument2)  # 在CATIA上切換各視窗
    # ======================================
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.3")
    body1.Name = "cut_cavity_insert"
    selection3 = partDocument1.Selection
    # ------------------------------------------------------------↓   判斷是否有沖切模組
    selection3.Clear()
    selection3.Search("Name=plate_line_" + str(now_plate_line_number) + "_op*_cut_punch_*_cutting_*,all")
    if selection3.Count > 0:
        partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_shear.CATPart")
        # ======================================
        defs.window_change(partDocument1, partDocument2)  # 在CATIA上切換各視窗
        # ======================================
        bodies1 = part1.Bodies
        # ------------------------------------------------------------↓   改變body名稱
        body1 = bodies1.Item("Body.4")
        body1.Name = "cut_cavity_insert_shear"
        # ------------------------------------------------------------↑
        selection3.Clear()
        selection3.Search("Name=*cut_cavity_insert_shear,all")
        selection3.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
        selection3.Clear()
    # ============================================================================(cut_plate_change)
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    length = [None] * 50
    formula = [None] * 20
    parameter = [None] * 20
    pad_Bolt_Hole = [0] * 20
    offest_invert = int()
    line_name = [""] * 5
    lower_pad_insert_hole_number = int()
    cut_plat_machining_explanation_pilot_punch = int()
    Working_parameter = int()
    total_op_number = int(gvar.strip_parameter_list[2])
    # ======================================================================================================
    length[0] = part1.Parameters.Item("pitch")
    length[0].Value = gvar.strip_parameter_list[4]
    # ======================================================================================================
    # ======================================================================================================
    length[1] = part1.Parameters.Item("plate_height")
    g = now_plate_line_number
    if gvar.die_type == "module":
        plate_H = float(gvar.strip_parameter_list[1])
    else:
        plate_H = -float(gvar.strip_parameter_list[29])
    if gvar.strip_parameter_list[29] != "":
        if gvar.die_type == "module":
            length[1].Value = -float(gvar.strip_parameter_list[26])
        else:
            length[1].Value = -float(gvar.strip_parameter_list[29])
    else:
        die_rule_file_name = "模板厚度選擇"
        excel_Sheet_name = "200T以下"
        Column_string_serch = "下墊板"
        Row_string_serch = "精密級"
        (serch_result) = defs.ExcelSearch(die_rule_file_name, excel_Sheet_name, Row_string_serch, Column_string_serch)
        length[1].Value = -serch_result
        gvar.strip_parameter_list[29] = length[1]
    # ======================================================================================================
    length[4] = part1.Parameters.Item("plate_up_plane")
    if gvar.die_type == "module":
        plate_position = 0  # (-die_open_height)
    else:
        plate_position = -0 - int(gvar.strip_parameter_list[26])  # die_open_height
    length[4].Value = plate_position
    # ======================================================================================================
    length[6] = part1.Parameters.Item("offset")
    if offest_invert == 1:
        length[6].Value = -1.5  # (-down_plate_hole_offset)
    else:
        length[6].Value = 1.5  # (down_plate_hole_offset)
    # ======================================================================================================
    file_name = "Data1"
    body_name1 = "Body.2"
    hybridBody_name = "die"
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(file_name, body_name1,
                                                                                             hybridBody_name)  # 環境設定(檔案名,body名,依據群組名)(全域變數改)
    if gvar.die_type == "module":
        M_plate_length = 35
        M_plate_wide = 180
        hybridShape1 = ElementBody.HybridShapes.Item("up_plane")  # 宣告平面(下)
        defs.material_tpye_palte_sketch(hybridShape1, M_plate_length, M_plate_wide,
                                        100 + 50 * (now_plate_line_number - 1),
                                        112.5 + 12.5, ElementDocument, ElementBody,
                                        ElementHybridBody)  # (平面的宣告,模板長,模板寬,中心座標_X,中心座標_Y)
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\plate_size")  # 草圖置換
        part1.Update()
        # M_cut_plate_design(M_plate_length, M_plate_wide)
        pass  # 未使用
    else:
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\number_" + str(g) + "_plate_line")  # 草圖置換
    part1.Update()
    specsAndGeomWindow1 = catapp.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    part1.Update()
    # =================================================================================================
    hybridShapeFactory1 = part1.HybridShapeFactory
    parameters1 = part1.Parameters
    for n in range(1, 1 + total_op_number):
        op_number = 10 * n
        xi = 0  # 重設xi
        xi_over = False
        if gvar.StripDataList[38][g][n] > 0:  # 普通沖孔沖頭
            for i in range(1, 1 + gvar.StripDataList[38][g][n]):
                xi = gvar.StripDataList[38][g][n]  # xi=當前工站入子數量
                # ------------------------------------------------------------↓   搜尋是否為對稱名稱
                selection3 = partDocument1.Selection
                selection3.Clear()
                selection3.Search(
                    "Name=plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_symmetric_" + str(i))
                # ------------------------------------------------------------↑
                symmetric_amuantity = selection3.Count  # 判斷對稱及記數
                # ------------------------------------------------------------↓   判斷是否對稱名稱
                if symmetric_amuantity > 0:
                    part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                        "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_symmetric_" + str(i))  # 草圖置換
                else:
                    part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                        "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i))  # 草圖置換
                # ------------------------------------------------------------↑
                # ------------------------------------------------------------↓
                if xi > 1:  # xi>1時預先記錄螺栓數量
                    # ------------------------------------------↓搜尋現有pad_Bolt_Hole點的螺栓數量
                    selection1.Clear()
                    selection1.Search("Name=pad_" + str(g) + "_Bolt_point_*,all")
                    pad_Bolt_Hole[g] = selection1.Count
                    selection1.Clear()
                    # ------------------------------------------↑搜尋現有pad_Bolt_Hole點的螺栓數量
                    PunchDef.cut_plate_clash_change(xi, op_number)
                    i = i + (xi - 1)  # 強迫i跳離迴圈
                    xi_over = True
                # ------------------------------------------------------------↑
                # ------------------------------------------------------------↓   參數宣告
                length[13] = part1.Parameters.Item("insert_line")  # 多出來的
                length[13].Value = 5  # 多出來的
                length[33] = part1.Parameters.Item("gap")
                length[33].Value = float(gvar.strip_parameter_list[1]) * 0.02  # (lower_die_space / 100 )  # 間隙
                length[30] = part1.Parameters.Item("cut_cavity_insert_height")
                length[30].Value = float(gvar.strip_parameter_list[26])
                length[31] = part1.Parameters.Item("die_open_height")
                length[31].Value = 0  # (die_open_height)
                length[34] = part1.Parameters.Item("x_to_x")
                length[35] = part1.Parameters.Item("y_to_y")
                length[36] = part1.Parameters.Item("int_x")
                length[37] = part1.Parameters.Item("int_y")
                # ------------------------------------------------------------↑
                # ------------------------------------------------------------↓   先放大入子以免更換時  孔>入子塊
                length[36].Value = 200
                length[37].Value = 200
                # ------------------------------------------------------------↑
                part1.Update()
                # ------------------------------------------------------------↓  修正為對稱孔位之參數 True=對稱  False=不對稱
                boolParam1 = part1.Parameters.Item("symmetry_switch")
                if symmetric_amuantity > 0:
                    boolParam1.Value = True
                else:
                    boolParam1.Value = False
                # ------------------------------------------------------------↑
                # ------------------------------------------------------------↓   整數化
                length[36].Value = math.ceil(length[34].Value)
                length[37].Value = math.ceil(length[35].Value)
                # ------------------------------------------------------------↑
                body_name1 = "cut_cavity_insert"
                sketch_name1 = "cut_insert_line_Sketch"
                insert_line_name = [""] * 4
                insert_line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i)
                if xi <= 1:
                    insert_Bolt_Hole(g, n, i, insert_line_name, sketch_name1, body_name1, type_name=None)  # 挖孔指令
                else:
                    auto_insert_Bolt_Hole(g, n, i, insert_line_name, sketch_name1, body_name1, type_name=None)  # 挖孔指令
                part1.Update()
                if xi_over == True:
                    xi_over = False
                    break
        # ------------------------------------------------------------------------------------------------- ↓補強型沖頭之入子
        if gvar.StripDataList[4][g][n] > 0:
            # for for_counter in range (1, 1 + gvar.StripDataList[4][g][n]):
            #     line_name1 = "plate_line_" + str(g) + "_op" + str(n * 10) + "_Reinforcement_cut_line_" + str(for_counter)
            #     insert_Sketch_Bolt_Hole(g, n, for_counter)
            pass  # 未使用
        # ------------------------------------------------------------------------------------------------- ↑補強型沖頭之入子
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↓I T M 對稱入子 2015-12-25
        if gvar.StripDataList[53][g][n] > 0:
            # for for_counter in range (1, 1 + gvar.StripDataList[53][g][n]):
            #     line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_T_symmetric_" + str(for_counter)
            #     line_name[2] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_T_" + str(for_counter)
            #     insert_Sketch_Bolt_Hole(g, n, for_counter)
            #     symmetric_insert_name_T[n][for_counter] = line_name[1]
            pass  # 未使用

        if gvar.StripDataList[54][g][n] > 0:
            # for for_counter in range (1, 1 + gvar.StripDataList[54][g][n]):
            #    line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_I_symmetric_" + str(for_counter)
            #    line_name[2] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_I_" + str(for_counter)
            #    insert_Sketch_Bolt_Hole(g, n, for_counter)
            #    symmetric_insert_name_I[n][for_counter] = line_name[1]
            pass  # 未使用
        if gvar.StripDataList[55][g][n] > 0:
            # for for_counter in range(1, 1 + gvar.StripDataList[55][g][n]):
            #     line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_M_symmetric_" + str(for_counter)
            #     line_name[2] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_M_" + str(for_counter)
            #     Call
            #     insert_Sketch_Bolt_Hole(g, n, for_counter)
            #     symmetric_insert_name_M(n, for_counter) = line_name(1)
            pass  # 未使用
        # ----------bend_up_shaping_cavity_down_number(g,n)>0↓-----------------------------整平模組孔down
        if gvar.StripDataList[44][g][n] > 0:
            # for i in range (1 , 1+ gvar.StripDataList[44][g][n]):
            #     pp_count = i
            #     Call bend_up_shaping_punch_hole_1
            pass  # 未使用
        # ----------bend_up_shaping_cavity_down_number(g,n)>0↑-----------------------------整平模組孔down
        # ----------bend_up_shaping_cavity_up_number(g,n)>0↓-----------------------------整平模組孔up
        if gvar.StripDataList[45][g][n] > 0:
            # for i in range (1 , 1+ gvar.StripDataList[45][g][n]):
            #     pp_count = i
            #     Call bend_up_shaping_punch_hole_2
            pass  # 未使用
        # ----------bend_up_shaping_cavity_up_number(g,n)>0↑-----------------------------整平模組孔up
        if gvar.StripDataList[27][g][n] > 0:  # --------------------切斷沖頭_下
            # for cut_punch_d_number in range(1, 1 + gvar.StripDataList[27][g][n]):
            #     #----↓參數名稱
            #     point_name[2] = "op" + str(op_number) + "_d_cut_X_min_point"   #點
            #     point_name[4] = "op"+str(op_number) + "_d_cut_X_max_point"
            #     point_name[6] = "op" +str(op_number) +  "_d_cut_Z_point"
            #     body_name[1] = "cut_cavity_insert_shear"    #body
            #     sketch_name[1] = "sketch_Coordinate"        #草圖
            #     type_name = "B_down"                        #型式
            #     insert_line_name[1] = "plate_line_" + str(g) + "_op" +str(op_number) +  "_cut_punch_d_cutting_" +str(cut_punch_d_number)
            #     insert_line_name[2] = "plate_line_" + str(g) + "_op" +str(op_number) +  "_cut_punch_d_insert_surface_" +str( cut_punch_d_number)
            #     #------------------------------------------------------------↑
            #     punch_cutting_change           #換草圖
            #     punch_cutting                  #挖槽
            #     sketch_name[1] = "down_cut_insert_Sketch"
            #     punch_cutting                  #挖槽
            #     type_name = "C_left"
            #     punch_cutting_change           #換草圖
            #     sketch_name[1] = "left_cut_insert_Sketch"
            #     punch_cutting                  #挖槽
            pass  # 未使用
        if gvar.StripDataList[28][g][n] > 0:  # --------------------切斷沖頭_上
            # for cut_punch_u_number in range(1, 1 + gvar.StripDataList[28][g][n]):
            # #----↓參數名稱
            #     point_name[2] = "op" + str(op_number) +  "_u_cut_X_min_point"          #點
            #     point_name[4] = "op" + str(op_number) +  "_u_cut_X_max_point"
            #     point_name[6] = "op" + str(op_number) +  "_u_cut_Z_point"
            #     body_name[1] = "cut_cavity_insert_shear"           #body
            #     sketch_name[1] = "sketch_Coordinate"               #草圖
            #     type_name = "A_up"                                 #型式
            #     insert_line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) +  "_cut_punch_u_cutting_" + str( cut_punch_u_number)
            #     insert_line_name[2] = "plate_line_" + str(g) + "_op" + str(op_number) +  "_cut_punch_u_insert_surface_" + str( cut_punch_u_number)
            #     #------------------------------------------------------------↑
            #     Call punch_cutting_change           #換草圖
            #     Call punch_cutting                  #挖槽
            #     sketch_name[1] = "up_cut_insert_Sketch"
            #     Call punch_cutting                  #挖槽
            #     type_name = "C_left"
            #     Call punch_cutting_change           #換草圖
            #     sketch_name[1] = "left_cut_insert_Sketch"
            #     Call punch_cutting                  #挖槽
            pass  # 未使用
            # --------------------------------------------------------------------------------------------整形模穴
        if gvar.StripDataList[74][g][n] > 0:
            # for now_data_number in range(1, 1 + gvar.StripDataList[74][g][n]):
            #     F_bending_cut_cavity
            pass  # 未使用
        part1.Update()
        if gvar.StripDataList[37][g][n] > 0:
            A_punch(now_plate_line_number, op_number)
    selection1 = partDocument1.Selection
    selection1.Clear()
    selection1.Search("Name=cut_line_assume_*,all")
    if selection1.Count > 0:
        selection1.Delete()
    selection1.Clear()
    selection1.Search("Name=cut_cavity_insert,all")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    selection1.Search("Name=bolt_point_*,all")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    selection1.Search("Name=Sketch*,all")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    part1 = partDocument1.Part
    hybridBodies1 = part1.HybridBodies
    hybridBody1 = hybridBodies1.Item("cut_line_assume1")
    selection1.Add(hybridBody1)
    if selection1.Count > 0:
        selection1.Delete()
    cut_plat_machining_explanation_shape = 0  # -------------------------------------------------加工說明
    # ================================================↓挖下料孔↓================================================
    ss_count = 0
    for n in range(1, 1 + total_op_number):
        op_number = 10 * n
        cut_plat_machining_explanation_shape = cut_plat_machining_explanation_shape + 1  # -----------加工說明
        # ===============================================================補強入子
        if gvar.StripDataList[4][g][n] > 0:
            # for i in range (1 , 1+ gvar.StripDataList[4][g][n]):
            #     cut_line_st = "plate_line_" + str(g) + "_op" + str(n * 10) + "_Reinforcement_cut_line_" + str(i)
            #     boolean_hole_body
            pass  # 未使用
        # ===============================================================補強入子
        # ===============================================================A沖 plate_line_A_punch
        if gvar.StripDataList[37][g][n] > 0:
            for i in range(1, 1 + gvar.StripDataList[37][g][n]):
                cut_line_st = "plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i)
                (lower_pad_insert_hole_number, ss_count) = boolean_hole_body(ss_count, op_number, cut_line_st,
                                                                             lower_pad_insert_hole_number)
        # ===============================================================
        # ===============================================================型沖 plate_line_allotype_cut_line
        if gvar.StripDataList[39][g][n] > 0:
            for i in range(1, 1 + gvar.StripDataList[39][g][n]):
                cut_line_st = "plate_line_" + str(g) + "_op" + str(op_number) + "_allotype_cut_line_" + str(i)
                (lower_pad_insert_hole_number, ss_count) = boolean_hole_body(ss_count, op_number, cut_line_st,
                                                                             lower_pad_insert_hole_number)
        # ===============================================================
        # ===============================================================沖方形孔 plate_line_cut_line_number
        if gvar.StripDataList[38][g][n] > 0:
            for i in range(1, 1 + gvar.StripDataList[38][g][n]):
                cut_line_st = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i)
                (lower_pad_insert_hole_number, ss_count) = boolean_hole_body(ss_count, op_number, cut_line_st,
                                                                             lower_pad_insert_hole_number)
        # ===============================================================
        # ===============================================================沖T型異形孔 plate_line_unnomal_cut_line_T
        if gvar.StripDataList[53][g][n] > 0:
            # for i in range(1, 1 + gvar.StripDataList[53][g][n]):
            #     cut_line_st = symmetric_insert_name_T[n][i]
            #     (lower_pad_insert_hole_number) = boolean_hole_body(ss_count, op_number, cut_line_st, lower_pad_insert_hole_number)
            pass  # 未使用
        # ===============================================================沖I型異形孔 plate_line_unnomal_cut_line_I
        if gvar.StripDataList[54][g][n] > 0:
            # for i in range(1, 1 + gvar.StripDataList[54][g][n]):
            #     cut_line_st = symmetric_insert_name_I[n][i]
            #     (lower_pad_insert_hole_number) = boolean_hole_body(ss_count, op_number, cut_line_st, lower_pad_insert_hole_number)
            pass  # 未使用
        # ===============================================================
        # ===============================================================沖M型異形孔 plate_line_unnomal_cut_line_M
        if gvar.StripDataList[55][g][n] > 0:
            # for i in range(1, 1 + gvar.StripDataList[55][g][n]):
            #     cut_line_st = symmetric_insert_name_M[n][i]
            #     (lower_pad_insert_hole_number) = boolean_hole_body(ss_count, op_number, cut_line_st, lower_pad_insert_hole_number)
            pass  # 未使用
        # ===============================================================
    for n in range(1, 1 + total_op_number):
        op_number = 10 * n
        # ------------------------------------------------------------------------------↓快拆沖頭
        if gvar.StripDataList[29][g][n] > 0:  # 沖切沖頭_右
            # data_type = "line"
            # data_number = gvar.StripDataList[29][g][n]
            # part_name = "op" + str(op_number) + "_right_quickly_remove_cut_punch_insert_"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[30][g][n] > 0:  # 沖切沖頭_左
            # data_type = "line"
            # data_number = gvar.StripDataList[30][g][n]
            # part_name = "op" + str(op_number) + "_left_quickly_remove_cut_punch_insert_"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[31][g][n] > 0:  # 沖切沖頭_上
            # data_type = "line"
            # data_number = gvar.StripDataList[31][g][n]
            # part_name = "op" + str(op_number) + "_up_quickly_remove_cut_punch_insert_"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[32][g][n] > 0:  # 沖切沖頭_下
            # data_type = "line"
            # data_number = gvar.StripDataList[32][g][n]
            # part_name = "op" + str(op_number) + "_down_quickly_remove_cut_punch_insert_"
            # quickly_remove_punch
            pass  # 未使用
        # ------------------------------------------------------------------------------↑快拆沖頭
    product1.PartNumber = "lower_pad_" + str(g)  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("lower_pad_" + str(g))
    parameters1 = product1.UserRefProperties
    strParam1 = parameters1.CreateString("NO.", "")
    strParam1.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam2 = parameters5.CreateString("Part Name", "")
    strParam2.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters6 = product1.UserRefProperties
    strParam3 = parameters6.CreateString("Size", "")
    strParam3.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters7 = product1.UserRefProperties
    strParam4 = parameters7.CreateString("Material_Data", "")
    strParam4.ValuateFromString(gvar.strip_parameter_list[35])
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam5 = parameters5.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(gvar.strip_parameter_list[36])
    product1 = product1.ReferenceProduct
    parameters9 = product1.UserRefProperties
    strParam6 = parameters9.CreateString("Quantity", "")
    strParam6.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters10 = product1.UserRefProperties
    strParam7 = parameters10.CreateString("Page", "")
    strParam7.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters11 = product1.UserRefProperties
    strParam8 = parameters11.CreateString("L1", "")  # 形狀孔
    strParam8.ValuateFromString("L1: " + str(cut_plat_machining_explanation_shape) + "-(..,銑)")
    product1 = product1.ReferenceProduct
    parameters12 = product1.UserRefProperties
    strParam9 = parameters12.CreateString("A", "")  # 螺栓孔
    strParam9.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters13 = product1.UserRefProperties
    strParam10 = parameters13.CreateString("HP", "")  # 合銷孔
    strParam10.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters14 = product1.UserRefProperties
    strParam11 = parameters13.CreateString("B", "")  # B型引導沖孔
    strParam11.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters15 = product1.UserRefProperties
    strParam12 = parameters13.CreateString("BP", "")  # B沖沖孔
    strParam12.ValuateFromString("BP :" + str(cut_plat_machining_explanation_pilot_punch) + "- %%C" + str(
        Working_parameter) + "(B沖沖孔)")  # 介面參數設定OP10定位銷直徑
    product1 = product1.ReferenceProduct
    parameters16 = product1.UserRefProperties
    strParam13 = parameters16.CreateString("TS", "")  # 浮升引導
    strParam13.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters17 = product1.UserRefProperties
    strParam14 = parameters17.CreateString("IG", "")  # 內導柱
    strParam14.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters18 = product1.UserRefProperties
    strParam15 = parameters18.CreateString("F", "")  # 外導柱
    strParam15.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters19 = product1.UserRefProperties
    strParam16 = parameters19.CreateString("CS", "")  # 等高套筒
    strParam16.ValuateFromString("")
    parameters20 = product1.UserRefProperties
    strParam17 = parameters20.CreateString("AP", "")  # A沖沖孔
    strParam17.ValuateFromString("")
    # ====↑設定性質↑=====================================
    part1.Update()
    # --------------↓刪除不需要的Data↓--------------
    selection1 = partDocument1.Selection
    if now_plate_line_number == 1:
        selection1.Clear()
        selection1.Search("Name=plate_line_2*,all")
        if selection1.Count > 0:
            selection1.Delete()
        selection1.Clear
    if now_plate_line_number == 2:
        selection1.Clear()
        selection1 = partDocument1.Selection
        selection1.Search("Name=plate_line_1*,all")
        selection1.Clear()
    # --------------↑刪除不需要的Data↑--------------
    part1.Update()
    time.sleep(2)
    partDocument1.SaveAs(gvar.save_path + "lower_pad_" + str(g) + ".CATPart")  # 存檔的檔案名稱
    time.sleep(1)
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    partDocument1.Close()
    # ===================================================================關閉挖孔件
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()
    partDocument2 = catapp.ActiveDocument
    partDocument2.Close()
    # ===================================================================關閉挖孔件


def insert_Bolt_Hole(a, b, c, insert_line_name, sketch_name1, body_name1, type_name):
    # ------------------------------------------------------------↓   宣告變數
    length = [None] * 99
    pad_Bolt_Hole = [0] * 99
    # ------------------------------------------------------------↑
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    if gvar.StripDataList[27][a][b] > 0 or gvar.StripDataList[27][a][b] > 0:
        parameters1.Item("cut_line_formula_2").OptionalRelation.Modify("die\\" + insert_line_name[1])  # 草圖置換
        parameters1.Item("shear_Surface").OptionalRelation.Modify("die\\" + insert_line_name[2])  # 草圖置換
    else:
        parameters1.Item("cut_line_formula_1").OptionalRelation.Modify("die\\" + insert_line_name[1])  # 草圖置換
    # part1.Update()
    length[1] = parameters1.Item("long")
    length[2] = parameters1.Item("width")
    hybridShapeFactory1 = part1.HybridShapeFactory
    # ------------------------------------------↓搜尋現有pad_Bolt_Hole點的螺栓數量
    selection1.Clear()
    selection1.Search("Name=pad_" + str(a) + "_Bolt_point_*,all")
    pad_Bolt_Hole[a] = selection1.Count
    selection1.Clear()
    # ------------------------------------------↑搜尋現有pad_Bolt_Hole點的螺栓數量
    # ------------------------------------------↓判斷一個螺栓孔還是四個
    if body_name1 == "cut_cavity_insert_shear":
        if type_name == "A_up" or type_name == "B_down":
            ff = 1
            fff = 1
        elif type_name == "C_left":
            ff = 2
            fff = 2
    else:
        if length[1].Value > 50 and length[2].Value > 50:
            ff = 2
            fff = 5
        elif length[1].Value > 50 and length[2].Value < 50:
            ff = 3
            fff = 4
        elif length[1].Value < 50 and length[2].Value > 50:
            ff = 3
            fff = 4
        else:
            ff = 1
            fff = 1
    # ------------------------------------------↑
    for CC in range(ff, 1 + fff):
        shapeFactory1 = part1.ShapeFactory
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        body2 = bodies1.Item(body_name1)
        sketches1 = body2.Sketches
        sketch1 = sketches1.Item(sketch_name1)
        part1.UpdateObject(sketch1)
        if gvar.StripDataList[27][a][b] > 0 or gvar.StripDataList[28][a][b] > 0:
            length[34] = parameters1.Item("x_to_x_shear")
            length[35] = parameters1.Item("y_to_y_shear")
            length[36] = parameters1.Item("int_x_shear")
            length[37] = parameters1.Item("int_y_shear")
            strParam1 = parameters1.Item("Type_shear")
            strParam1.Value = type_name
        else:
            length[34] = parameters1.Item("x_to_x")
            length[35] = parameters1.Item("y_to_y")
            length[36] = parameters1.Item("int_x")
            length[37] = parameters1.Item("int_y")
            strParam1 = parameters1.Item("Type")
        # ------------------------------------------------------------↓   在已有平面建立條件             (方法二)
        hybridShapes1 = body1.HybridShapes
        hybridShapePlaneOffset1 = hybridShapes1.Item("up_plane")
        reference1 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   零建檔建立點之語法"HybridShapes"包涵在body裡
        hybridShapes2 = body2.HybridShapes
        hybridShapePointOnPlane1 = hybridShapes2.Item("Bolt_point_" + str(CC))
        reference2 = part1.CreateReferenceFromObject(hybridShapePointOnPlane1)
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   建立點
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, -11)  # (x,y,z)
        hybridShapePointCoord1.PtRef = reference2  # 基礎點 = reference2
        body1.InsertHybridShape(hybridShapePointCoord1)  # 建立點
        # part1.Update()
        # ------------------------------------------------------------↑
        reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
        pad_Bolt_Hole[a] = pad_Bolt_Hole[a] + 1
        # ------------------------------------------------------------↓   建立點(打斷關連的)
        hybridShapePointExplicit1 = hybridShapeFactory1.AddNewPointDatum(reference3)  # 將點置於元素3
        body1.InsertHybridShape(hybridShapePointExplicit1)
        part1.InWorkObject = hybridShapePointExplicit1
        hybridShapePointExplicit1.Name = "pad_" + str(a) + "_Bolt_point_" + str(pad_Bolt_Hole[a])  # 點改名字
        # part1.Update()
        hybridShapeFactory1.DeleteObjectForDatum(reference3)  # 刪除元素3
        # ------------------------------------------------------------↑
        reference4 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
        hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference4, reference1, 15)
        hole1.Name = "Hole_OP" + str(b * 10) + "_pick_" + str(c) + "_" + str(CC)  # 孔改名字
        hole1.Type = 0
        hole1.AnchorMode = 0
        hole1.BottomType = 0
        # ------------------------------------------------------------↓     孔的型態設定
        hole1.ThreadingMode = 0  # (螺紋孔catThreadedHoleThreading  OR 無螺紋孔catSmoothHoleThreading)
        hole1.ThreadSide = 0  # 左OR右螺紋
        hole1.Type = 0  # 直孔  (沉頭孔令外再用草圖挖)
        # ------------------------------------------------------------↑
        limit1 = hole1.BottomLimit
        limit1.LimitMode = 0
        length1 = hole1.Diameter
        hole1.CreateStandardThreadDesignTable(1)
        strParam1 = hole1.HoleThreadDescription
        strParam1.Value = "M8"
        # =============================================↓   極限設定
        limit2 = hole1.BottomLimit
        limit2.LimitMode = 0  # 未貫穿(盲孔)
        # =============================================↑
        # =============================================↓   牙深設定 孔深>牙深(否則錯誤)
        length2 = limit2.dimension  # 孔深
        length2.Value = int(gvar.strip_parameter_list[29])  # 16
        # =============================================↑
        # =============================================↓   牙深設定 孔深>牙深(否則錯誤)
        length3 = hole1.ThreadDepth  # 牙深
        length3.Value = length2.Value - 2
        # =============================================↑
        sketch1 = hole1.sketch
        selection1.Add(sketch1)
        selection1.Add(hybridShapePointExplicit1)
        visPropertySet1 = visPropertySet1.Parent
        visPropertySet1.SetShow(1)
        selection1.Clear()
        part1.Update()


def auto_insert_Bolt_Hole(a, b, c, insert_line_name, sketch_name1, body_name1, type_name):
    # ------------------------------------------------------------↓   宣告變數
    length = [None] * 99
    pad_Bolt_Hole = [0] * 99
    # ------------------------------------------------------------↑
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    parameters1 = part1.Parameters
    if gvar.StripDataList[27][a][b] > 0 or gvar.StripDataList[28][a][b] > 0:
        parameters1.Item("cut_line_formula_2").OptionalRelation.Modify("die\\" + insert_line_name[1])  # 草圖置換
        parameters1.Item("shear_Surface").OptionalRelation.Modify("die\\" + insert_line_name[2])  # 草圖置換
    else:
        parameters1.Item("cut_line_formula_1").OptionalRelation.Modify("die\\" + insert_line_name[1])  # 草圖置換
    part1.Update()
    length[1] = parameters1.Item("long")
    length[2] = parameters1.Item("width")
    hybridShapeFactory1 = part1.HybridShapeFactory
    # ------------------------------------------↓判斷一個螺栓孔還是四個
    if body_name1 == "cut_cavity_insert_shear":
        if type_name == "A_up" or type_name == "B_down":
            ff = 1
            fff = 1
        elif type_name == "C_left":
            ff = 2
            fff = 2
    else:
        if length[1].Value > 50 and length[2].Value > 50:
            ff = 2
            fff = 5
        elif length[1].Value > 50 and length[2].Value < 50:
            ff = 3
            fff = 4
        elif length[1].Value < 50 and length[2].Value > 50:
            ff = 3
            fff = 4
        else:
            ff = 1
            fff = 1
    # ------------------------------------------↑
    for CC in range(ff, 1 + fff):
        shapeFactory1 = part1.ShapeFactory
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        body2 = bodies1.Item(body_name1)
        sketches1 = body2.Sketches
        sketch1 = sketches1.Item(sketch_name1)
        part1.UpdateObject(sketch1)
        if gvar.StripDataList[27][a][b] > 0 or gvar.StripDataList[28][a][b] > 0:
            length[34] = parameters1.Item("x_to_x_shear")
            length[35] = parameters1.Item("y_to_y_shear")
            length[36] = parameters1.Item("int_x_shear")
            length[37] = parameters1.Item("int_y_shear")
            strParam1 = parameters1.Item("Type_shear")
            strParam1.Value = type_name
        else:
            length[34] = parameters1.Item("x_to_x")
            length[35] = parameters1.Item("y_to_y")
            length[36] = parameters1.Item("int_x")
            length[37] = parameters1.Item("int_y")
            strParam1 = parameters1.Item("Type")
        # ------------------------------------------------------------↓   在已有平面建立條件             (方法二)
        hybridShapes1 = body1.HybridShapes
        hybridShapePlaneOffset1 = hybridShapes1.Item("up_plane")
        reference1 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   零建檔建立點之語法"HybridShapes"包涵在body裡
        hybridShapes2 = body2.HybridShapes
        hybridShapePointOnPlane1 = hybridShapes2.Item("Bolt_point_" + str(CC))
        reference2 = part1.CreateReferenceFromObject(hybridShapePointOnPlane1)
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   建立點
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, -11)  # (x,y,z)
        hybridShapePointCoord1.PtRef = reference2  # 基礎點 = reference2
        body1.InsertHybridShape(hybridShapePointCoord1)  # 建立點
        part1.Update()
        # ------------------------------------------------------------↑
        reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
        pad_Bolt_Hole[a] = pad_Bolt_Hole[a] + 1
        # ------------------------------------------------------------↓   建立點(打斷關連的)
        hybridShapePointExplicit1 = hybridShapeFactory1.AddNewPointDatum(reference3)  # 將點置於元素3
        body1.InsertHybridShape(hybridShapePointExplicit1)
        part1.InWorkObject = hybridShapePointExplicit1
        hybridShapePointExplicit1.Name = "pad_" + str(a) + "_Bolt_point_" + str(pad_Bolt_Hole[a])  # 點改名字
        part1.Update()
        hybridShapeFactory1.DeleteObjectForDatum(reference3)  # 刪除元素3
        # ------------------------------------------------------------↑
        reference4 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
        hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference4, reference1, 15)
        hole1.Name = "Hole_OP" + str(b) * 10 + "_pick_" + str(c) + "_" + str(CC)  # 孔改名字
        hole1.Type = 0
        hole1.AnchorMode = 0
        hole1.BottomType = 0
        # ------------------------------------------------------------↓     孔的型態設定
        hole1.ThreadingMode = 0  # (螺紋孔catThreadedHoleThreading  OR 無螺紋孔catSmoothHoleThreading)
        hole1.ThreadSide = 0  # 左OR右螺紋
        hole1.Type = 0  # 直孔  (沉頭孔令外再用草圖挖)
        # ------------------------------------------------------------↑
        limit1 = hole1.BottomLimit
        limit1.LimitMode = 0
        length1 = hole1.Diameter
        hole1.CreateStandardThreadDesignTable(1)
        strParam1 = hole1.HoleThreadDescription
        strParam1.Value = "M8"
        # =============================================↓   極限設定
        limit2 = hole1.BottomLimit
        limit2.LimitMode = 0  # 未貫穿(盲孔)
        # =============================================↑
        # =============================================↓   牙深設定 孔深>牙深(否則錯誤)
        # length2 As length
        length2 = limit2.dimension  # 孔深

        length2.Value = int(gvar.strip_parameter_list[29])  # 16
        # =============================================↑
        # =============================================↓   牙深設定 孔深>牙深(否則錯誤)
        length3 = hole1.ThreadDepth  # 牙深
        length3.Value = length2.Value - 2
        # =============================================↑
        selection1.Delete()
        sketch1 = hole1.sketch
        selection1.Add(sketch1)
        selection1.Add(hybridShapePointExplicit1)
        visPropertySet1 = visPropertySet1.Parent
        visPropertySet1.SetShow(1)
        selection1.Clear()
        part1.UpdateObject(hole1)


def boolean_hole_body(ss_count, op_number, cut_line_st, lower_pad_insert_hole_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument2 = catapp.ActiveDocument
    length = [None] * 99
    formula = [None] * 20
    parameter = [None] * 20
    if ss_count < 1:
        partDocument1 = documents1.Open(gvar.open_path + "cut_line.CATPart")
        product1 = partDocument1.getItem("Part1")
        part1 = partDocument1.Part
        # =======================================================================偏移距離
        length[1] = part1.Parameters.Item("operation_length")
        length[1].Value = 1
        # =======================================================================偏移距離
        part1.Update()
        ss_count = ss_count + 1
        # =======================================================================視窗切換
    elif ss_count > 0:
        windows1 = catapp.Windows
        specsAndGeomWindow1 = windows1.Item("cut_line.CATPart")
        specsAndGeomWindow1.Activate()
        ss_count = ss_count + 1
    # =======================================================================視窗切換
    # ======================
    time.sleep(1)
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    selection1.Clear()
    selection2 = partDocument1.Selection
    selection2.Clear()
    part1 = partDocument1.Part
    # 搜尋所有自設規則
    # ===================================================================
    relations1 = part1.Relations
    formulal_Count = part1.Relations.Count
    for form_number in range(1, 1 + formulal_Count):
        formula1 = relations1.Item(form_number)
        selection1.Add(formula1)
        time.sleep(0.1)
    # ===================================================================
    # 搜尋所有自設變數
    # ===================================================================
    parameters1 = part1.Parameters
    parameter_Count = part1.Parameters.RootParameterSet.DirectParameters.Count
    for parame_number in range(1, 1 + parameter_Count):
        # paramet1 As Object #formul與Rule同時出現的話必須使用Object
        paramet1 = parameters1.RootParameterSet.DirectParameters.Item(parame_number)
        selection1.Add(paramet1)
        time.sleep(0.1)
    # ===================================================================
    # 搜尋所有自設變數支點
    # ===================================================================
    parameters2 = part1.Parameters
    parameter_Count = parameters2.RootParameterSet.DirectParameters.Count
    parameter_Count = parameters2.RootParameterSet.ParameterSets.Count
    for parame_number in range(1, 1 + parameter_Count):
        paramet1 = parameters1.RootParameterSet.ParameterSets.Item(parame_number)
        selection1.Add(paramet1)
        time.sleep(0.1)
    # ===================================================================
    # 搜尋所有自設規則(群組)
    # ===================================================================
    anyObject1 = part1.Relations.getItem("Relations")
    selection1.Add(anyObject1)
    # ===================================================================
    constraints1 = part1.Constraints
    Constraints_Count = part1.Constraints.Count
    for Constraints_number in range(1, 1 + Constraints_Count):
        constraint1 = constraints1.Item(Constraints_number)
        selection1.Add(constraint1)
        time.sleep(0.1)
    # ===================================================================
    bodies1 = part1.Bodies
    bodies_Count = part1.Bodies.Count
    for bodies_number in range(1, 1 + bodies_Count):
        bodie1 = bodies1.Item(bodies_number)
        selection1.Add(bodie1)
        time.sleep(0.1)
    # 搜尋座標設置
    # AxisSystems1 As AxisSystems
    AxisSystems1 = part1.AxisSystems
    AxisSystems_Count = part1.AxisSystems.Count
    for Axis_number in range(1, 1 + AxisSystems_Count):
        Axis1 = AxisSystems1.Item(Axis_number)
        selection1.Add(Axis1)
        time.sleep(0.1)
    # 搜尋所有自設幾何(群組)
    # ===================================================================
    hybridBody_Count = part1.HybridBodies.Count
    for hybridBody_number in range(1, 1 + hybridBody_Count):
        hybridBody1 = part1.HybridBodies.Item(hybridBody_number)
        selection1.Add(hybridBody1)
        time.sleep(0.1)
    time.sleep(1)
    selection1.Copy()
    time.sleep(1)
    partDocument2.Activate()
    # ===================================================================
    partDocument1 = catapp.ActiveDocument
    selection3 = partDocument1.Selection
    time.sleep(1)
    selection3.Clear()
    part1 = partDocument1.Part
    selection3.Add(part1)
    time.sleep(1)
    selection3.Paste()
    bodies1 = part1.Bodies
    for eq_sum in range(50, 1 + 1, -1):
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=Body." + str(eq_sum) + ",all")
        if selection1.Count > 0:
            break
    body1 = bodies1.Item("Body." + str(eq_sum))
    part1.InWorkObject = body1
    part1.InWorkObject.Name = "boolean_hole_body_op" + str(op_number) + "_" + str(ss_count)
    # ======================================================
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("boolean_hole_body_op" + str(op_number) + "_" + str(ss_count))
    part1.InWorkObject = body1
    # ======================================================
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    hybridShapeCurveExplicit1 = parameters1.Item("punch_cut_line_assume_1")
    part1.InWorkObject = hybridShapeCurveExplicit1
    part1.InWorkObject.Name = "punch_cut_line_assume_" + str(ss_count + 1)
    # ======================================================
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("boolean_hole_body_op" + str(op_number) + "_" + str(ss_count))
    shapes1 = body1.Shapes
    pad1 = shapes1.Item("finally_pad")
    part1.InWorkObject = pad1
    part1.Update()
    # ======================================================
    part1.Parameters.Item("punch_cut_line_assume_" + str(ss_count + 1)).OptionalRelation.Modify(
        "die\\" + cut_line_st)  # 草圖置換
    parameters2 = part1.Parameters
    part1.Update()
    # ======================================================
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    part1.InWorkObject = body1
    body2 = bodies1.Item("Body.2")
    part1.InWorkObject = body2
    shapeFactory1 = part1.ShapeFactory
    body3 = bodies1.Item("boolean_hole_body_op" + str(op_number) + "_" + str(ss_count))
    remove1 = shapeFactory1.AddNewRemove(body3)
    lower_pad_insert_hole_number = lower_pad_insert_hole_number + 1
    remove1.Name = "Lower-pad-Insert-hole-" + str(lower_pad_insert_hole_number)
    part1.Update()
    # ======================================================
    return lower_pad_insert_hole_number, ss_count


def A_punch(now_plate_line_number, op_number):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    hybridShapeFactory1 = part1.HybridShapeFactory
    parameters1 = part1.Parameters
    length = [None] * 50
    insert_line_name = [""] * 4
    g = now_plate_line_number
    n = int(op_number / 10)
    cut_plat_machining_explanation_A_punch = 0  # -----------------------------------------------------加工說明
    for i in range(1, 1 + gvar.StripDataList[37][g][n]):
        # ----------------------------------------------------------------------------↓對稱入子
        # ------------------------------------------------------------↓   搜尋是否為對稱名稱
        selection3 = partDocument1.Selection
        selection3.Clear()
        selection3.Search("Name=plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_symmetric_" + str(i))
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   判斷是否對稱名稱
        if selection3.Count > 0:
            part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_symmetric_" + str(i))  # 草圖置換
        else:
            part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i))  # 草圖置換
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   參數宣告
        length[13] = part1.Parameters.Item("insert_line")  # 多出來的
        length[13].Value = 5  # 多出來的
        length[33] = part1.Parameters.Item("gap")
        length[33].Value = float(gvar.strip_parameter_list[1]) * 0.02  # (lower_die_space / 100 )                #間隙
        length[30] = part1.Parameters.Item("cut_cavity_insert_height")
        length[30].Value = float(gvar.strip_parameter_list[26])
        length[31] = part1.Parameters.Item("die_open_height")
        length[31].Value = 28  # (die_open_height)
        length[34] = part1.Parameters.Item("x_to_x")
        length[35] = part1.Parameters.Item("y_to_y")
        length[36] = part1.Parameters.Item("int_x")
        length[37] = part1.Parameters.Item("int_y")
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   先放大入子以免更換時  孔>入子塊
        length[36].Value = 500
        length[37].Value = 500
        # ------------------------------------------------------------↑
        part1.Update()
        # ------------------------------------------------------------↓  修正為對稱孔位之參數 True=對稱  False=不對稱
        boolParam1 = part1.Parameters.Item("symmetry_switch")
        if selection3.Count > 0:
            boolParam1.Value = True
        else:
            boolParam1.Value = False
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   整數化
        length[36].Value = math.ceil(length[34].Value)
        length[37].Value = math.ceil(length[35].Value)
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓
        body_name1 = "cut_cavity_insert"
        sketch_name1 = "cut_insert_line_Sketch"
        if selection3.Count > 0:
            insert_line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_symmetric_" + str(i)
        else:
            insert_line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i)
        # ------------------------------------------------------------↑
        part1.Update()
        insert_Bolt_Hole(g, n, i, insert_line_name, sketch_name1, body_name1, type_name=None)  # 挖孔指令
        cut_plat_machining_explanation_A_punch = cut_plat_machining_explanation_A_punch + 1  # --------加工說明
        part1.Update()

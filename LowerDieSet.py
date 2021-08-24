import win32com.client as win32
import global_var as gvar
import defs
import time

lower_die_set_machining_explanation_shape = 0  # -----------------------------------加工說明


def LowerDieSet():
    total_op_number = int(gvar.strip_parameter_list[2])
    catapp = win32.Dispatch('CATIA.Application')
    out_Guide_Material = "MYJP"
    out_Guide_Diameter = 32
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    partDocument2 = documents1.Open(gvar.open_path + "Data1.CATPart")
    # ======================================
    defs.window_change(partDocument1, partDocument2)
    # ======================================
    # ======================================(lower_die_set_change)
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    selection1 = partDocument1.Selection
    length = [None] * 20
    formula = [None] * 20
    parameter = [None] * 20
    lower_die_set_A_punch_number = int()
    lower_die_set_insert_hole_number = int()
    # ======================================================================================================
    length[1] = part1.Parameters.Item("plate_height")
    length[1].Value = -float(gvar.strip_parameter_list[32])
    length[4] = part1.Parameters.Item("plate_up_plane")
    length[4].Value = -0 - float(gvar.strip_parameter_list[26]) - float(
        gvar.strip_parameter_list[29])  # die_open_height
    length[5] = part1.Parameters.Item("center_point_distance_X")
    length[5].Value = 0
    length[6] = part1.Parameters.Item("Pillar_seat_depth")
    length[6].Value = 30
    part1.Update()
    # ======================================================================================================
    parameters1 = part1.Parameters
    strParam20 = parameters1.Item("outer_guiding_post_bolt")  # 外導柱固定螺栓直徑
    length[7] = part1.Parameters.Item("outer_guiding_post_pin")  # 外導柱固定螺栓合銷
    if out_Guide_Material == "MYJP":
        if out_Guide_Diameter == 20:
            strParam20.Value = "M6"
            length[7].Value = 6
        elif out_Guide_Diameter == 25:
            strParam20.Value = "M8"
            length[7].Value = 8
        elif out_Guide_Diameter == 32:
            strParam20.Value = "M10"
            length[7].Value = 8
        elif out_Guide_Diameter == 38:
            strParam20.Value = "M10"
            length[7].Value = 10
        elif out_Guide_Diameter == 50:
            strParam20.Value = "M12"
            length[7].Value = 10
    # ----------------------------------------------------------
    elif out_Guide_Material == "MYKP":
        if out_Guide_Diameter == 20:
            strParam20.Value = "M8"
            length[7].Value = 8
        elif out_Guide_Diameter == 25:
            strParam20.Value = "M8"
            length[7].Value = 8
        elif out_Guide_Diameter == 32:
            strParam20.Value = "M10"
            length[7].Value = 8
        elif out_Guide_Diameter == 38:
            strParam20.Value = "M10"
            length[7].Value = 10
        elif out_Guide_Diameter == 50:
            strParam20.Value = "M12"
            length[7].Value = 10
    # ----------------------------------------------------------
    elif out_Guide_Material == "DANLY":
        if out_Guide_Diameter == 25:
            strParam20.Value = "M6"
        elif out_Guide_Diameter == 32:
            strParam20.Value = "M8"
        elif out_Guide_Diameter == 40:
            strParam20.Value = "M8"
        elif out_Guide_Diameter == 50:
            strParam20.Value = "M8"
        elif out_Guide_Diameter == 63:
            strParam20.Value = "M8"
        elif out_Guide_Diameter == 80:
            strParam20.Value = "M8"
        if out_Guide_Diameter > 25:
            # lower_die_set_DANLY_MODEL_1  # 螺栓孔形式變更
            pass  # 未使用
        elif out_Guide_Diameter <= 25:
            # lower_die_set_DANLY_MODEL_2  # 螺栓孔形式變更
            pass  # 未使用
    # ======================================================================================================
    part1.Update()
    part1 = partDocument1.Part
    part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\lower_die_seat_line")  # 草圖置換
    part1.Update()
    specsAndGeomWindow1 = catapp.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    part1.Update()
    parameters1 = part1.Parameters
    relations1 = part1.Relations
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    sketches1 = body1.Sketches
    sketch1 = sketches1.Item("Sketch.135")
    part1.UpdateObject(sketch1)
    # ======================================================================================================
    length1 = parameters1.Item("plate_length")
    lower_die_set_length = length1.Value
    parameters2 = part1.Parameters
    length2 = parameters2.Item("plate_width")
    lower_die_set_width = length2.Value
    # ======================================================================================================
    length[8] = part1.Parameters.Item("guide plate_bolt_distance")
    length[8].Value = 0  # (part_plate)
    length[9] = part1.Parameters.Item("guide plate_bolt_pitch")
    length[9].Value = 0  # (bolt_Length)
    # ======================================================================================================
    # ------------------------------整塊式及分塊式孔位選擇
    part_plate = 0
    if part_plate < 300:
        parameters18 = part1.Parameters
        strParam1 = parameters18.Item("String.1")
        strParam1.Value = "B"
    # ------------------------------整塊式及分塊式孔位選擇
    # ======================================================================================================
    part1.Update()
    hybridShapeFactory1 = part1.HybridShapeFactory
    parameters3 = part1.Parameters
    part1.Update()
    if gvar.die_type == "module":
        # M_lower_hole
        pass  # 未使用
    # =================↓挖下料孔↓================================================
    if gvar.die_type == "moudle":
        X_pin1 = 17.5
        Y_pin1 = 175
        X_pin2 = 117.5
        Y_pin2 = 175
        X_pin3 = 182.5
        Y_pin3 = 175
        X_pin4 = 252.5
        Y_pin4 = 175
        part_file_name = "lower_die_set"
        body_name_1 = "PartBody"
        hybridBody_name = "die"
        (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(part_file_name,
                                                                                                 body_name_1,
                                                                                                 hybridBody_name)  # 環境設定(檔案名,body名,依據群組名)(全域變數改)
        hybridShape1 = ElementHybridBody.HybridShapes.Item("number_1_plate_line")  # 宣告元素(當下模板的曲線)
        (element_Reference1) = defs.ExtremumPoint("X_min", "Y_min", 0, 2, hybridShape1, ElementDocument, ElementBody,
                                                  ElementHybridBody)  # 建立極值點   (極值方向1,極值方向2,極值方向3,需要幾個方向,在哪個元素上的極值)  element_Reference[1]為OUT
        element_Reference1.Name = "number_1_plate_min"  # 建立最小點
        element_Reference10 = element_Reference1  # 建點之依據
        element_point = [None] * 5
        # ----------------------------------------------------------------------------------------------------------------pin_1
        (element_point5) = defs.BuildPoint(element_Reference10, ElementDocument, ElementBody, ElementHybridBody,
                                           SketchPosition="HybridBody")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point[5] 為out
        element_point5.X.Value = X_pin1
        element_point5.Y.Value = Y_pin1
        element_point[1] = element_point5
        # ----------------------------------------------------------------------------------------------------------------pin_2
        (element_point5) = defs.BuildPoint(element_Reference10, ElementDocument, ElementBody, ElementHybridBody,
                                           SketchPosition="HybridBody")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point[5] 為out
        element_point5.X.Value = X_pin2
        element_point5.Y.Value = Y_pin2
        element_point[2] = element_point5
        # ----------------------------------------------------------------------------------------------------------------pin_2
        (element_point5) = defs.BuildPoint(element_Reference10, ElementDocument, ElementBody, ElementHybridBody,
                                           SketchPosition="HybridBody")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point[5] 為out
        element_point5.X.Value = X_pin3
        element_point5.Y.Value = Y_pin3
        element_point[3] = element_point5
        # ----------------------------------------------------------------------------------------------------------------pin_4
        (element_point5) = defs.BuildPoint(element_Reference10, ElementDocument, ElementBody, ElementHybridBody,
                                           SketchPosition="HybridBody")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point[5] 為out
        element_point5.X.Value = X_pin4
        element_point5.Y.Value = Y_pin4
        element_point[4] = element_point5
        # ----------------------------------------------------------------------------------------------------------------pin_4
        hybridShape2 = ElementHybridBody.HybridShapes.Item("up_plane")  # 宣告平面
        element_Reference12 = hybridShape2
        for location_pin_number in range(1, 5):
            element_Reference11 = element_point[location_pin_number]
            (hole) = defs.HoleSimpleD(8, float(gvar.strip_parameter_list[32]), 0, ElementDocument, ElementBody,
                                      element_Reference11,
                                      element_Reference12)  # 直孔  (M,深度,out孔陳述式,方向) element_Reference(11)=point #element_Reference(12)=plane direction=>0=下 1=上
    ss_count = 0
    for g in range(1, gvar.PlateLineNumber + 1):
        if gvar.die_type == "moudle":
            # location_block
            pass  # 未使用
        for n in range(1, 1 + total_op_number):
            op_number = 10 * n
            # ------------------------------------------------------------------------------------------------↓ 補強入子
            if gvar.StripDataList[4][g][n] > 0:  # 補強入子
                for for_counter in range(1, round(gvar.StripDataList[4][g][n]) + 1):
                    cut_line_st = "plate_line_" + str(g) + "_op" + str(n * 10) + "_Reinforcement_cut_line_" + str(
                        for_counter)
                    (lower_die_set_A_punch_number, lower_die_set_insert_hole_number) = boolean_hole_body("insert",
                                                                                                         lower_die_set_A_punch_number,
                                                                                                         lower_die_set_insert_hole_number,
                                                                                                         cut_line_st)
            # ------------------------------------------------------------------------------------------------↑ 補強入子
            # ===============================================================A沖 plate_line_A_punch
            if gvar.StripDataList[37][g][n] > 0:
                for i in range(1, 1 + gvar.StripDataList[37][g][n]):
                    cut_line_st = "plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i)
                    (lower_die_set_A_punch_number, lower_die_set_insert_hole_number) = boolean_hole_body("A_punch",
                                                                                                         lower_die_set_A_punch_number,
                                                                                                         lower_die_set_insert_hole_number,
                                                                                                         cut_line_st)
            # ===============================================================
            # ===============================================================型沖 plate_line_allotype_cut_line
            if gvar.StripDataList[39][g][n] > 0:
                for i in range(1, 1 + gvar.StripDataList[39][g][n]):
                    cut_line_st = "plate_line_" + str(g) + "_op" + str(op_number) + "_allotype_cut_line_" + str(i)
                    (lower_die_set_A_punch_number, lower_die_set_insert_hole_number) = boolean_hole_body("insert",
                                                                                                         lower_die_set_A_punch_number,
                                                                                                         lower_die_set_insert_hole_number,
                                                                                                         cut_line_st)
            # ===============================================================
            # ===============================================================沖方形孔 plate_line_cut_line_number
            if gvar.StripDataList[38][g][n] > 0:
                for i in range(1, 1 + gvar.StripDataList[38][g][n]):
                    cut_line_st = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i)
                    (lower_die_set_A_punch_number, lower_die_set_insert_hole_number) = boolean_hole_body("insert",
                                                                                                         lower_die_set_A_punch_number,
                                                                                                         lower_die_set_insert_hole_number,
                                                                                                         cut_line_st)
            # ===============================================================
            # ===============================================================沖T型異形孔 plate_line_unnomal_cut_line_T
            if gvar.StripDataList[53][g][n] > 0:
                for i in range(1, 1 + gvar.StripDataList[53][g][n]):
                    line_name = [""] * 4
                    line_name[2] = "plate_line_" + str(g) + "_op" + str(
                        op_number) + "_unnomal_cut_line_T_symmetric_" + str(i)
                    line_name[3] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_T_" + str(i)
                    # ------------------------------------------------------------↓   判斷是否對稱名稱
                    selection1 = partDocument1.Selection
                    selection1.Clear()
                    selection1.Search("Name=" + line_name[2])
                    # ------------------------------------------------------------↑
                    # ------------------------------------------------------------↓   判斷是否對稱名稱
                    if selection1.Count > 0:
                        line_name[1] = line_name[2]
                    else:
                        line_name[1] = line_name[3]
                    # ------------------------------------------------------------↑
                    cut_line_st = line_name[1]
                    (lower_die_set_A_punch_number, lower_die_set_insert_hole_number) = boolean_hole_body("insert",
                                                                                                         lower_die_set_A_punch_number,
                                                                                                         lower_die_set_insert_hole_number,
                                                                                                         cut_line_st)
            # ===============================================================
            # ===============================================================沖I型異形孔 plate_line_unnomal_cut_line_I
            if gvar.StripDataList[54][g][n] > 0:
                for i in range(1, 1 + gvar.StripDataList[54][g][n]):
                    line_name = [""] * 4
                    line_name[2] = "plate_line_" + str(g) + "_op" + str(
                        op_number) + "_unnomal_cut_line_I_symmetric_" + str(i)
                    line_name[3] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_I_" + str(i)
                    # ------------------------------------------------------------↓   判斷是否對稱名稱
                    selection1 = partDocument1.Selection
                    selection1.Clear()
                    selection1.Search("Name=" + line_name[2])
                    # ------------------------------------------------------------↑
                    # ------------------------------------------------------------↓   判斷是否對稱名稱
                    if selection1.Count > 0:
                        line_name[1] = line_name[2]
                    else:
                        line_name[1] = line_name[3]
                    # ------------------------------------------------------------↑
                    cut_line_st = line_name[1]
                    (lower_die_set_A_punch_number, lower_die_set_insert_hole_number) = boolean_hole_body("insert",
                                                                                                         lower_die_set_A_punch_number,
                                                                                                         lower_die_set_insert_hole_number,
                                                                                                         cut_line_st)
            # ===============================================================
            # ===============================================================沖M型異形孔 plate_line_unnomal_cut_line_M
            if gvar.StripDataList[55][g][n] > 0:
                for i in range(1, 1 + gvar.StripDataList[55][g][n]):
                    line_name = [""] * 4
                    line_name[2] = "plate_line_" + str(g) + "_op" + str(
                        op_number) + "_unnomal_cut_line_M_symmetric_" + str(i)
                    line_name[3] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_M_" + str(i)
                    # ------------------------------------------------------------↓   判斷是否對稱名稱
                    selection1 = partDocument1.Selection
                    selection1.Clear()
                    selection1.Search("Name=" + line_name[2])
                    # ------------------------------------------------------------↑
                    # ------------------------------------------------------------↓   判斷是否對稱名稱
                    if selection1.Count > 0:
                        line_name[1] = line_name[2]
                    else:
                        line_name[1] = line_name[3]
                    # ------------------------------------------------------------↑
                    cut_line_st = line_name[1]
                    (lower_die_set_A_punch_number, lower_die_set_insert_hole_number) = boolean_hole_body("insert",
                                                                                                         lower_die_set_A_punch_number,
                                                                                                         lower_die_set_insert_hole_number,
                                                                                                         cut_line_st)
            # ===============================================================
    # =================↑挖下料孔↑================================================
    for g in range(1, gvar.PlateLineNumber + 1):
        for n in range(1, 1 + total_op_number):
            op_number = 10 * n
            # ------------------------------------------------------------------------------↓快拆沖頭
            if gvar.StripDataList[29][g][n] > 0:  # '沖切沖頭_右
                # data_type = "line"
                # data_number = round(gvar.StripDataList[29][g][n])
                # part_name = "op" + str(op_number) + "_right_quickly_remove_cut_punch_"
                # PunchDef.quickly_remove_punch(data_type, data_number, part_name)
                pass  # 未使用
            if gvar.StripDataList[30][g][n] > 0:  # 沖切沖頭_左
                # data_type = "line"
                # data_number = round(gvar.StripDataList[30][g][n])
                # part_name = "op" + str(op_number) + "_left_quickly_remove_cut_punch_"
                # PunchDef.quickly_remove_punch(data_type, data_number, part_name)
                pass  # 未使用
            if gvar.StripDataList[31][g][n] > 0:  # 沖切沖頭_上
                # data_type = "line"
                # data_number = gvar.StripDataList[31][g][n]
                # part_name = "op" + str(op_number) + "_up_quickly_remove_cut_punch_"
                # PunchDef.quickly_remove_punch(data_type, data_number, part_name)
                pass  # 未使用
            if gvar.StripDataList[32][g][n] > 0:  # 沖切沖頭_下
                # data_type = "line"
                # data_number = gvar.StripDataList[32][g][n]
                # part_name = "op" + str(op_number) + "_down_quickly_remove_cut_punch_"
                # PunchDef.quickly_remove_punch(data_type, data_number, part_name)
                pass  # 未使用
            # ------------------------------------------------------------------------------↑快拆沖頭
            if gvar.StripDataList[27][g][n] > 0:
                # punch_d_cutting
                pass  # 未使用
            if gvar.StripDataList[28][g][n] > 0:
                # punch_u_cutting
                pass  # 未使用
    # ============================================================================================================================感測器-----前計畫
    product1.PartNumber = "lower_die_set"  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("lower_die_set")
    parameters1 = product1.UserRefProperties
    strParam1 = parameters1.CreateString("NO.", "")
    strParam1.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters4 = product1.UserRefProperties
    strParam2 = parameters4.CreateString("Part Name", "")
    strParam2.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam3 = parameters5.CreateString("Size", "")
    strParam3.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters6 = product1.UserRefProperties
    strParam4 = parameters6.CreateString("Material_Data", "")
    strParam4.ValuateFromString(gvar.strip_parameter_list[33])
    product1 = product1.ReferenceProduct
    parameters7 = product1.UserRefProperties
    strParam5 = parameters7.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(gvar.strip_parameter_list[34])
    product1 = product1.ReferenceProduct
    parameters8 = product1.UserRefProperties
    strParam6 = parameters8.CreateString("Quantity", "")
    strParam6.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters9 = product1.UserRefProperties
    strParam7 = parameters9.CreateString("Page", "")
    strParam7.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters10 = product1.UserRefProperties
    strParam8 = parameters10.CreateString("L1", "")  # 形狀孔
    strParam8.ValuateFromString("L1: " + str(lower_die_set_machining_explanation_shape) + "-(.., 銑)")
    product1 = product1.ReferenceProduct
    strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
    strParam9.ValuateFromString("")
    product1 = product1.ReferenceProduct
    strParam10 = parameters10.CreateString("HP", "")  # 合銷孔
    strParam10.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters11 = product1.UserRefProperties
    strParam11 = parameters11.CreateString("B", "")  # B型引導沖孔
    strParam11.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters12 = product1.UserRefProperties
    strParam12 = parameters12.CreateString("BP", "")  # B沖沖孔
    strParam12.ValuateFromString("BP: " + str(0) + "- %%C" + str(
        4.98) + "(B沖沖孔)")  # 介面參數設定OP10定位銷直徑 #lower_die_set_machining_explanation_pilot_punch Working_parameter
    product1 = product1.ReferenceProduct
    parameters13 = product1.UserRefProperties
    strParam13 = parameters13.CreateString("TS", "")  # 浮升引導
    strParam13.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters14 = product1.UserRefProperties
    strParam14 = parameters14.CreateString("IG", "")  ##內導柱
    strParam14.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters15 = product1.UserRefProperties
    strParam15 = parameters15.CreateString("F", "")  # 外導柱
    strParam15.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters16 = product1.UserRefProperties
    strParam16 = parameters16.CreateString("CS", "")  # 等高套筒
    strParam16.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters17 = product1.UserRefProperties
    strParam17 = parameters17.CreateString("AP", "")  # A沖沖孔
    strParam17.ValuateFromString("")
    product1 = product1.ReferenceProduct
    # ====↑設定性質↑=====================================
    part1.Update()
    # --------------↑刪除不需要的Data↑--------------
    partDocument1.SaveAs(gvar.save_path + "lower_die_set.CATPart")  # 存檔的檔案名稱
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    partDocument1.Close()
    # -------------------------------------------------關閉挖孔件
    # ======================================(lower_die_set_change)
    return lower_die_set_length,lower_die_set_width


def boolean_hole_body(mode, lower_die_set_A_punch_number, lower_die_set_insert_hole_number, cut_line_st):
    catapp = win32.Dispatch('CATIA.Application')
    BodyName = 0  # 輸入須要移除的Body名稱
    PadName = 0  # 輸入須要投影的平面名稱
    if BodyName == 0:
        BodyName = "PartBody"
    if PadName == 0:
        PadName = "Pad.1"
    # ===============新增Body=========
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Add()
    part1.Update()
    # ===============投影=========
    hybridShapeFactory1 = part1.HybridShapeFactory
    parameters1 = part1.Parameters
    hybridShapeCurveExplicit1 = parameters1.Item(cut_line_st)  # 投影線段cut_line_st為須要挖洞的線段
    reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
    body2 = bodies1.Item(BodyName)
    shapes1 = body2.Shapes
    pad1 = shapes1.Item(PadName)
    reference2 = part1.CreateReferenceFromObject(pad1)
    hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference1, reference2)
    hybridShapeProject1.SolutionType = 0
    hybridShapeProject1.Normal = True
    hybridShapeProject1.SmoothingType = 0
    body1.InsertHybridShape(hybridShapeProject1)
    part1.InWorkObject = hybridShapeProject1
    part1.Update()
    # ===============長出=========
    shapeFactory1 = part1.ShapeFactory
    reference3 = part1.CreateReferenceFromObject(hybridShapeProject1)
    pad2 = shapeFactory1.AddNewPadFromRef(reference3, 40)
    relations1 = part1.Relations
    limit1 = pad2.FirstLimit
    length1 = limit1.dimension
    formula1 = relations1.Createformula("formula.346", "", length1, "plate_height ")
    formula1.rename("formula.346")
    pad2.IsSymmetric = True
    part1.Update()
    # ===============Remove=========
    part1.InWorkObject = body2
    remove1 = shapeFactory1.AddNewRemove(body1)
    if mode == "A_punch":
        lower_die_set_A_punch_number = lower_die_set_A_punch_number + 1
        remove1.Name = "Lower-die-set-A-punch-" + str(lower_die_set_A_punch_number)
    if mode == "insert":
        lower_die_set_insert_hole_number = lower_die_set_insert_hole_number + 1
        remove1.Name = "Lower-die-set-Insert-hole-" + str(lower_die_set_insert_hole_number)
    part1.Update()
    return lower_die_set_A_punch_number, lower_die_set_insert_hole_number

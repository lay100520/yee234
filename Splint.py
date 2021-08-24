import win32com.client as win32
import global_var as gvar
import defs
import PunchDef
import time

def Splint(now_plate_line_number):
    total_op_number = int(gvar.strip_parameter_list[2])
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument2 = catapp.ActiveDocument
    partDocument1 = documents1.Open(gvar.open_path + "Splint.CATPart")
    # ======================================
    defs.window_change(partDocument2, partDocument1)  # 複製catia檔案(留著的視窗,被複製的視窗)
    # ======================================
    # -------------------------------------------------------------↓補強入子
    selection3 = partDocument2.Selection
    selection3.Clear()
    selection3.Search("Name=plate_line_" + str(now_plate_line_number) + "_op*_Reinforcement_cut_line_*,all")
    if selection3.Count > 0:
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.open_path + "QR_punch_Reinforcement.CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("cut_line_formula_1")
        hybridShapeCurveExplicit1.Name = "cut_line_formula_3"
        # ======================================
        defs.window_change(partDocument2, partDocument1)  # 複製catia檔案(留著的視窗,被複製的視窗)
        # ======================================
        bodies1 = part1.Bodies
        body_number = bodies1.Count
        # ------------------------------------------------------------↓   改變body名稱
        body1 = bodies1.Item("Body." + str(body_number))
        body1.Name = "Reinforcement_cut_punch"
        # ------------------------------------------------------------↑
        selection3.Clear()
        selection3.Add(body1)
        selection3.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
        selection3.Clear()
    # ------------------------------------------------------------↑ 補強入子
    # ==================================================================================Splint_change.CATMain
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    length = [None] * 6
    formula = [None] * 6
    length[0] = part1.Parameters.Item("pitch")
    length[0].Value = int(gvar.strip_parameter_list[4])
    length[1] = part1.Parameters.Item("plate_height")
    g = now_plate_line_number
    try:
        if gvar.die_type == "module":
            length[1].Value = 0  # (back_splint_height)
        else:
            length[1].Value = float(gvar.strip_parameter_list[14])
    except:
        die_rule_file_name = "模板厚度選擇"
        excel_Sheet_name = "200T以下"
        Column_string_serch = "上夾板"
        Row_string_serch = "精密級"
        (serch_result) = defs.ExcelSearch(die_rule_file_name, excel_Sheet_name, Row_string_serch, Column_string_serch)
        length[1].Value = -serch_result
        ff = serch_result
        gvar.strip_parameter_list[14] = ff
    # ======================================================================================================
    length[4] = part1.Parameters.Item("plate_down_plane")
    if gvar.die_type == "module":
        plate_position = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[14]) + 0  # (back_stripper_plate_height)
    else:
        plate_position = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17])
    if gvar.Mold_status == "開模":
        length[4].Value = plate_position + 28  # (upper_die_open_height)
    elif gvar.Mold_status == "閉模":
        length[4].Value = plate_position
    # ======================================================================================================
    file_name = "Data1"
    body_name1 = "Body.2"
    hybridBody_name = "die"
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(file_name, body_name1,
                                                                                             hybridBody_name)  # 環境設定(檔案名,body名,依據群組名)
    if gvar.die_type == "module":
        M_plate_length = 35
        M_plate_wide = 116
        hybridShape1 = ElementBody.HybridShapes.Item("down_die_plate_up_plane")  # 宣告平面(下)
        (ElementSketch) = defs.material_type_palte_sketch(hybridShape1, M_plate_length, M_plate_wide,
                                                          100 + 50 * (now_plate_line_number - 1),
                                                          112.5 + 12.5, ElementDocument, ElementBody, ElementHybridBody,
                                                          "Body")  # (平面的宣告,模板長,模板寬,中心座標_X,中心座標_Y)
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\plate_size")  # 草圖置換
        part1.Update()
        M_splint_design(M_plate_length, M_plate_wide, ElementSketch, ElementDocument, ElementBody, ElementHybridBody,
                        "Body")
    else:
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\number_" + str(g) + "_plate_line")  # 草圖置換
    part1.Update()
    Splint_machining_explanation_shape = 0  # ------------------------------------------------------------加工說明
    for n in range(1, total_op_number + 1):
        now_op_number = n
        op_number = 10 * n
        splint_insert_hole_number = int()
        if gvar.StripDataList[38][g][n] > 0:
            for i in range(1, round(gvar.StripDataList[38][g][n]) + 1):
                Splint_machining_explanation_shape = Splint_machining_explanation_shape + 1  # ------------------------------------加工說明
                bodies1 = part1.Bodies
                body1 = bodies1.Item("Body.2")
                part1.InWorkObject = body1
                hybridShapeFactory1 = part1.HybridShapeFactory
                parameters1 = part1.Parameters
                hybridShapeCurveExplicit1 = parameters1.Item(
                    "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i))
                reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
                hybridShapes1 = body1.HybridShapes
                hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
                reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
                hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference1, reference2)
                hybridShapeProject1.SolutionType = 0
                hybridShapeProject1.Normal = True
                hybridShapeProject1.SmoothingType = 0
                body1.InsertHybridShape(hybridShapeProject1)
                part1.InWorkObject = hybridShapeProject1
                part1.Update()
                reference3 = part1.CreateReferenceFromObject(hybridShapeProject1)
                hybridShapeCurveExplicit2 = hybridShapeFactory1.AddNewCurveDatum(reference3)
                body1.InsertHybridShape(hybridShapeCurveExplicit2)
                part1.InWorkObject = hybridShapeCurveExplicit2
                part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(
                    i) + "_project_line"  # 更改外形線名稱
                hybridShapeFactory1.DeleteObjectForDatum(reference3)
                part1.Update()
                hybridShapeFactory1.DeleteObjectForDatum(reference3)
                shapeFactory1 = part1.ShapeFactory
                parameters1 = part1.Parameters
                hybridShapeCurveExplicit2 = parameters1.Item(
                    "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i) + "_project_line")  # 外型線
                reference4 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit2)
                pocket1 = shapeFactory1.AddNewPocketFromRef(reference4, 20)
                limit1 = pocket1.FirstLimit
                limit1.LimitMode = 3
                hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
                reference5 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
                limit1.LimitingElement = reference5
                splint_insert_hole_number = splint_insert_hole_number + 1
                pocket1.Name = "Splint-Insert-hole-" + str(splint_insert_hole_number)
                part1.Update()
                # ------------------------------------------------------
        if gvar.StripDataList[40][g][n] > 0:
            for i in range(1, round(gvar.StripDataList[40][g][n]) + 1):
                Splint_machining_explanation_shape = Splint_machining_explanation_shape + 1  # ------------------------------------加工說明
                # --------------Boundary 取得型面外形線↓---------------------
                hybridShapeFactory1 = part1.HybridShapeFactory
                parameters1 = part1.Parameters
                hybridShapeSurfaceExplicit1 = parameters1.Item(
                    "die\plate_line_" + str(g) + "_op" + str(op_number) + "_forming_cavity_surface_" + str(i))  # 型面名稱
                reference1 = part1.CreateReferenceFromObject(hybridShapeSurfaceExplicit1)
                hybridShapeBoundary1 = hybridShapeFactory1.AddNewBoundaryOfSurface(reference1)
                bodies1 = part1.Bodies
                body1 = bodies1.Item("Body.2")  # 設定工作物件為Body.2
                body1.InsertHybridShape(hybridShapeBoundary1)
                part1.InWorkObject = hybridShapeBoundary1
                reference2 = part1.CreateReferenceFromObject(hybridShapeBoundary1)
                hybridShapeCurveExplicit1 = hybridShapeFactory1.AddNewCurveDatum(reference2)
                body1.InsertHybridShape(hybridShapeCurveExplicit1)
                part1.InWorkObject = hybridShapeCurveExplicit1
                part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(
                    op_number) + "_forming_cavity_Boundary_" + str(i)  # 更改外形線名稱
                hybridShapeFactory1.DeleteObjectForDatum(reference2)
                # --------------隱藏Boundary型面外形線↓---------------------
                selection1 = partDocument1.Selection
                selection1.Clear()
                selection1.Search("Name=*_Boundary*, All ")
                selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
                selection1.Clear()
                # --------------隱藏Boundary型面外形線↑---------------------
                # --------------Boundary 取得型面外形線↑---------------------
                # ----------------Project 投影外形線至指定平面↓--------------------------------
                hybridShapeCurveExplicit1 = parameters1.Item(
                    "plate_line_" + str(g) + "_op" + str(op_number) + "_forming_cavity_Boundary_" + str(i))  # 欲投影外形線之名稱
                reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
                hybridShapes1 = body1.HybridShapes
                hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")  # 投影到哪個平面
                reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
                hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference1, reference2)
                hybridShapeProject1.SolutionType = 1  #
                hybridShapeProject1.Normal = True
                hybridShapeProject1.SmoothingType = 0
                body1.InsertHybridShape(hybridShapeProject1)
                part1.InWorkObject = hybridShapeProject1
                reference3 = part1.CreateReferenceFromObject(hybridShapeProject1)
                hybridShapeCurveExplicit2 = hybridShapeFactory1.AddNewCurveDatum(reference3)
                body1.InsertHybridShape(hybridShapeCurveExplicit2)

                part1.InWorkObject = hybridShapeCurveExplicit2
                part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(
                    op_number) + "_forming_cavity_Project_" + str(i)  # 更改外形線名稱
                hybridShapeFactory1.DeleteObjectForDatum(reference3)
                # ----------------Project 投影外形線至指定平面↑--------------------------------
                # ---------------------------------------
                shapeFactory1 = part1.ShapeFactory
                reference1 = part1.CreateReferenceFromName("")
                pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 50)
                limit1 = pocket1.FirstLimit
                limit1.LimitMode = 3
                parameters2 = part1.Parameters
                hybridShapeCurveExplicit2 = parameters2.Item(
                    "plate_line_" + str(g) + "_op" + str(op_number) + "_forming_cavity_Project_" + str(i))
                reference2 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit2)
                pocket1.SetProfileElement(reference2)
                reference3 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit2)
                pocket1.SetProfileElement(reference3)
                bodies1 = part1.Bodies
                body1 = bodies1.Item("Body.2")
                hybridShapes1 = body1.HybridShapes
                hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
                reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
                limit1.LimitingElement = reference2
                parameters1 = part1.Parameters
        if gvar.StripDataList[4][g][n] > 0:  # 補強入子
            for for_counter in range(1, round(gvar.StripDataList[4][g][n]) + 1):
                parameter_digital1 = 0
                PunchDef.punch_Reinforcement_Ecxavation(g, n, for_counter, parameter_digital1)
        if gvar.StripDataList[42][g][n] > 0:  # 整平模組孔down
            for j in range(1, gvar.StripDataList[42][g][n] + 1):
                pp_count = j
                PunchDef.bend_up_shaping_cavity_hole_1(op_number, pp_count)
        if gvar.StripDataList[43][g][n] > 0:  # 整平模組孔up
            for j in range(1, round(gvar.StripDataList[43][g][n]) + 1):
                pp_count = j
                PunchDef.bend_up_shaping_cavity_hole_2(op_number, pp_count)
        if gvar.StripDataList[27][g][n] > 0:  # 切斷沖頭_下
            PunchDef.punch_d_cutting(g, n)
        if gvar.StripDataList[28][g][n] > 0:  # 切斷沖頭_上
            PunchDef.punch_u_cutting(g, n)
        if gvar.StripDataList[37][g][n] > 0:  # A沖
            A_punch(g, n)
        if gvar.StripDataList[21][g][n] > 0:  # 打凸包沖頭_左
            PunchDef.emboss_forming_punch_left(g, n)
        if gvar.StripDataList[22][g][n] > 0:  # 打凸包沖頭_右
            PunchDef.emboss_forming_punch_right(g, n)
        if gvar.StripDataList[3][g][n] > 0:  # 半沖切
            PunchDef.QR_half_cut_punch(g, n)
        # ------------------------------------------------------------------------------↓快拆沖頭
    for n in range(1, total_op_number + 1):
        now_op_number = n
        op_number = 10 * n
        if gvar.StripDataList[29][g][n] > 0:  # '沖切沖頭_右
            data_type = "line"
            data_number = round(gvar.StripDataList[29][g][n])
            part_name = "op" + str(op_number) + "_right_quickly_remove_cut_punch_"
            PunchDef.quickly_remove_punch(data_type, data_number, part_name)
        if gvar.StripDataList[30][g][n] > 0:  # 沖切沖頭_左
            data_type = "line"
            data_number = round(gvar.StripDataList[30][g][n])
            part_name = "op" + str(op_number) + "_left_quickly_remove_cut_punch_"
            PunchDef.quickly_remove_punch(data_type, data_number, part_name)
        if gvar.StripDataList[31][g][n] > 0:  # 沖切沖頭_上
            data_type = "line"
            data_number = gvar.StripDataList[31][g][n]
            part_name = "op" + str(op_number) + "_up_quickly_remove_cut_punch_"
            PunchDef.quickly_remove_punch(data_type, data_number, part_name)
        if gvar.StripDataList[32][g][n] > 0:  # 沖切沖頭_下
            data_type = "line"
            data_number = gvar.StripDataList[32][g][n]
            part_name = "op" + str(op_number) + "_down_quickly_remove_cut_punch_"
            PunchDef.quickly_remove_punch(data_type, data_number, part_name)
        if gvar.StripDataList[33][g][n] > 0:  # 折彎沖頭_右
            data_type = "surface"
            data_number = gvar.StripDataList[33][g][n]
            part_name = "op" + str(op_number) + "_right_quickly_remove_bending_punch_"
            PunchDef.quickly_remove_punch(data_type, data_number, part_name)
        if gvar.StripDataList[34][g][n] > 0:  # 折彎沖頭_左
            data_type = "surface"
            data_number = gvar.StripDataList[34][g][n]
            part_name = "op" + str(op_number) + "_left_quickly_remove_bending_punch_"
            PunchDef.quickly_remove_punch(data_type, data_number, part_name)
        if gvar.StripDataList[35][g][n] > 0:  # 折彎沖頭_上
            data_type = "surface"
            data_number = gvar.StripDataList[35][g][n]
            part_name = "op" + str(op_number) + "_up_quickly_remove_bending_punch_"
            PunchDef.quickly_remove_punch(data_type, data_number, part_name)
        if gvar.StripDataList[36][g][n] > 0:  # 折彎沖頭_下
            data_type = "surface"
            data_number = gvar.StripDataList[36][g][n]
            part_name = "op" + str(op_number) + "_down_quickly_remove_bending_punch_"
            PunchDef.quickly_remove_punch(data_type, data_number, part_name)
        # ------------------------------------------------------------------------------↑快拆沖頭
    part1.Update()
    product1.PartNumber = "Splint_" + str(g)
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Splint_" + str(g))
    parameters3 = product1.UserRefProperties
    strParam1 = parameters3.CreateString("NO.", "")
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
    strParam4.ValuateFromString(gvar.strip_parameter_list[15])
    product1 = product1.ReferenceProduct
    parameters7 = product1.UserRefProperties
    strParam5 = parameters7.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(gvar.strip_parameter_list[16])
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
    strParam8.ValuateFromString("P1: " + str(Splint_machining_explanation_shape) + "-(沖頭, 割), 單+0.005")
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
    strParam12.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters13 = product1.UserRefProperties
    strParam13 = parameters13.CreateString("TS", "")  # 浮升引導
    strParam13.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters14 = product1.UserRefProperties
    strParam14 = parameters14.CreateString("IG", "")  # 內導柱
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
    # --------------↓刪除不需要的Data↓--------------
    selection1 = partDocument1.Selection

    if now_plate_line_number == 1:
        selection1.Clear()
        selection1.Search("Name=plate_line_2*,all")
        try:
            selection1.Delete()
            selection1.Clear()
        except:
            pass
    if now_plate_line_number == 2:
        selection1.Clear()
        selection1 = partDocument1.Selection
        selection1.Search("Name=plate_line_1*,all")
        try:
            selection1.Delete()
            selection1.Clear()
        except:
            pass
    # --------------↑刪除不需要的Data↑--------------
    part1.Update()
    time.sleep(2)
    partDocument1.SaveAs(gvar.save_path + "Splint_" + str(g) + ".CATPart")  # 存檔的檔案名稱
    time.sleep(2)
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    partDocument1.Close()
    # ==================================================================================Splint_change.CATMain


def M_splint_design(M_plate_length, M_plate_wide, ElementSketch, ElementDocument, ElementBody, ElementHybridBody,
                    SketchPosition):
    (ElementReference1) = defs.ExtremumPoint("X_min", "Y_min", "Z_max", 2, ElementSketch, ElementDocument, ElementBody,
                                             ElementHybridBody)
    ElementReference2 = ElementReference1
    ElementReference2.Name = "up_common_seat_plate_min"  # 建立最小點
    part1 = ElementDocument.Part
    guild_D = 10.5
    L = 0  # (back_splint_height)
    for up_pin in range(1, 3):
        x_value = M_plate_length / 2
        if up_pin == 1:
            y_value = 34
        else:
            y_value = M_plate_wide - 34
        (ElementPoint5) = defs.BuildPoint(ElementReference2, ElementDocument, ElementBody, ElementHybridBody,
                                          SketchPosition)
        ElementPoint5.X.Value = x_value
        ElementPoint5.Y.Value = y_value
        ElementPoint5.Z.Value = gvar.strip_parameter_list[14]
        ElementPoint5.Name = "inner_guild_" + str(up_pin) + "_Start"
        ElementReference10 = ElementPoint5
        (ElementPoint5) = defs.BuildPoint(ElementReference10, ElementDocument, ElementBody, ElementHybridBody,
                                          SketchPosition)
        ElementPoint5.X.Value = 0
        ElementPoint5.Y.Value = 0
        hybridShape1 = ElementBody.HybridShapes.Item("down_die_plate_up_plane")  # 宣告平面(下)
        hybridShape2 = ElementBody.HybridShapes.Item("down_die_plate_down_plane")  # 宣告平面(上)
        ElementReference11 = ElementPoint5
        ElementReference12 = hybridShape2
        ElementReference10 = ElementPoint5
        (ElementPoint5) = defs.BuildPoint(ElementReference10, ElementDocument, ElementBody, ElementHybridBody,
                                          SketchPosition)
        ElementPoint5.Z.Value = 0  # (back_splint_height)
        ElementPoint5.Y.Value = 0
        ElementPoint5.X.Value = 0
        ElementPoint5.Name = "Blot_guide_hole_" + str(up_pin) + "_End"
        (hole_pin) = defs.HoleSimpleD(guild_D, L, 0, ElementDocument, ElementBody,
                                      ElementReference11, ElementReference12)
        if up_pin == 1:
            hole_pin.Name = "guide_hole_12_R"
        else:
            hole_pin.Name = "guide_hole_12_L"
        ElementReference10 = ElementReference2
        defs.BuildSketch("E_sketch", hybridShape2, ElementDocument, SketchPosition, ElementBody,
                         ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
        x_value = 17.5
        y_value = M_plate_wide / 2
        (ElementPoint5) = defs.BuildPoint(ElementReference10, ElementDocument, ElementBody, ElementHybridBody,
                                          SketchPosition)  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
        ElementPoint5.X.Value = x_value
        ElementPoint5.Y.Value = y_value
        (ElementPoint1,element_line1, element_line2, element_line3, element_line4) = defs.SketchRectangle(ElementSketch, M_plate_length, 22, ElementDocument,
                                               ElementHybridBody)  # (草圖陳述句,長,寬)  element_point(1)=>output中心點之陳述句  [需經環境副程式跑過]
        ElementPoint3 = ElementPoint5
        ElementPoint4 = ElementPoint1
        part1.Update()
        defs.SketchBuildCallout(ElementSketch, "free", "Binding", 0, ElementDocument, ElementPoint3,
                                ElementPoint4)  # 建立標註(草圖陳述句,標註方向,式標OR拘束,改變OR讀取之數值,標註兩點)
        # ---------------------------line_Excavation(挖除線段)
        part1 = ElementDocument.Part
        shapeFactory1 = part1.ShapeFactory
        reference1 = part1.CreateReferenceFromObject(ElementSketch)
        part1.InWorkObject = ElementBody
        pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
        ElementReference20 = pocket1
        # ---------------------------line_Excavation(挖除線段)
        ElementReference20.FirstLimit.dimension.Value = 10
        element_point = [None] * 20
        for triangle_number in range(1, 3):
            for point_N in range(6, 9):
                if point_N == 6 and triangle_number == 1:
                    x_value = 21
                    y_value = M_plate_wide / 2
                elif point_N == 7 and triangle_number == 1:
                    x_value = 0
                    y_value = M_plate_wide / 2 - 21
                elif point_N == 8 and triangle_number == 1:
                    x_value = 0
                    y_value = M_plate_wide / 2 + 21
                elif point_N == 6 and triangle_number == 2:
                    x_value = M_plate_length - 21
                    y_value = M_plate_wide / 2
                elif point_N == 7 and triangle_number == 2:
                    x_value = M_plate_length
                    y_value = M_plate_wide / 2 - 21
                elif point_N == 8 and triangle_number == 2:
                    x_value = M_plate_length
                    y_value = M_plate_wide / 2 + 21
                (ElementPoint5) = defs.BuildPoint(ElementReference10, ElementDocument, ElementBody, ElementHybridBody,
                                                  SketchPosition)  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
                ElementPoint5.X.Value = x_value
                ElementPoint5.Y.Value = y_value
                element_point[point_N] = ElementPoint5
            (ElementSketch) = defs.BuildSketch("E_sketch", hybridShape2, ElementDocument, SketchPosition, ElementBody,
                                               ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
            # ========================sketch_triangle(main_sketch)
            part1 = ElementDocument.Part
            part1.InWorkObject = ElementSketch
            factory2D1 = ElementSketch.OpenEdition()
            geometricElements1 = ElementSketch.GeometricElements
            axis2D1 = geometricElements1.Item("AbsoluteAxis")
            line2D5 = axis2D1.getItem("HDirection")
            line2D6 = axis2D1.getItem("VDirection")
            point2D1 = factory2D1.CreatePoint(50, 60)
            point2D2 = factory2D1.CreatePoint(50, 50)
            point2D3 = factory2D1.CreatePoint(60, 50)
            line2D1 = factory2D1.CreateLine(50, 60, 50, 50)
            line2D1.StartPoint = point2D1
            line2D1.EndPoint = point2D2
            line2D2 = factory2D1.CreateLine(50, 50, 60, 50)
            line2D2.EndPoint = point2D2
            line2D2.StartPoint = point2D3
            line2D3 = factory2D1.CreateLine(60, 50, 50, 60)
            line2D3.StartPoint = point2D3
            line2D3.EndPoint = point2D1
            part1.Update()
            ElementPoint11 = point2D1
            ElementPoint12 = point2D2
            ElementPoint13 = point2D3
            ElementSketch.CloseEdition()
            # ========================sketch_triangle(element_sketch(1))
            for point_N in range(15, 18):
                ElementPoint18 = element_point[point_N - 9]
                ElementPoint19 = element_point[point_N - 4]
                part1.Update()
                defs.SketchBuildCallout(ElementSketch, "free", "Binding", 0, ElementDocument, ElementPoint18,
                                        ElementPoint19)  # 建立標註(草圖陳述句,標註方向,式標OR拘束,改變OR讀取之數值,標註兩點)
            # ---------------------------line_Excavation(挖除線段)
            part1 = ElementDocument.Part
            shapeFactory1 = part1.ShapeFactory
            reference1 = part1.CreateReferenceFromObject(ElementSketch)
            part1.InWorkObject = ElementBody
            pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
            ElementReference20 = pocket1
            # ---------------------------line_Excavation(挖除線段)
            ElementReference20.FirstLimit.dimension.Value = 10


def A_punch(g, n):  # A沖
    catapp = win32.Dispatch('CATIA.Application')
    partDocument2 = catapp.ActiveDocument
    product1 = partDocument2.getItem("Part1")
    part1 = partDocument2.Part
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.open_path + "QR_Stop_line.CATPart")
    # 在CATIA上切換各視窗
    # ======================================
    defs.window_change(partDocument2, partDocument1)
    # ======================================
    length = [None] * 99
    partDocument1 = documents1.Open(gvar.open_path + "SJAS.CATPart")#(A_Punch_Module+.CATPart)
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length[2] = part1.Parameters.Item("D")
    length[2].Value = int(gvar.strip_parameter_list[23])
    length[3] = part1.Parameters.Item("H")
    partDocument1.Close()
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length[0] = part1.Parameters.Item("D")
    length[0].Value = int(gvar.strip_parameter_list[23])
    length[1] = part1.Parameters.Item("H")
    length[1].Value = length[3].Value
    op_number = n * 10
    cut_cavity_machining_explanation_shape = int()
    stop_plate_A_punch_number = int()
    part1.Update()
    for i in range(1, round(gvar.StripDataList[37][g][n]) + 1):
        cut_cavity_machining_explanation_shape = cut_cavity_machining_explanation_shape + 1  # --------------------------------------------------加工說明
        part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
            "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i))  # 草圖置換
        part1.Update()
        for B_n in range(50, 0, -1):
            selection1 = partDocument1.Selection
            selection1.Clear()
            selection1.Search("Name=Body." + str(B_n) + "*,all")
            Body_n = selection1.Count
            selection1.Clear()
            if Body_n > 0:
                body_number = B_n
                break
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        hybridShapeFactory1 = part1.HybridShapeFactory
        body2 = bodies1.Item("Body." + str(body_number))
        sketches1 = body2.Sketches
        sketch1 = sketches1.Item("A_punch_insert_hole_line_Sketch")
        reference1 = part1.CreateReferenceFromObject(sketch1)
        hybridShapes1 = body1.HybridShapes
        hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
        reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
        hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference1, reference2)
        hybridShapeProject1.SolutionType = 0
        hybridShapeProject1.Normal = True
        hybridShapeProject1.SmoothingType = 0
        body1.InsertHybridShape(hybridShapeProject1)
        part1.InWorkObject = hybridShapeProject1
        part1.Update()
        reference3 = part1.CreateReferenceFromObject(hybridShapeProject1)
        hybridShapeCurveExplicit1 = hybridShapeFactory1.AddNewCurveDatum(reference3)
        body1.InsertHybridShape(hybridShapeCurveExplicit1)
        part1.InWorkObject = hybridShapeCurveExplicit1
        part1.Update()
        hybridShapeFactory1.DeleteObjectforDatum(reference3)
        shapeFactory1 = part1.ShapeFactory
        reference4 = part1.CreateReferenceFromName("")
        pocket1 = shapeFactory1.AddNewPocketFromRef(reference4, 20)
        reference5 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        pocket1.SetProfileElement(reference5)
        reference6 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        pocket1.SetProfileElement(reference6)
        limit1 = pocket1.FirstLimit
        limit1.LimitMode = 2
        stop_plate_A_punch_number = stop_plate_A_punch_number + 1
        pocket1.Name = "Stop-plate-A-punch-" + str(stop_plate_A_punch_number)
        part1.Update()
        # ===============================================挖孔
    # --------------------------------------------------------------  隱藏線
    selection2 = partDocument1.Selection
    selection2.Clear()
    selection2.Search("Name=Point.*, All ")
    selection2.VisProperties.SetShow(1)  # 1為隱藏,0為顯示 線
    selection2.Clear()
    selection2.Search("Name=Sketch.*, All ")
    selection2.VisProperties.SetShow(1)
    selection2.Clear()
    # ----------------------------------------------------------------------------
    selection2.Search("Name=cut_line_assume")
    selection2.Delete()
    selection2.Clear()
    selection2.Search("Name=Body." + str(body_number))
    selection2.Delete()
    selection2.Clear()

# import csv
# with open(gvar.strip_parameters_file_root) as csvFile:
#     rows = csv.reader(csvFile)
#     strip_parameter_list = tuple(tuple(rows)[0])
#     gvar.strip_parameter_list = strip_parameter_list
# gvar.die_type = "common"
# gvar.StripDataList[37][1][1] = 2
# gvar.StripDataList[38][1][2] = 1
# gvar.StripDataList[38][1][3] = 4
# gvar.StripDataList[38][1][4] = 1
# gvar.StripDataList[38][1][6] = 4
# gvar.StripDataList[38][1][7] = 1
# catapp = win32.Dispatch('CATIA.Application')
# documents1 = catapp.Documents
# partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
# Splint(1)
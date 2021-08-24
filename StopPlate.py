import win32com.client as win32
import global_var as gvar
import defs
import PunchDef
import time

def StopPlate(now_plate_line_number):
    gvar.die_type = "common"
    total_op_number = int(gvar.strip_parameter_list[2])
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument2 = catapp.ActiveDocument
    partDocument1 = documents1.Open(gvar.open_path + "stop_plate.CATPart")
    defs.window_change(partDocument2, partDocument1)
    selection3 = partDocument2.Selection
    selection3.Clear()
    selection3.Search("Name=plate_line_" + str(now_plate_line_number) + "_op*_Reinforcement_cut_line_*,all")
    if selection3.Count > 0:
        catapp = win32.Dispatch('CATIA.Application')
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.open_path + "QR_punch_Reinforcement.CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("cut_line_formula_1")
        hybridShapeCurveExplicit1.Name = "cut_line_formula_3"
        # 在CATIA上切換各視窗
        defs.window_change(partDocument2, partDocument1)
        part2 = partDocument2.Part
        bodies1 = part2.Bodies
        body_number = bodies1.Count
        # ------------------------------------------------------------↓   改變body名稱
        body1 = bodies1.Item("Body." + str(body_number))
        body1.Name = "Reinforcement_cut_punch"
        # ------------------------------------------------------------↑
        selection3.Clear()
        selection3.Add(body1)
        selection3.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
        selection3.Clear()
        # --------------------------------另一個檔案複製-------------------------------------
        catapp = win32.Dispatch('CATIA.Application')
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.open_path + "cut_Reinforcement_insert.CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("cut_line_assume_1")
        hybridShapeCurveExplicit1.Name = "cut_line_formula_4"
        bodies1 = part1.Bodies
        body1 = bodies1.Item("PartBody")
        body1.Name = "Reinforcement_insert"
        # 在CATIA上切換各視窗
        defs.window_change(partDocument2, partDocument1)
        bodies1 = part2.Bodies
        # ------------------------------------------------------------↓   改變body名稱
        body1 = bodies1.Item("Reinforcement_insert")
        # ------------------------------------------------------------↑
        selection3.Clear()
        selection3.Add(body1)
        selection3.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
        selection3.Clear()
    # =========================================================================stop_plate_change
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    length = [None] * 7
    formula = [None] * 7
    # =====================================================================
    length[0] = part1.Parameters.Item("pitch")
    length[0].Value = int(gvar.strip_parameter_list[4])
    # =====================================================================
    length[1] = part1.Parameters.Item("plate_height")
    g = now_plate_line_number
    try:
        if gvar.die_type == "module":
            length[1].Value = float(gvar.strip_parameter_list[14])
        else:
            length[1].Value = float(gvar.strip_parameter_list[17])
    except:
        die_rule_file_name = "模板厚度選擇"
        excel_Sheet_name = "200T以下"
        Column_string_serch = "止擋板"
        Row_string_serch = "精密級"
        (serch_result) = defs.ExcelSearch(die_rule_file_name, excel_Sheet_name, Row_string_serch, Column_string_serch)
        length[1].Value = serch_result
        ee = serch_result
        gvar.strip_parameter_list[17] = ee
    length[4] = part1.Parameters.Item("plate_down_plane")
    plate_position = float(float(gvar.strip_parameter_list[1]) + int(gvar.strip_parameter_list[20]) + 0)
    if gvar.Mold_status == "開模":
        length[4].Value = plate_position + 28  # (upper_die_open_height)
    else:
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
        M_stop_design(M_plate_length, M_plate_wide, ElementSketch, ElementDocument, ElementBody, ElementHybridBody,
                      "Body")
    else:
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\number_" + str(g) + "_plate_line")  # 草圖置換
    part1.Update()
    specsAndGeomWindow1 = catapp.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    part1.Update()
    stop_plate_machining_explanation_shape = 0
    stop_plate_insert_hole_number = 0
    for n in range(1, total_op_number + 1):
        now_op_number = n
        op_number = 10 * n
        if round(gvar.StripDataList[38][g][n]) > 0:
            for i in range(1, round(gvar.StripDataList[38][g][n]) + 1):
                stop_plate_machining_explanation_shape = stop_plate_machining_explanation_shape + 1
                partDocument1 = catapp.ActiveDocument
                part1 = partDocument1.Part
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
                selection1 = partDocument1.Selection
                selection1.Clear()
                selection1.Search("Name=*_project_line,all")
                selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
                selection1.Clear()
                # ----------------------------------↑投影外型線↑--------------------------------------
                # ----------------------------------↓OFFSET↓------------------------------------------
                reference4 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit2)
                hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)  # 改Z方向
                hybridShape3DCurveOffset1 = hybridShapeFactory1.AddNew3DCurveOffset(reference4, hybridShapeDirection1,
                                                                                    0, 0.5, 1)  # 偏移距離
                hybridShape3DCurveOffset1.InvertDirection = False
                body1.InsertHybridShape(hybridShape3DCurveOffset1)
                part1.InWorkObject = hybridShape3DCurveOffset1
                part1.Update()
                reference5 = part1.CreateReferenceFromObject(hybridShape3DCurveOffset1)
                hybridShapeCurveExplicit3 = hybridShapeFactory1.AddNewCurveDatum(reference5)
                body1.InsertHybridShape(hybridShapeCurveExplicit3)
                part1.InWorkObject = hybridShapeCurveExplicit3
                part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(
                    i) + "_offset_line"
                part1.Update()
                hybridShapeFactory1.DeleteObjectForDatum(reference5)
                shapeFactory1 = part1.ShapeFactory
                reference6 = part1.CreateReferenceFromName("")
                pocket1 = shapeFactory1.AddNewPocketFromRef(reference6, 20)
                limit1 = pocket1.FirstLimit
                limit1.LimitMode = 3
                reference7 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit3)
                pocket1.SetProfileElement(reference7)
                reference8 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit3)
                pocket1.SetProfileElement(reference8)
                limit1.LimitMode = 3
                hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
                reference9 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
                limit1.LimitingElement = reference9
                stop_plate_insert_hole_number = stop_plate_insert_hole_number + 1
                pocket1.Name = "Stop-plate-Insert-hole" + str(stop_plate_insert_hole_number)
                part1.Update()
                # ------------------------------------------------------
        if round(gvar.StripDataList[40][g][n]) > 0:
            for i in range(1, round(gvar.StripDataList[40][g][n]) + 1):
                stop_plate_machining_explanation_shape = stop_plate_machining_explanation_shape + 1
                # --------------Boundary 取得型面外形線↓---------------------
                hybridShapeFactory1 = part1.HybridShapeFactory
                parameters1 = part1.Parameters
                hybridShapeSurfaceExplicit1 = parameters1.Item(
                    "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_forming_cavity_surface_" + str(i))  # 型面名稱
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
                hybridShapeProject1.SolutionType = 1  ###
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
                selection3 = partDocument1.Selection
                selection3.Clear()
                selection3.Search("Name=*_Project*, All ")
                selection3.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
                selection3.Clear()
                # ----------------Project 投影外形線至指定平面↑--------------------------------
                # ----------------------------------↓OFFSET↓------------------------------------------
                reference11 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit2)
                hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)  # 改Z方向
                hybridShape3DCurveOffset4 = hybridShapeFactory1.AddNew3DCurveOffset(reference11, hybridShapeDirection2,
                                                                                    0, 1, 0.5)  # 偏移距離
                hybridShape3DCurveOffset4.InvertDirection = False
                body1.InsertHybridShape(hybridShape3DCurveOffset4)
                part1.InWorkObject = hybridShape3DCurveOffset4
                part1.Update()
                reference15 = part1.CreateReferenceFromObject(hybridShape3DCurveOffset4)
                hybridShapeCurveExplicit9 = hybridShapeFactory1.AddNewCurveDatum(reference15)
                body1.InsertHybridShape(hybridShapeCurveExplicit9)
                part1.InWorkObject = hybridShapeCurveExplicit9
                part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(
                    op_number) + "_forming_cavity_offset_" + str(i)
                part1.Update()
                hybridShapeFactory1.DeleteObjectForDatum(reference15)
                selection2 = partDocument1.Selection
                selection2.Clear()
                selection2.Search("Name=*_offset*, All ")
                selection2.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
                selection2.Clear()
                # ----------------------------------↑OFFSET↑------------------------------------------
                # ---------------------------------------
                shapeFactory1 = part1.ShapeFactory
                reference1 = part1.CreateReferenceFromName("")
                pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 50)
                limit1 = pocket1.FirstLimit
                limit1.LimitMode = 3
                parameters2 = part1.Parameters
                hybridShapeCurveExplicit2 = parameters2.Item(
                    "plate_line_" + str(g) + "_op" + str(op_number) + "_forming_cavity_offset_" + str(i))
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
        if gvar.StripDataList[42][g][n] > 0:  # 整平模組孔down
            for j in range(1, gvar.StripDataList[42][g][n] + 1):
                pp_count = j
                PunchDef.bend_up_shaping_cavity_hole_1(op_number, pp_count)
        if gvar.StripDataList[43][g][n] > 0:  # 整平模組孔up
            for j in range(1, round(gvar.StripDataList[43][g][n]) + 1):
                pp_count = j
                PunchDef.bend_up_shaping_cavity_hole_2(op_number, pp_count)
        if gvar.StripDataList[4][g][n] > 0:  # 補強入子
            for for_counter in range(1, round(gvar.StripDataList[4][g][n]) + 1):
                sketch_name2 = "Excavation_offset_Sketch"
                PunchDef.punch_Reinforcement_Ecxavation(g, n, for_counter, 0)  # (stop_plate_space)
                Reinforcement_insert_Bolt_Hole(g, n, for_counter)
        if gvar.StripDataList[27][g][n] > 0:  # 切斷沖頭_下
            PunchDef.punch_d_cutting(g, n)
        if gvar.StripDataList[28][g][n] > 0:  # 切斷沖頭_上
            PunchDef.punch_u_cutting(g, n)
        if gvar.StripDataList[37][g][n] > 0:  # A沖
            A_punch(g, n)
        if gvar.StripDataList[21][g][n] > 0:  # 打凸包沖頭_左
            emboss_forming_punch_left(g, n)
        if gvar.StripDataList[22][g][n] > 0:  # 打凸包沖頭_右
            PunchDef.emboss_forming_punch_right(g, n)
        if gvar.StripDataList[3][g][n] > 0:  # 半沖切
            PunchDef.QR_half_cut_punch(g, n)
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
        if gvar.StripDataList[73][g][n] > 0:  # 整形沖頭
            for now_data_number in range(1, gvar.StripDataList[73][g][n]):
                F_bending_stop_plate(g, op_number, now_data_number, "Body")
    part1.Update()
    product1.PartNumber = "Stop_plate_" + str(g)  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Stop_plate_" + str(g))
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
    strParam4.ValuateFromString(gvar.strip_parameter_list[18])
    product1 = product1.ReferenceProduct
    parameters7 = product1.UserRefProperties
    strParam5 = parameters7.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(gvar.strip_parameter_list[19])
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
    strParam8.ValuateFromString("S1: " + str(stop_plate_machining_explanation_shape) + "-(沖頭孔,銑)")
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
    strParam16 = parameters15.CreateString("CS", "")  # 等高套筒
    strParam16.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters17 = product1.UserRefProperties
    strParam17 = parameters17.CreateString("AP", "")  # A沖沖孔
    strParam17.ValuateFromString("")
    product1 = product1.ReferenceProduct
    # ====↑設定性質↑=====================================
    part1.Update()
    selection1 = partDocument1.Selection
    if now_plate_line_number == 1:
        selection1.Clear()
        selection1.Search("Name=plate_line_2*,all")
        if selection1.Count > 0:
            selection1.Delete()
        selection1.Clear()
    elif now_plate_line_number == 2:
        selection1.Clear()
        selection1 = partDocument1.Selection
        selection1.Search("Name=plate_line_1*,all")
        if selection1.Count > 0:
            selection1.Delete()
        selection1.Clear()
    # --------------↑刪除不需要的Data↑--------------
    part1.Update()
    time.sleep(2)
    partDocument1.SaveAs(gvar.save_path + "Stop_plate_" + str(g) + ".CATPart")  # 存檔的檔案名稱
    time.sleep(2)
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    partDocument1.Close()
    # =========================================================================stop_plate_change


def M_stop_design(M_plate_length, M_plate_wide, ElementSketch, ElementDocument, ElementBody, ElementHybridBody,
                  SketchPosition):
    (ElementReference1) = defs.ExtremumPoint("X_min", "Y_min", "Z_max", 2, ElementSketch, ElementDocument, ElementBody,
                                             ElementHybridBody)
    ElementReference2 = ElementReference1
    ElementReference2.Name = "up_common_seat_plate_min"  # 建立最小點
    ElementPoint3 = ElementReference2
    for up_pin in range(1, 3):
        x_value = M_plate_length / 2
        if up_pin == 1:
            y_value = 13
        else:
            y_value = M_plate_wide - 13
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
        ElementPoint5.Z.Value = -5
        ElementPoint5.Name = "inner_guild_" + str(up_pin) + "_End"
        hybridShape1 = ElementBody.HybridShapes.Item("down_die_plate_up_plane")  # 宣告平面(下)
        hybridShape2 = ElementBody.HybridShapes.Item("down_die_plate_down_plane")  # 宣告平面(上)
        ElementReference11 = ElementPoint5
        ElementReference12 = hybridShape2
        (hole_pin) = defs.HoleSimpleD(13.5, 5, 0, ElementDocument, ElementBody, ElementReference11, ElementReference12)
        if up_pin == 1:
            hole_pin.Name = "inner_guild_R_1"
        else:
            hole_pin.Name = "inner_guild_L_1"
        (hole_pin) = defs.HoleSimpleD(10, int(gvar.strip_parameter_list[14]), 0, ElementDocument, ElementBody,
                                      ElementReference11, ElementReference12)
        if up_pin == 1:
            hole_pin.Name = "inner_guild_R_2"
        else:
            hole_pin.Name = "inner_guild_L_2"
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
        ElementReference11 = ElementPoint5
        hybridShape2 = ElementBody.HybridShapes.Item("down_die_plate_down_plane")  # 宣告平面(上)
        ElementReference12 = hybridShape2
        (hole_pin) = defs.HoleSimpleD(10, int(gvar.strip_parameter_list[14]), 0, ElementDocument, ElementBody,
                                      ElementReference11, ElementReference12)
        if up_pin == 1:
            hole_pin.Name = "pin_hole_10_R_2"
        else:
            hole_pin.Name = "pin_hole_10_L_2"


def stop_plate_point_build(g, point, high, body_name1, body_name2):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    parameters1 = part1.Parameters
    bodies1 = part1.Bodies
    body1 = bodies1.Item(body_name1)
    body2 = bodies1.Item(body_name2)
    point_conter = [] * 30
    stop_plate_Bolt_Hole = [] * 10
    # '------------------------------------------↓搜尋現有pad_Bolt_Hole點的螺栓數量
    selection1.Clear()
    selection1.Search("Name=stop_plate_" + str(g) + "_Bolt_point_*,all")
    point_conter[g] = selection1.Count
    selection1.Clear()
    # '------------------------------------------↑搜尋現有pad_Bolt_Hole點的螺栓數量
    # '------------------------------------------------------------↓   零建檔建立點之語法"HybridShapes"包涵在body裡
    hybridShapes2 = body2.HybridShapes
    hybridShapePointCoord1 = hybridShapes2.Item(point)
    reference2 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
    # '------------------------------------------------------------↑
    # '------------------------------------------------------------↓   建立點
    hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, high)  # '(x,y,z)
    hybridShapePointCoord1.PtRef = reference2  # '基礎點 = reference2
    body1.InsertHybridShape(hybridShapePointCoord1)  # '建立點
    part1.Update()
    # '------------------------------------------------------------↑
    reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
    point_conter[g] = point_conter[g] + 1
    stop_plate_Bolt_Hole[g] = point_conter[g]
    # '------------------------------------------------------------↓   建立點(打斷關連的)
    hybridShapePointExplicit1 = hybridShapeFactory1.AddNewPointDatum(reference3)  # '將點置於元素3
    body1.InsertHybridShape(hybridShapePointExplicit1)
    part1.InWorkObject = hybridShapePointExplicit1
    hybridShapePointExplicit1.Name = "stop_plate_" + str(g) + "_Bolt_point_" + str(point_conter[g])  # '點改名字
    point_name = "stop_plate_" + str(g) + "_Bolt_point_" + str(point_conter[g])  # '點名字輸出
    part1.Update()
    hybridShapeFactory1.DeleteObjectForDatum(reference3)  # '刪除元素3
    selection1.Add(hybridShapePointExplicit1)
    visPropertySet1 = visPropertySet1.Parent
    visPropertySet1.SetShow(1)
    selection1.Clear()
    # '------------------------------------------------------------↑
    return point_name


def F_Hole_up(n, i, hole_digital, hole_high, plane, point_name, body_name1):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    shapeFactory1 = part1.ShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item(body_name1)
    hybridShapeFactory1 = part1.HybridShapeFactory
    parameters1 = part1.Parameters
    # '------------------------------------------------------------↓   在已有平面建立條件             (方法二)
    hybridShapes1 = body1.HybridShapes
    hybridShapePlaneOffset1 = hybridShapes1.Item(plane)
    reference1 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
    # '------------------------------------------------------------↑
    hybridShapePointExplicit1 = parameters1.Item(point_name)
    reference2 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
    hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference2, reference1, 15)
    hole1.Name = "Hole_OP" + str(n * 10) + "_pick_" + str(i) + "_" + str(hole_digital)  # 孔改名字
    hole1.Type = 0
    hole1.AnchorMode = 0
    hole1.BottomType = 0
    # '------------------------------------------------------------↓     孔的型態設定
    hole1.ThreadingMode = 0  # (螺紋孔catThreadedHoleThreading  OR 無螺紋孔catSmoothHoleThreading)
    hole1.ThreadSide = 0  # 左OR右螺紋
    hole1.Type = 0  # 直孔  (沉頭孔令外再用草圖挖)
    # '------------------------------------------------------------↑
    limit1 = hole1.BottomLimit
    limit1.LimitMode = 0
    length1 = hole1.Diameter
    hole1.CreateStandardThreadDesignTable(1)
    strParam1 = hole1.HoleThreadDescription
    strParam1.Value = "M8"
    # '=============================================↓   極限設定
    limit2 = hole1.BottomLimit
    limit2.LimitMode = 0  # 未貫穿(盲孔)
    # '=============================================↑
    # '=============================================↓   牙深設定 孔深>牙深(否則錯誤)
    length2 = limit2.dimension  # 孔深
    length2.Value = hole_high
    # '=============================================↑
    # '=============================================↓   牙深設定 孔深>牙深(否則錯誤)
    length3 = hole1.ThreadDepth  # 牙深
    length3.Value = length2.Value - 2
    # '=============================================↑
    sketch1 = hole1.sketch
    selection1.Add(sketch1)
    visPropertySet1 = visPropertySet1.Parent
    visPropertySet1.SetShow(1)
    selection1.Clear()
    hole1.Reverse()
    part1.UpdateObject(hole1)


def Reinforcement_insert_Bolt_Hole(g, n, i):
    formula_name1 = "cut_line_formula_4"
    line_name2 = "plate_line_" + str(g) + "_op" + str(n * 10) + "_Reinforcement_cut_line_" + str(i)
    body_name1 = "Body.2"
    body_name2 = "Reinforcement_insert"
    sketch_name1 = "Bolt_Sketch"
    element_name1 = "down_die_plate_up_plane"
    point_name2 = " "
    line_name1 = "plate_line_" + str(g) + "_op" + str(n * 10) + "_Reinforcement_cut_line_" + str(i)
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    parameters2 = parameters1.RootParameterSet.ParameterSets.Item("Reinforcement_Parameters")
    defs.FormulaChange(body_name2, formula_name1, line_name1,
                       sketch_name1)  # '需求參數(參數所在本體=body_name(2),需更換的參數=formula_name(1),被換成曲線NAME=line_name(1),需更先進行更新草圖(假如沒有就隨便抓一個)=sketch_name(1))
    parameter1 = parameters2.DirectParameters.Item("Reinforcement_insert_height")
    parameter1.Value = gvar.strip_parameter_list[20]
    # '入子高度
    # '======================================================================================================
    parameter1 = parameters2.DirectParameters.Item("Reinforcement_open_height")
    if gvar.Mold_status == "開模":
        parameter1.Value = float(gvar.strip_parameter_list[1]) + 28  # (upper_die_open_height)
    elif gvar.Mold_status == "閉模":
        parameter1.Value = float(gvar.strip_parameter_list[1])
    # '======================================================================================================
    part1.Update()
    point_name1 = "Bolt_Point_1"
    (point_name2) = stop_plate_point_build(g, point_name1, 3, body_name1, body_name2)
    (point_name2) = stop_plate_point_build(g, point_name1, 11, body_name1, body_name2)
    F_Hole_up(n, i, 1, 16, element_name1, point_name2, body_name1)
    point_name1 = "Bolt_Point_2"
    (point_name2) = stop_plate_point_build(g, point_name1, 3, body_name1, body_name2)
    (point_name2) = stop_plate_point_build(g, point_name1, 11, body_name1, body_name2)
    F_Hole_up(n, i, 2, 16, element_name1, point_name2, body_name1)


def F_bending_stop_plate(g, op_number, now_data_number, SketchPosition):
    OP = op_number
    i = now_data_number
    EX_file_name = "op" + str(OP) + "_bending_punch_" + str(i)
    plate_name = "Data1"
    sketch_name_line = "judgment_sketch"
    deep = "max"
    body_name2 = "bending_punch"
    body_name3 = "Body.2"
    parameter_name5 = "bending_punch_parameter"
    gap = 0
    defs.FUN_pad_gap(EX_file_name, plate_name, sketch_name_line, deep, gap, body_name2, body_name3, parameter_name5,
                     SketchPosition,
                     now_data_number)  # 貼上元素   body_name(2)=copy"body1"  body_name(3)=paste "body2" deep="max"=>挖到底

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
        for j in range(1, 3):
            part1 = partDocument1.Part
            hybridShapeFactory1 = part1.HybridShapeFactory
            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, 0)
            hybridBodies1 = part1.HybridBodies
            hybridShapes1 = body2.HybridShapes
            hybridShapePointCoord2 = hybridShapes1.Item("Bolt_point_" + str(j))
            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord2)
            hybridShapePointCoord1.PtRef = reference1
            bodies1 = part1.Bodies
            body1 = bodies1.Item("Body.2")
            hybridShapePointExplicit1 = hybridShapeFactory1.AddNewPointDatum(reference1)
            body1.InsertHybridShape(hybridShapePointExplicit1)
            part1.InWorkObject = hybridShapePointExplicit1
            reference1 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
            hybridShapes1 = body1.HybridShapes
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 40)
            hole1.Type = 0
            hole1.AnchorMode = 1
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.BottomType = 1
            length2 = limit1.dimension
            length2.Value = 24
            hole1.ThreadingMode = 0
            hole1.CreateStandardThreadDesignTable(1)
            strParam1 = hole1.HoleThreadDescription
            strParam1.Value = "M8"
            length3 = hole1.ThreadDepth
            length3.Value = 16
            hole1.Reverse()
            hole1.Name = "Stop-plate-A-punch-bolt-" + str(stop_plate_A_punch_number) + "-" + str(j)
            try:
                part1.UpdateObject(hole1)
            except:
                selection1 = partDocument1.Selection
                selection1.Clear()
                selection1.add(hole1)
                selection1.Delete()
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
# StopPlate(1)

def emboss_forming_punch_left(g, n):  # 打凸包沖頭_左
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    op_number = n * 10
    for i in range(1, round(gvar.StripDataList[21][g][n]) + 1):
        partDocument1 = documents1.Open(
            gvar.save_path + "op" + str(op_number) + "_emboss_forming_punch_left_0" + str(i) + ".CATPart")
        documents1 = catapp.Documents
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        # ===============================================
        length1 = parameters1.Item("stop_plate_gap")
        length1.Value = 0
        # ===============================================
        part1.Update()
        # ------------------------------------------------------------↓offset
        hybridShapeFactory1 = part1.HybridShapeFactory
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        hybridShapes1 = body1.HybridShapes
        hybridShape3DCurveOffset1 = hybridShapes1.Item("demise_hole_left_offset")
        reference1 = part1.CreateReferenceFromObject(hybridShape3DCurveOffset1)
        hybridShapePlaneOffset1 = hybridShapes1.Item("punch_up_plane")
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
        part1.InWorkObject.Name = "demise_hole_left_offset_project"
        part1.Update()
        hybridShapeFactory1.DeleteObjectForDatum(reference3)
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=demise_hole_left_offset_project*,all")
        selection1.VisProperties.SetShow(1)  # 1隱藏/0顯示
        selection1.Clear()
        # ------------------------------------------------------------↑offset
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        selection1 = partDocument1.Selection
        selection1.Clear()
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("demise_hole_left_offset_project")
        selection1.Add(hybridShapeCurveExplicit1)
        selection1.Copy()
        # ===================================================================
        windows1 = catapp.Windows
        specsAndGeomWindow1 = windows1.Item("Data1.CATPart")
        specsAndGeomWindow1.Activate()
        # ===================================================================
        partDocument2 = catapp.ActiveDocument
        selection2 = partDocument2.Selection
        selection2.Clear()
        part2 = partDocument2.Part
        bodies1 = part2.Bodies
        body1 = bodies1.Item("Body.2")
        part2.InWorkObject = body1
        selection2.Add(part2)
        selection2.Paste()
        part2.InWorkObject.Name = "demise_hole_left_offset_project_" + str(i)
        partDocument1.Close()
        # ---------------------------------------------
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Add()
        body1.Name = "remove_body_" + str(i)
        part1.Update()
        hybridShapeFactory1 = part1.HybridShapeFactory
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("demise_hole_left_offset_project_" + str(i))
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        hybridShapes1 = body1.HybridShapes
        hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
        reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
        hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference1, reference2)
        hybridShapeProject1.SolutionType = 0
        hybridShapeProject1.Normal = True
        hybridShapeProject1.SmoothingType = 0
        body2 = bodies1.Item("remove_body_" + str(i))
        body2.InsertHybridShape(hybridShapeProject1)
        part1.InWorkObject = hybridShapeProject1
        part1.Update()
        shapeFactory1 = part1.ShapeFactory
        reference3 = part1.CreateReferenceFromName("")
        pad1 = shapeFactory1.AddNewPadFromRef(reference3, 20)
        limit1 = pad1.FirstLimit
        limit1.LimitMode = 3
        hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_up_plane")
        reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
        limit1.LimitingElement = reference4
        reference5 = part1.CreateReferenceFromObject(hybridShapeProject1)
        pad1.SetProfileElement(reference5)
        reference6 = part1.CreateReferenceFromObject(hybridShapeProject1)
        pad1.SetProfileElement(reference6)
        part1.Update()
        part1.InWorkObject = body1
        remove1 = shapeFactory1.AddNewRemove(body2)
        part1.Update()
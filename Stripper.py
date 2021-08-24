import win32com.client as win32
import defs
import global_var as gvar
import time

def Stripper(now_plate_line_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    partDocument2 = documents1.Open(gvar.open_path + "Stripper.CATPart")
    # ======================================
    defs.window_change(partDocument1, partDocument2)  # 在catapp上切換各視窗
    # ======================================
    # ------------------------------------------------------------↓判斷是否有沖切模組
    selection3 = partDocument1.Selection
    selection3.Clear()
    selection3.Search("Name=plate_line_" + str(now_plate_line_number) + "_op*_cut_punch_*_cutting_*,all")
    if selection3.Count > 0:
        documents1 = catapp.Documents
        partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_shear.CATPart")
        # ======================================
        defs.window_change(partDocument1, partDocument2)  # 在catapp上切換各視窗
        # ======================================
        part2 = partDocument1.Part
        bodies1 = part2.Bodies
        # ------------------------------------------------------------↓改變body名稱
        body1 = bodies1.Item("Body.3")
        body1.Name = "cut_cavity_insert_shear"
        # ------------------------------------------------------------↑
        selection3.Clear()
        selection3.Search("Name=*cut_cavity_insert_shear,all")
        selection3.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
        selection3.Clear()
    # -----------------------------------------------------------------------------↓判斷是否有補強沖頭模組
    selection3.Clear()
    selection3.Search("Name=plate_line_" + str(now_plate_line_number) + "_op*_Reinforcement_cut_line_*,all")
    if selection3.Count > 0:
        documents1 = catapp.Documents
        partDocument2 = documents1.Open(gvar.open_path + "QR_punch_Reinforcement.CATPart")
        part1 = partDocument2.Part
        parameters1 = part1.Parameters
        # hybridShapeCurveExplicit1 As HybridShapeCurveExplicit
        hybridShapeCurveExplicit1 = parameters1.Item("cut_line_formula_1")
        hybridShapeCurveExplicit1.Name = "cut_line_formula_3"
        # ======================================
        defs.window_change(partDocument1, partDocument2)  # 在catapp上切換各視窗
        # ======================================
        part2 = partDocument2.Part
        bodies1 = part2.Bodies
        body_number = bodies1.Count
        # ------------------------------------------------------------↓   改變body名稱
        body1 = bodies1.Item("Body." + body_number)
        body1.Name = "Reinforcement_cut_punch"
        # ------------------------------------------------------------↑
        selection3.Clear()
        selection3.Add(body1)
        selection3.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
        selection3.Clear()
        # --------------------------------另一個檔案複製-------------
        partDocument2 = documents1.Open(gvar.open_path + "cut_Reinforcement_insert.CATPart")
        part1 = partDocument2.Part
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("cut_line_assume_1")
        part1.InWorkObject = hybridShapeCurveExplicit1
        hybridShapeCurveExplicit1.Name = "cut_line_formula_4"
        bodies1 = part1.Bodies
        body1 = bodies1.Item("PartBody")
        body1.Name = "Reinforcement_insert"
        # ======================================
        defs.window_change(partDocument1, partDocument2)  # 在catapp上切換各視窗
        # ======================================
        selection3.Clear()
        selection3.Add(body1)
        selection3.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
        selection3.Clear()
    # -----------------------------------------------------------------------------↑判斷是否有補強沖頭模組
    # ============================================================================Stripper_change.CATMain
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    length = [None] * 99
    formula = [None] * 99
    total_op_number = int(gvar.strip_parameter_list[2])
    stripper_insert_hole_number = int()
    point_name = [""] * 7
    body_name = [""] * 10
    sketch_name = [""] * 10
    insert_line_name = [""] * 5
    # ======================================================================================================
    length[0] = part1.Parameters.Item("pitch")
    length[0].Value = int(gvar.strip_parameter_list[4])
    # ======================================================================================================
    # ======================================================================================================
    length[1] = part1.Parameters.Item("plate_height")
    g = now_plate_line_number
    if gvar.strip_parameter_list[20] != "":
        if gvar.die_type == "module":
            length[1].Value = 0  # (back_stripper_plate_height)
        else:
            length[1].Value == float(gvar.strip_parameter_list[20])
    else:
        die_rule_file_name = "模板厚度選擇"
        excel_Sheet_name = "200T以下"
        Column_string_serch = "脫料板"
        Row_string_serch = "精密級"
        (serch_result) = defs.ExcelSearch(die_rule_file_name, excel_Sheet_name, Row_string_serch, Column_string_serch)
        length[1].Value = -serch_result
        gvar.strip_parameter_list[20] = serch_result
    # ======================================================================================================
    length[4] = part1.Parameters.Item("plate_down_plane")
    if gvar.die_type == "module":
        plate_position = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20])
    else:
        if gvar.Mold_status == "開模":
            plate_position = float(gvar.strip_parameter_list[1]) + 0  # (-stripper_stroke)
        else:
            plate_position = float(gvar.strip_parameter_list[1])
    if gvar.Mold_status == "開模":
        length[4].Value = float(gvar.strip_parameter_list[1]) + 28  # (upper_die_open_height)
    elif gvar.Mold_status == "閉模":
        length[4].Value = plate_position
    # ======================================================================================================
    file_name = "Data1"
    body_name1 = "Body.2"
    hybridBody_name = "die"
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(file_name, body_name1,
                                                                                             hybridBody_name)  # 環境設定(檔案名,body名,依據群組名)(全域變數改)
    if gvar.die_type == "module":
        # M_plate_length = 35
        # M_plate_wide = 116
        # hybridShape1 = ElementBody.HybridShapes.Item("down_die_plate_up_plane") #宣告平面(下)
        # defs.material_tpye_palte_sketch(hybridShape1, M_plate_length, M_plate_wide, 100 + 50 * (now_plate_line_number - 1), 112.5 + 12.5) #(平面的宣告,模板長,模板寬,中心座標_X,中心座標_Y)
        # part1.Parameters.Item("1_formula_1").OptionalRelation.Modify ("die\\plate_size") #草圖置換
        # part1.Update()
        # M_stripper_design(M_plate_length, M_plate_wide)
        pass  ##未使用
    else:
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\number_" + str(g) + "_plate_line")  # 草圖置換
    part1.Update()
    specsAndGeomWindow1 = catapp.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    part1.Update()
    Stripper_machining_explanation_shape = 0  # ------------------------------------------------------------加工說明
    ss_count = 0  # 異形沖挖孔計數
    for now_op_number in range(1, 1 + total_op_number):
        n = now_op_number
        op_number = 10 * n
        if gvar.StripDataList[38][g][n] > 0:
            for i in range(1, 1 + gvar.StripDataList[38][g][n]):
                Stripper_machining_explanation_shape = Stripper_machining_explanation_shape + 1  # ------------------------------------加工說明
                # ----------cut_line_number(h) > 0 ↓------------------------
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
                hybridShapeFactory1.DeleteObjectforDatum(reference3)
                part1.Update()
                hybridShapeFactory1.DeleteObjectforDatum(reference3)
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
                                                                                    0, 0.5, 1)  # 沖孔偏移距離
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
                hybridShapeFactory1.DeleteObjectforDatum(reference5)
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
                stripper_insert_hole_number = stripper_insert_hole_number + 1
                pocket1.Name = "Stripper-plate-Insert-hole-" + str(stripper_insert_hole_number)
                part1.Update()
            # ------------------------------------------------------
        # ----------cut_line_number(h) > 0↑------------------------
        # ----------bend_up_shaping_cavity_1_number(g,n)>0↓-----------------------------整平模組孔1邊
        if gvar.StripDataList[42][g][n] > 0:
            # for i in range( 1 , 1+ gvar.StripDataList[42][g][n]):
            #     pp_count = i
            #     Call bend_up_shaping_cavity_hole_1
            pass  # 未使用
        # ----------bend_up_shaping_cavity_1_number(g,n)>0↑-----------------------------整平模組孔1邊
        # ----------bend_up_shaping_cavity_2_number(g,n)>0↓-----------------------------整平模組孔2邊
        if gvar.StripDataList[43][g][n] > 0:
            # for i in range( 1 , 1+ gvar.StripDataList[43][g][n]):
            #     pp_count = i
            #     Call bend_up_shaping_cavity_hole_2
            pass  # 未使用
        # ----------bend_up_shaping_cavity_2_number(g,n)>0↑-----------------------------整平模組孔2邊
        # ----------forming_punch_surface_number(h) > 0 ↓------------------------
        if gvar.StripDataList[40][g][n] > 0:
            for i in range(1, 1 + gvar.StripDataList[43][g][n]):
                Stripper_machining_explanation_shape = Stripper_machining_explanation_shape + 1  # ------------------------------------加工說明
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
                hybridShapeFactory1.DeleteObjectforDatum(reference2)
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
                    op_number) + "_forming_punch_Project_" + str(i)  # 更改外形線名稱
                hybridShapeFactory1.DeleteObjectforDatum(reference3)
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
                    op_number) + "_forming_punch_offset_" + str(i)
                part1.Update()
                hybridShapeFactory1.DeleteObjectforDatum(reference15)
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
                    "plate_line_" + str(g) + "_op" + str(op_number) + "_forming_punch_offset_" + str(i))
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
        # ---------- forming_punch_surface_number(h) > 0 ↑------------------------
        # ------------------------------------------------------------↓補強沖頭
        if gvar.StripDataList[4][g][n] > 0:
            # for for_counter in range (1 , 1+ gvar.StripDataList[4][g][n]):
            #     parameter_digital[1] = 0
            #     punch_Reinforcement_Ecxavation(g, n, for_counter)
            #     parameter_digital[1] = 0#(stripper_plate_space)
            #     insert_Reinforcement_Ecxavation(g, n, for_counter)
            pass  # 未使用
        # ------------------------------------------------------------↑補強沖頭
        # ----------------------------------------------------------------------------------------------------------------↓切斷工程
        if gvar.StripDataList[27][g][n] > 0:
            # for cut_punch_u_number in range( 1 , 1+ gvar.StripDataList[27][g][n])
            #     #------------------------------------------------------------↓參數名稱
            #     insert_Gap = 0#(stripper_plate_space)
            #     point_name[1] = "cut_base_point_X_min"
            #     point_name[2] = "op" + str(op_number) + "_d_cut_X_min_point"
            #     point_name[3] = "cut_base_point_X_max"
            #     point_name[4] = "op" + str(op_number) + "_d_cut_X_max_point"
            #     point_name[5] = "cut_base_point_Z"
            #     point_name[6] = "op" + str(op_number) + "_d_cut_Z_point"
            #     body_name[1] = "cut_cavity_insert_shear"
            #     sketch_name[1] = "down_cut_insert_Sketch"
            #     type_name = "B_down"
            #     insert_line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_d_cutting_" + str(cut_punch_u_number) #草圖置換名稱
            #     insert_line_name[2] = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_d_insert_surface_" + str(cut_punch_u_number)
            #     #------------------------------------------------------------↑
            #     punch_d_cutting                            #沖頭讓位
            #     cut_cavity_change.punch_cutting_change     #換草圖
            #     cut_cavity_change.punch_cutting            #挖槽
            pass  # 未使用
        if gvar.StripDataList[27][g][n] > 0:
            # for cut_punch_u_number in range( 1 , 1+ gvar.StripDataList[27][g][n])
            #     #------------------------------------------------------------↓參數名稱
            #     insert_Gap = 0#(stripper_plate_space)
            #     point_name[1] = "cut_base_point_X_min"
            #     point_name[2] = "op" + str(op_number) + "_d_cut_X_min_point"
            #     point_name[3] = "cut_base_point_X_max"
            #     point_name[4] = "op" + str(op_number) + "_d_cut_X_max_point"
            #     point_name[5] = "cut_base_point_Z"
            #     point_name[6] = "op" + str(op_number) + "_d_cut_Z_point"
            #     body_name[1] = "cut_cavity_insert_shear"
            #     sketch_name[1] = "down_cut_insert_Sketch"
            #     type_name = "B_down"
            #     insert_line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_u_cutting_" + str(cut_punch_u_number) #草圖置換名稱
            #     insert_line_name[2] = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_u_insert_surface_" + str(cut_punch_u_number)
            #     #------------------------------------------------------------↑
            #     punch_d_cutting                            #沖頭讓位
            #     cut_cavity_change.punch_cutting_change     #換草圖
            #     cut_cavity_change.punch_cutting            #挖槽
            pass  # 未使用
        # ----------------------------------------------------------------------------------------------------------------↑切斷工程
        if gvar.StripDataList[37][g][n] > 0:
            A_punch(now_plate_line_number, op_number)
        if gvar.StripDataList[3][g][n] > 0:  # 半衝切
            # Call QR_half_cut_punch
            pass  # 未使用
    for now_op_number in range(1, 1 + total_op_number):
        n = now_op_number
        op_number = 10 * n
        # ----------------------------------------------------------------------------------↓快拆沖頭
        if gvar.StripDataList[29][g][n] > 0:  # 沖切沖頭_右
            # data_type = "line"
            # data_number = gvar.StripDataList[29][g][n]
            # product_name = "op" + str(op_number) + "_right_quickly_remove_cut_punch_stripper_insert_"
            # part_name = "right_quickly_remove_cut_punch_stripper_insert_line_"
            # parameter_name[1] = "insert_line_stripper_demise_Sketch_Xmax_length"
            # parameter_name[2] = "insert_line_stripper_demise_Sketch_offset_Xmax_length"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[30][g][n] > 0:  # 沖切沖頭_左
            # data_type = "line"
            # data_number = gvar.StripDataList[30][g][n]
            # product_name = "op" + str(op_number) + "_left_quickly_remove_cut_punch_stripper_insert_"
            # part_name = "left_quickly_remove_cut_punch_stripper_insert_line_"
            # parameter_name[1] = "insert_line_stripper_demise_Sketch_Xmax_length"
            # parameter_name[2] = "insert_line_stripper_demise_Sketch_offset_Xmax_length"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[31][g][n] > 0:  # 沖切沖頭_上
            # data_type = "line"
            # data_number = gvar.StripDataList[31][g][n]
            # product_name = "op" + str(op_number) + "_up_quickly_remove_cut_punch_stripper_insert_"
            # part_name = "up_quickly_remove_cut_punch_stripper_insert_line_"
            # parameter_name[1] = "insert_line_stripper_demise_Sketch_Ymax_length"
            # parameter_name[2] = "insert_line_stripper_demise_Sketch_offset_Ymax_length"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[32][g][n] > 0:  # 沖切沖頭_下
            # # data_type = "line"
            # # data_number = gvar.StripDataList[32][g][n]
            # product_name = "op" + str(op_number) + "_down_quickly_remove_cut_punch_stripper_insert_"
            # part_name = "down_quickly_remove_cut_punch_stripper_insert_line_"
            # parameter_name[1] = "insert_line_stripper_demise_Sketch_Ymax_length"
            # parameter_name[2] = "insert_line_stripper_demise_Sketch_offset_Ymax_length"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[33][g][n] > 0:  # 成形沖頭_右
            # data_type = "surface"
            # data_number = gvar.StripDataList[33][g][n]
            # product_name = "op" + str(op_number) + "_right_quickly_remove_bending_punch_stripper_insert_"
            # part_name = "right_quickly_remove_bending_punch_stripper_insert_line_"
            # parameter_name[1] = "insert_line_stripper_demise_Sketch_Xmax_length"
            # parameter_name[2] = "insert_line_stripper_demise_Sketch_offset_Xmax_length"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[34][g][n] > 0:  # 成形沖頭_左
            # data_type = "surface"
            # data_number =gvar.StripDataList[34][g][n]
            # product_name = "op" + str(op_number) + "_left_quickly_remove_bending_punch_stripper_insert_"
            # part_name = "left_quickly_remove_bending_punch_stripper_insert_line_"
            # parameter_name[1] = "insert_line_stripper_demise_Sketch_Xmax_length"
            # parameter_name[2] = "insert_line_stripper_demise_Sketch_offset_Xmax_length"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[35][g][n] > 0:  # 成形沖頭_上
            # data_type = "surface"
            # data_number = gvar.StripDataList[35][g][n]
            # product_name = "op" + str(op_number) + "_up_quickly_remove_bending_punch_stripper_insert_"
            # part_name = "up_quickly_remove_bending_punch_stripper_insert_line_"
            # parameter_name[1] = "insert_line_stripper_demise_Sketch_Ymax_length"
            # parameter_name[2] = "insert_line_stripper_demise_Sketch_offset_Ymax_length"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[36][g][n] > 0:  # 成形沖頭_下
            # data_type = "surface"
            # data_number = gvar.StripDataList[36][g][n]
            # product_name = "op" + str(op_number) + "_down_quickly_remove_bending_punch_stripper_insert_"
            # part_name = "down_quickly_remove_bending_punch_stripper_insert_line_"
            # parameter_name[1] = "insert_line_stripper_demise_Sketch_Ymax_length"
            # parameter_name[2] = "insert_line_stripper_demise_Sketch_offset_Ymax_length"
            # quickly_remove_punch
            pass  # 未使用
        # ----------------------------------------------------------------------------------↑快拆沖頭
        # --------------------------------------------------------------------------------------------整形沖頭
        if gvar.StripDataList[74][g][n] > 0:
            # for now_data_number in range(1, 1 + gvar.StripDataList[74][g][n]):
            #     F_bending_stripper
            pass  # 未使用
    # --------------------------↓讓位↓---------------------------
    for o in range(1, 1 + total_op_number):
        for j in range(1, 1 + 5):
            plate_op = g * 100 + o
            if gvar.StripDataList[2][g][n] > 0:
                # length[7] = part1.Parameters.Item("demise_height")
                # length[7].Value = plate_line_demise_surface_up_number(plate_op, j)
                # #-------↓Function demise用↓---------
                # demise_plane_h = length[7].Value
                # data_line = j
                # demise_op = o * 10
                # #-------↑Function demise用↑---------
                # Call demise
                pass  # 未使用
    # --------------------------↑讓位↑---------------------------
    selection1 = partDocument1.Selection
    selection1.Clear()
    selection1.Search("Name=Reinforcement_insert,all")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    part1.Update()
    product1.PartNumber = "Stripper_" + str(g)  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Stripper_" + str(g))
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
    strParam4.ValuateFromString(gvar.strip_parameter_list[21])
    product1 = product1.ReferenceProduct
    parameters7 = product1.UserRefProperties
    strParam5 = parameters7.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(gvar.strip_parameter_list[22])
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
    strParam8.ValuateFromString("S1: " + str(Stripper_machining_explanation_shape) + "-(,割), 單+0.005")
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
    strParam15 = parameters14.CreateString("F", "")  # 外導柱
    strParam15.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters16 = product1.UserRefProperties
    strParam16 = parameters14.CreateString("CS", "")  # 等高套筒
    strParam16.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters17 = product1.UserRefProperties
    strParam17 = parameters17.CreateString("AP", "")  # A沖沖孔
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
        selection1.Clear()
    if now_plate_line_number == 2:
        selection1.Clear()
        selection1 = partDocument1.Selection
        selection1.Search("Name=plate_line_1*,all")
        selection1.Delete()
        selection1.Clear()
    selection1 = partDocument1.Selection
    selection1.Search("Name=cut_line_assume,all")
    if selection1.Count > 0:
        selection1.Delete()
    selection1.Clear()
    # --------------↑刪除不需要的Data↑--------------
    part1.Update()
    time.sleep(2)
    partDocument1.SaveAs(gvar.save_path + "Stripper_" + str(g) + ".CATPart")  # 存檔的檔案名稱
    time.sleep(2)
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    partDocument1.Close()


def A_punch(now_plate_line_number, op_number):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    partDocument2 = documents1.Open(gvar.open_path + "QR_Stripper_line.CATPart")
    # ======================================
    defs.window_change(partDocument1, partDocument2)  # 在CATIA上切換各視窗
    # ======================================
    length = [None] * 5
    partDocument1 = documents1.Open(gvar.open_path + "SJAS.CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length[2] = part1.Parameters.Item("D")
    length[2].Value = float(gvar.strip_parameter_list[23])
    length[3] = part1.Parameters.Item("H")
    partDocument1.Close()
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length[0] = part1.Parameters.Item("D")
    length[0].Value = float(gvar.strip_parameter_list[23])
    length[1] = part1.Parameters.Item("H")
    length[1].Value = length[3].Value
    g = now_plate_line_number
    n = int(op_number / 10)
    part1.Update()
    cut_cavity_machining_explanation_shape = int()
    stripper_A_punch_number = int()
    for i in range(1, 1 + gvar.StripDataList[37][g][n]):
        cut_cavity_machining_explanation_shape = cut_cavity_machining_explanation_shape + 1  # --------------------------------------------------加工說明
        part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
            "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i))  # 草圖置換
        part1.Update()
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        hybridShapeFactory1 = part1.HybridShapeFactory
        for B_n in range(50, 1 + 1, -1):
            selection1 = partDocument1.Selection
            selection1.Clear()
            selection1.Search("Name=Body." + str(B_n) + "*,all")
            Body_n = selection1.Count
            selection1.Clear()
            if Body_n > 0:
                body_number = B_n
                break
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
        stripper_A_punch_number = stripper_A_punch_number + 1
        pocket1.Name = "Stripper-plate-A-punch-" + str(stripper_A_punch_number)
        part1.Update()
    selection2 = partDocument1.Selection
    selection2.Clear()
    selection2.Search(
        "Name=Body." + str(body_number))
    if selection2.Count > 0:
        selection2.Delete()
    selection2.Clear()
    selection2.Search("Name=D")
    if selection2.Count > 0:
        selection2.Delete()
    selection2.Clear()
    selection2.Search("Name=H")
    if selection2.Count > 0:
        selection2.Delete()
    selection2.Clear()

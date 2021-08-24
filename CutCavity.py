import win32com.client as win32
import defs
import global_var as gvar
import PunchDef
import time
import math
def CutCavity(now_plate_line_number):
    total_op_number = int(gvar.strip_parameter_list[2])
    element_name = [""] * 3
    catapp = win32.Dispatch('CATIA.Application')
    windows1 = catapp.Windows
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    partDocument2 = documents1.Open(gvar.open_path + "cut_cavity.CATPart")
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
    # ------------------------------------------------------------↓   改變body名稱
    body1 = bodies1.Item("Body.3")
    body1.Name = "cut_cavity_insert"
    # ------------------------------------------------------------↑
    hybridShapes1 = body1.HybridShapes
    # ------------------------------------------------------------↓   定義刪除的東西
    sketches1 = body1.Sketches
    sketch1 = sketches1.Item("Bolt_point_Sketch")
    shapes1 = body1.Shapes
    pad1 = shapes1.Item("Pad.4")  # 原本是2
    pocket1 = shapes1.Item("cut_Pocket")
    hybridShapeExtremum1 = hybridShapes1.Item("insert_X_max")
    hybridShapeExtremum2 = hybridShapes1.Item("insert_X_min")
    hybridShapeExtremum3 = hybridShapes1.Item("insert_Y_max")
    hybridShapeExtremum4 = hybridShapes1.Item("insert_Y_min")
    hybridShapeExtremum5 = hybridShapes1.Item("ReferenceExtremum.1")
    hybridShapeExtremum6 = hybridShapes1.Item("ReferenceExtremum.2")
    hybridShapeExtremum7 = hybridShapes1.Item("ReferenceExtremum.3")
    hybridShapePointBetween1 = hybridShapes1.Item("Point_1")
    hybridShapePlaneNormal1 = hybridShapes1.Item("ReferencePlane")
    hybridShapeLinePtPt1 = hybridShapes1.Item("ReferenceLine.1")
    # ------------------------------------------------------------↑
    # ------------------------------------------------------------↓   刪除
    selection3 = partDocument1.Selection
    time.sleep(2)
    selection3.Clear()
    selection3.Add(hybridShapeExtremum1)
    selection3.Add(hybridShapeExtremum2)
    selection3.Add(hybridShapeExtremum3)
    selection3.Add(hybridShapeExtremum4)
    selection3.Add(hybridShapeExtremum5)
    selection3.Add(hybridShapeExtremum6)
    selection3.Add(hybridShapeExtremum7)
    selection3.Add(hybridShapePointBetween1)
    selection3.Add(hybridShapePlaneNormal1)
    selection3.Add(hybridShapeLinePtPt1)
    selection3.Add(sketch1)
    selection3.Add(pocket1)
    for i in range(1, 1 + 5):
        try:
            hybridShapePointOnPlane1 = hybridShapes1.Item("Bolt_point_" + str(i))
            hole1 = shapes1.Item("Bolt_Hole_" + str(i))
            selection3.Add(hole1)
            selection3.Add(hybridShapePointOnPlane1)
            selection3.Add(pad1)
        except:
            pass
    selection3.Delete()
    selection3.Clear()
    # ------------------------------------------------------------↑
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
        # ------------------------------------------------------------↓   隱藏
        selection3.Clear()
        selection3.Search("Name=*cut_cavity_insert_shear,all")
        selection3.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
        selection3.Clear()
        # ------------------------------------------------------------↑
    # ======================================(cut_cavity_Change)======================================
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    length = [None] * 99
    formula = [None] * 99
    pitch = int(gvar.strip_parameter_list[4])
    lower_die_insert_hole_count = int()
    # ======================================================================================================
    length[0] = part1.Parameters.Item("pitch")
    length[0].Value = pitch
    # ======================================================================================================
    # ======================================================================================================
    length[1] = part1.Parameters.Item("plate_height")
    g = now_plate_line_number
    if gvar.strip_parameter_list[26] != "":
        if gvar.die_type == "module":
            length[1].Value = int(gvar.strip_parameter_list[20])
        else:
            length[1].Value = -int(gvar.strip_parameter_list[26])
    else:
        die_rule_file_name = "模板厚度選擇"
        excel_Sheet_name = "200T以下"
        Column_string_serch = "下模板"
        Row_string_serch = "精密級"
        (serch_result) = defs.ExcelSearch(die_rule_file_name, excel_Sheet_name, Row_string_serch, Column_string_serch)
        length[1].Value = -serch_result
        gvar.strip_parameter_list[26] = str(serch_result)
    # ======================================================================================================
    length[4] = part1.Parameters.Item("plate_up_plane")
    if gvar.die_type == "module":
        length[4].Value = float(gvar.strip_parameter_list[1])
    else:
        length[4].Value = -0  # die_open_height
    # ======================================================================================================
    # ======================================================================================================
    length[33] = part1.Parameters.Item(
        "gap")  # ************入子間隙************入子間隙************入子間隙************入子間隙************入子間隙
    length[33].Value = float(gvar.strip_parameter_list[1]) * 0.02  # (lower_die_space / 100 )
    file_name = "Data1"
    body_name1 = "Body.2"
    hybridBody_name = "die"
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(file_name, body_name1,
                                                                                             hybridBody_name)  # 環境設定(檔案名,body名,依據群組名)(全域變數改)
    if gvar.die_type == "module":
        M_plate_length = 35
        M_plate_wide = 116
        hybridShape1 = ElementBody.HybridShapes.Item("down_die_plate_up_plane")  # 宣告平面(下)
        defs.material_tpye_palte_sketch(hybridShape1, M_plate_length, M_plate_wide,
                                        100 + 50 * (now_plate_line_number - 1), 112.5 + 12.5, ElementDocument,
                                        ElementBody, ElementHybridBody)  # (平面的宣告,模板長,模板寬,中心座標_X,中心座標_Y)
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\plate_size")  # 草圖置換
        # Call M_cavity_design(M_plate_length, M_plate_wide)
        pass  # 未使用
    else:
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\number_" + str(g) + "_plate_line")  # 草圖置換
    part1.Update()
    specsAndGeomWindow1 = catapp.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    viewer3D1.Reframe()
    part1.Update()
    # =================↓挖下料孔↓================================================
    cut_cavity_machining_explanation_shape = 0  # ----------------------------------------------------------加工說明
    for n in range(1, 1 + total_op_number):
        op_number = 10 * n
        cut_type = [""] * 5
        xi = int()
        if gvar.StripDataList[38][g][n] > 0:
            if gvar.StripDataList[38][g][n] > 1:
                xi = gvar.StripDataList[38][g][n]
            for i in range(1, 1 + 1):
                cut_cavity_machining_explanation_shape = cut_cavity_machining_explanation_shape + 1  # ----------------------------------加工說明
                # ---------------------------------------------------------------------------------------------------------------------------------------------------↓對稱入子
                # ------------------------------------------------------------↓   搜尋是否為對稱名稱
                selection3 = partDocument1.Selection
                selection3.Clear()
                selection3.Search(
                    "Name=plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_symmetric_" + str(i))
                # ------------------------------------------------------------↑
                # ------------------------------------------------------------↓   判斷是否對稱名稱
                if selection3.Count > 0:
                    part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                        "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_symmetric_" + str(i))  # 草圖置換
                else:
                    part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                        "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i))  # 草圖置換
                # ------------------------------------------------------------↑
                # ------------------------------------------------------------↓  修正為對稱孔位之參數 True=對稱  False=不對稱
                boolParam1 = part1.Parameters.Item("symmetry_switch")
                if selection3.Count > 0:
                    boolParam1.Value = True
                else:
                    boolParam1.Value = False
                # ------------------------------------------------------------↑
                # ------------------------------------------------------------↓   參數宣告
                length[34] = part1.Parameters.Item("x_to_x")
                length[35] = part1.Parameters.Item("y_to_y")
                length[36] = part1.Parameters.Item("int_x")
                length[37] = part1.Parameters.Item("int_y")
                # ------------------------------------------------------------↑
                if xi > 1:
                    PunchDef.cut_cavity_clash_change(xi, op_number)
                # ------------------------------------------------------------↓   先放大入子以免更換時  孔>入子塊
                length[36].Value = 500
                length[37].Value = 500
                part1.Update()
                # ------------------------------------------------------------↑
                q = length[34].Value
                R = length[35].Value
                length[36].Value = math.ceil(q)
                length[37].Value = math.ceil(R)
                # ------------------------------------------------------------↑
                part1.Update()
                # ---------------------------------------------------------------------------------------------------------------------------------------------------↑對稱入子
                bodies1 = part1.Bodies
                body1 = bodies1.Item("Body.2")
                part1.InWorkObject = body1
                hybridShapeFactory1 = part1.HybridShapeFactory
                body2 = bodies1.Item("cut_cavity_insert")
                sketches1 = body2.Sketches
                sketch1 = sketches1.Item("cut_insert_line_Sketch")
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
                part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(
                    i) + "_project_line"
                part1.Update()
                hybridShapeFactory1.DeleteObjectforDatum(reference3)
                selection1 = partDocument1.Selection
                selection1.Clear()
                selection1.Search("Name=*_project_line,all")
                selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
                selection1.Clear()
                # ----------------------------------↑投影外型線↑--------------------------------------
                # ----------------------------------↓OFFSET↓------------------------------------------
                reference4 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
                # ---------------------------------------------------------------------------------------------------------------------------------------------------↓對稱入子
                hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)  # 改Z方向
                hybridShape3DCurveOffset1 = hybridShapeFactory1.AddNew3DCurveOffset(reference4, hybridShapeDirection1,
                                                                                    0.005, 0, 0)  # 偏移距離
                # ---------------------------------------------------------------------------------------------------------------------------------------------------↑對稱入子
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
                lower_die_insert_hole_count = lower_die_insert_hole_count + 1
                pocket1.Name = ("Lower-die-Insert-hole-" + str(lower_die_insert_hole_count))
                part1.Update()
                if xi > 1:
                    selection1 = partDocument1.Selection
                    selection1.Clear()
                    selection1.Search("Name=projection.*,all")
                    selection1.Delete()
                    xi = 0
        # ====================================================================
        # ------------------------------------------------------------↓點名稱參數
        point_name = [''] * 6
        point_name[1] = "cut_base_point_X_min"
        point_name[3] = "cut_base_point_X_max"
        point_name[5] = "cut_base_point_Z"
        # ------------------------------------------------------------↑
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↓ 補強入子
        if gvar.StripDataList[4][g][n] > 0:
            # for for_counter in range (1 , 1+ gvar.StripDataList[4][g][n]):
            #     Reinforcement_cut_Excavation(g, n, for_counter)
            pass  # 未使用
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↑ 補強入子
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
        if gvar.StripDataList[37][g][n] > 0:
            A_punch(now_plate_line_number, op_number)
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↓I T M 對稱入子 2015-12-25
        if gvar.StripDataList[53][g][n] > 0:
            # for for_counter in range(1, 1 + gvar.StripDataList[53][g][n]):
            #     body_name[1] = "Body.2"
            #     body_name[2] = "cut_cavity_insert"
            #     formula_name[1] = "cut_line_formula_1"
            #     element_name[1] = "down_die_plate_up_plane"
            #     line_name[2] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_T_symmetric_" +str( for_counter)
            #     line_name[3] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_T_" +str( for_counter)
            #     line_name[4] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_insert_line_T_" +str( for_counter)
            #     sketch_name[1] = "cut_insert_line_Sketch"
            #     unnomal_insert
            pass  # 未使用
        if gvar.StripDataList[54][g][n] > 0:
            # for for_counter in range(1, 1 + gvar.StripDataList[54][g][n]):
            #     body_name[1] = "Body.2"
            #     body_name[2] = "cut_cavity_insert"
            #     formula_name[1] = "cut_line_formula_1"
            #     element_name[1] = "down_die_plate_up_plane"
            #     line_name[2] = "plate_line_" + str(g) + "_op" +str(op_number) +  "_unnomal_cut_line_I_symmetric_" +str( for_counter)
            #     line_name[3] = "plate_line_" + str(g) + "_op" +str(op_number) + "_unnomal_cut_line_I_" +str( for_counter)
            #     line_name[4] = "plate_line_" + str(g) + "_op" +str(op_number) +  "_unnomal_cut_insert_line_I_" +str( for_counter)
            #
            #     sketch_name[1] = "cut_insert_line_Sketch"
            #
            #     unnomal_insert
            pass  # 未使用
        if gvar.StripDataList[55][g][n] > 0:
            # # for for_counter in range(1, 1 + gvar.StripDataList[55][g][n]):
            #     body_name[1] = "Body.2"
            #     body_name[2] = "cut_cavity_insert"
            #     formula_name[1] = "cut_line_formula_1"
            #     element_name[1] = "down_die_plate_up_plane"
            #     line_name[2] = "plate_line_" + str(g) + "_op" + str(op_number) + "_unnomal_cut_line_M_symmetric_" +str( for_counter)
            #     line_name[3] = "plate_line_" + str(g) + "_op" +str(op_number) + "_unnomal_cut_line_M_" +str( for_counter)
            #     line_name[4] = "plate_line_" + str(g) + "_op" +str(op_number) + "_unnomal_cut_insert_line_M_" +str( for_counter)
            #     sketch_name[1] = "cut_insert_line_Sketch"
            #     Call unnomal_insert
            pass  # 未使用
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↑I T M 對稱入子 2015-12-25
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↓靠肩沖頭(2種:衝切,凸包) 2016-8-25
        specsAndGeomWindow1 = windows1.Item("Data1" + ".CATPart")
        # --------------------------------------------------------------------------------------------整形模穴
        if gvar.StripDataList[74][g][n] > 0:
            # for now_data_number in range(1, 1 + gvar.StripDataList[74][g][n]):
            #     F_bending_cut_cavity
            pass  # 未使用
        element_name[0] = "cut_insert_line_Sketch"
        X = 0  # 名稱命名
        # --------------------------------------------------------------------------------------------異型沖頭
        if gvar.StripDataList[39][g][n] > 0:
            # for now_data_number in range(1, 1 + gvar.StripDataList[39][g][n]):
            #     #----------------------------------------------------參數
            #     emboss_forming_punch_direction = " "
            #     file_name = "op" + str(op_number) + "_allotype_cut_insert_" + str(now_data_number) + ".CATPart"
            #     insert_line_name[1] = "plate_line" + str(g) + "_op" + str(op_number) + "_allotype_cut_insert_" + str(now_data_number)
            #     #----------------------------------------------------參數
            #     #----------------------------------------------------執行
            #     partDocument2 = documents1.Open(gvar.save_path + file_name)
            #     specsAndGeomWindow2 = windows1.Item(file_name)
            #     Call F_function.projection_subroutine("cut_cavity_insert", "up_plane", element_name[0], "Sketch", insert_line_name[1])
            #     selection1.Add (body1)
            #     selection1.Paste   ()        #將副程式中F_function.projection_subroutine 所複製的cruve 複製入子
            #     specsAndGeomWindow1.Activate()
            #     Call F_function.Excavation_subroutine("Body.2", "down_die_plate_up_plane", insert_line_name[1], 0, "Last", "Regular")
            #     partDocument2.Close()
            #     #----------------------------------------------------執行
            pass  # 未使用
        # --------------------------------------------------------------------------------------------異型沖頭
        # --------------------------------------------------------------------------------------------打凸包沖頭_左
        if gvar.StripDataList[21][g][n] > 0:
            # if gvar.StripDataList[67][g][n] > 0 :
            #     for now_data_number in range(1, 1 + gvar.StripDataList[67][g][n]):
            #     #----------------------------------------------------參數
            #         emboss_forming_punch_direction = "left"
            #         file_name = "op" + str(op_number) + "_left_emboss_forming_insert_" + str(now_data_number) + ".CATPart"
            #         insert_line_name[1] = "plate_line" + str(g) + "_op" + str(op_number) + "_left_emboss_forming_insert_" + str(now_data_number)
            #         #----------------------------------------------------參數
            #         #----------------------------------------------------執行
            #         partDocument2 = documents1.Open(save_path & file_name)
            #         specsAndGeomWindow2 = windows1.Item(file_name)
            #         Call F_function.projection_subroutine("cut_cavity_insert", "up_plane", element_name[0], "Sketch", insert_line_name[1])
            #         selection1.Add (body1)
            #         selection1.Paste ()          #將副程式中F_function.projection_subroutine 所複製的cruve 複製入子
            #         specsAndGeomWindow1.Activate()
            #         Call F_function.Excavation_subroutine("Body.2", "down_die_plate_up_plane", insert_line_name[1], 0, "Last", "Regular")
            #         partDocument2.Close()
            # ----------------------------------------------------執行
            pass  # 未使用
        # --------------------------------------------------------------------------------------------打凸包沖頭_左

        # --------------------------------------------------------------------------------------------打凸包沖頭_右
        if gvar.StripDataList[22][g][n] > 0:
            # if gvar.StripDataList[71][g][n] > 0 :
            #     for now_data_number in range(1, 1 + gvar.StripDataList[71][g][n]):
            #         #----------------------------------------------------參數
            #         emboss_forming_punch_direction = "right"
            #         file_name = "op" + str(op_number) + "_right_emboss_forming_insert_" + str(now_data_number) + ".CATPart"
            #         insert_line_name[1] = "plate_line" + str(g) + "_op" + str(op_number) + "_right_emboss_forming_insert_" + str(now_data_number)
            #         #----------------------------------------------------參數
            #         #----------------------------------------------------執行
            #         partDocument2 = documents1.Open(save_path & file_name)
            #         specsAndGeomWindow2 = windows1.Item(file_name)
            #         Call F_function.projection_subroutine("cut_cavity_insert", "up_plane", element_name[0], "Sketch", insert_line_name[1])
            #         selection1.Add (body1)
            #         selection1.Paste ()          #將副程式中F_function.projection_subroutine 所複製的cruve 複製入子
            #         specsAndGeomWindow1.Activate()
            #         Call F_function.Excavation_subroutine("Body.2", "down_die_plate_up_plane", insert_line_name[1], 0, "Last", "Regular")
            #         partDocument2.Close()
            #         #----------------------------------------------------執行
            pass  # 未使用
        # --------------------------------------------------------------------------------------------打凸包沖頭_右
        if gvar.StripDataList[41][g][n] > 0:
            for i in range(1, 1 + gvar.StripDataList[41][g][n]):
                cut_cavity_machining_explanation_shape = cut_cavity_machining_explanation_shape + 1  # ----------------------------------加工說明
                # -----------------------------Boundary 取得型面外形線↓-----------------------------
                hybridShapeFactory1 = part1.HybridShapeFactory
                parameters1 = part1.Parameters
                hybridShapeSurfaceExplicit1 = parameters1.Item(
                    "plate_line_" + str(g) + "_op" + str(op_number) + "_forming_punch_surface_" + str(i))  # 型面名稱
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
                    op_number) + "_forming_punch_Boundary_" + str(i)  # 更改外形線名稱
                hybridShapeFactory1.DeleteObjectforDatum(reference2)
                # --------------隱藏Boundary型面外形線↓---------------------
                selection1 = partDocument1.Selection
                selection1.Clear()
                selection1.Search("Name=*_Boundary*, All ")
                selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
                selection1.Clear()
                # --------------隱藏Boundary型面外形線↑---------------------
                # -----------------------------Boundary 取得型面外形線↑-----------------------------
                # -----------------------------Project 投影外形線至指定平面↓-----------------------------
                hybridShapeCurveExplicit1 = parameters1.Item(
                    "plate_line_" + str(g) + "_op" + str(op_number) + "_forming_punch_Boundary_" + str(i))  # 欲投影外形線之名稱
                reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
                hybridShapes1 = body1.HybridShapes
                hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")  # 投影到哪個平面
                reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
                hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference1, reference2)
                hybridShapeProject1.SolutionType = 1  ###
                hybridShapeProject1.Normal = True
                hybridShapeProject1.SmoothingType = 0
                body1.InsertHybridShape(hybridShapeProject1)
                part1.InWorkObject = (hybridShapeProject1)
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
                # -----------------------------Project 投影外形線至指定平面↑-----------------------------
                # ----------------------------------↓OFFSET↓------------------------------------------
                reference11 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit2)
                # ---------------------------------------------------------------------------------------------------------------------------------------------------↓對稱入子
                hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)  # 改Z方向
                hybridShape3DCurveOffset4 = hybridShapeFactory1.AddNew3DCurveOffset(reference11, hybridShapeDirection2,
                                                                                    0.005, 1, 0.5)  # 偏移距離
                # ---------------------------------------------------------------------------------------------------------------------------------------------------↑對稱入子
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
                # ----------------Pad 對投影外形線挖除至型面↓--------------------------------
                partDocument1 = catapp.ActiveDocument
                part1 = partDocument1.Part
                shapeFactory1 = part1.ShapeFactory
                reference1 = part1.CreateReferenceFromName("")
                pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 40)
                parameters1 = part1.Parameters
                hybridShapeCurveExplicit1 = parameters1.Item(
                    "plate_line_" + str(g) + "_op" + str(op_number) + "_forming_punch_offset_" + str(i))
                reference2 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
                pocket1.SetProfileElement(reference2)
                reference3 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
                pocket1.SetProfileElement(reference3)
                limit1 = pocket1.FirstLimit
                limit1.LimitMode = 3
                bodies1 = part1.Bodies
                body1 = bodies1.Item("Body.2")
                hybridShapes1 = body1.HybridShapes
                hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
                reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
                limit1.LimitingElement = reference4
                part1.Update()
        # ----------------Pad 對投影外形線挖除至型面↑--------------------------------
    for n in range(1, 1 + total_op_number):
        op_number = 10 * n
        # ----------------------------------------------------------------------------------------------↓快拆沖頭
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
        if gvar.StripDataList[33][g][n] > 0:  # 成形沖頭_右
            # data_type = "surface"
            # data_number = gvar.StripDataList[33][g][n]
            # part_name = "op" + str(op_number) + "_right_quickly_remove_bending_punch_insert_"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[34][g][n] > 0:  # 成形沖頭_左
            # data_type = "surface"
            # data_number =gvar.StripDataList[34][g][n]
            # part_name = "op" + str(op_number) + "_left_quickly_remove_bending_punch_insert_"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[35][g][n] > 0:  # 成形沖頭_上
            # data_type = "surface"
            # data_number = gvar.StripDataList[35][g][n]
            # part_name = "op" + str(op_number) + "_up_quickly_remove_bending_punch_insert_"
            # quickly_remove_punch
            pass  # 未使用
        if gvar.StripDataList[36][g][n] > 0:  # 成形沖頭_下
            # data_type = "surface"
            # data_number = gvar.StripDataList[36][g][n]
            # part_name = "op" + str(op_number) + "_down_quickly_remove_bending_punch_insert_"
            # quickly_remove_punch
            pass  # 未使用
        # ----------------------------------------------------------------------------------------------↑快拆沖頭
    # --------------------------↓讓位↓---------------------------
    for o in range(1, 1 + total_op_number):
        # for j in range(1, 1 + 5):
        #     plate_op = g * 100 + o
        #     if gvar.StripDataList[44][plate_op][j] > 0:
        #         length[7] = part1.Parameters.Item("demise_height")
        #         length[7].Value = gvar.StripDataList[44][plate_op][j]
        #         # -------↓Function demise用↓---------
        #         demise_plane_h = length[7].Value
        #         data_line = j
        #         demise_op = o * 10
                # -------↑Function demise用↑---------
                # Call demise
                pass  # 未使用
    # --------------------------↑讓位↑---------------------------
    product1.PartNumber = "lower_die_" + str(g)  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("lower_die_" + str(g))
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
    strParam4.ValuateFromString(gvar.strip_parameter_list[35])
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam5 = parameters5.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(gvar.strip_parameter_list[36])
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
    strParam8.ValuateFromString("DL1 : " + str(cut_cavity_machining_explanation_shape) + "-(入塊孔, 割), 單+0.005")
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
    part1 = partDocument1.Part
    parameters12 = part1.Parameters
    length[5] = parameters12.Item("plate_length")
    plate_length = length[5].Value
    part1.Update()
    parameters13 = part1.Parameters
    length[6] = parameters13.Item("plate_width")
    plate_width = length[6].Value
    part1.Update()
    # ------------------↓隱藏點↓-------------------
    selection1 = partDocument1.Selection
    selection1.Clear()
    selection1.Search("Name=start_point*, All ")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    selection1.Clear()
    selection1.Search("Name=end_point*, All ")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    selection1.Clear()
    selection1.Search("Name=*sketch*, All ")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    # ------------------↑隱藏點↑-------------------
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
    # --------------↑刪除不需要的Data↑--------------
    part1.Update()
    time.sleep(2)
    partDocument1.SaveAs(gvar.save_path + "lower_die_" + str(g) + ".CATPart")  # 存檔的檔案名稱
    time.sleep(2)
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    partDocument1.Close()
    # ======================================(cut_cavity_Change)======================================


def A_punch(now_plate_line_number, op_number):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    length = [None] * 99
    g = now_plate_line_number
    n = int(op_number / 10)
    lower_die_A_punch_count = int()
    length[33] = part1.Parameters.Item("gap")
    length[33].Value = float(gvar.strip_parameter_list[1]) * 0.02  # (lower_die_space / 100 )
    for i in range(1,1+gvar.StripDataList[37][g][n]):
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↓對稱入子
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
        part1.Update()
        # ------------------------------------------------------------↓  修正為對稱孔位之參數 True=對稱  False=不對稱
        boolParam1 = part1.Parameters.Item("symmetry_switch")
        if selection3.Count > 0:
            boolParam1.Value = True
        else:
            boolParam1.Value = False
        part1.Update()
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   參數宣告
        length[34] = part1.Parameters.Item("x_to_x")
        length[35] = part1.Parameters.Item("y_to_y")
        length[36] = part1.Parameters.Item("int_x")
        length[37] = part1.Parameters.Item("int_y")
        q = length[34].Value
        R = length[35].Value
        # ------------------------------------------------------------↑
        # ------------------------------------------------------------↓   整數化
        length[36].Value = math.ceil(q)
        length[37].Value = math.ceil(R)
        # ------------------------------------------------------------↑
        part1.Update()
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↑對稱入子
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        hybridShapeFactory1 = part1.HybridShapeFactory
        for B_n in range(20, 1 + 1, -1):
            selection1 = partDocument1.Selection
            selection1.Clear()
            selection1.Search("Name=Body." + str(B_n) + "*,all")
            Body_n = selection1.Count
            selection1.Clear()
            if Body_n > 0:
                body_number = B_n
                break
        body2 = bodies1.Item("cut_cavity_insert")
        sketches1 = body2.Sketches
        sketch1 = sketches1.Item("cut_insert_line_Sketch")
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
        part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(
            i) + "_project_line"
        part1.Update()
        hybridShapeFactory1.DeleteObjectforDatum(reference3)
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=*_project_line,all")
        selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
        selection1.Clear()
        # ----------------------------------↑投影外型線↑--------------------------------------
        # ----------------------------------↓OFFSET↓------------------------------------------
        reference4 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↓對稱入子
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)  # 改Z方向
        hybridShape3DCurveOffset1 = hybridShapeFactory1.AddNew3DCurveOffset(reference4, hybridShapeDirection1, 0.005, 0,
                                                                            0)  # 偏移距離
        # ---------------------------------------------------------------------------------------------------------------------------------------------------↑對稱入子
        hybridShape3DCurveOffset1.InvertDirection = False
        body1.InsertHybridShape(hybridShape3DCurveOffset1)
        part1.InWorkObject = hybridShape3DCurveOffset1
        part1.Update()
        reference5 = part1.CreateReferenceFromObject(hybridShape3DCurveOffset1)
        hybridShapeCurveExplicit3 = hybridShapeFactory1.AddNewCurveDatum(reference5)
        body1.InsertHybridShape(hybridShapeCurveExplicit3)
        part1.InWorkObject = hybridShapeCurveExplicit3
        part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(
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
        lower_die_A_punch_count = lower_die_A_punch_count + 1
        pocket1.Name = "Lower-die-A-punch-" + str(lower_die_A_punch_count)
        limit1.LimitMode = 3
        hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
        reference9 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
        limit1.LimitingElement = reference9
        part1.Update()

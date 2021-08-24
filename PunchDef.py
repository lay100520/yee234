import win32com.client as win32
import global_var as gvar
import defs
import time


def bend_up_shaping_cavity_hole_1(op_number, pp_count):  # 整平模組孔down
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(
        gvar.save_path + "op" + str(op_number) + "_bend_up_shaping_punch_down_" + str(pp_count) + ".CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length1 = parameters1.Item("offset_sketch_length")
    length1.Value = -int(gvar.strip_parameter_list[20])
    part1.Update()
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    sketches1 = body1.Sketches
    sketch1 = sketches1.Item("offset_sketch")
    reference1 = part1.CreateReferenceFromObject(sketch1)
    originElements1 = part1.OriginElements
    hybridShapePlaneExplicit1 = originElements1.PlaneXY
    reference2 = part1.CreateReferenceFromObject(hybridShapePlaneExplicit1)
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
    part1.InWorkObject.Name = "offset_sketch_hole"
    part1.Update()
    hybridShapeFactory1.DeleteObjectForDatum(reference3)
    part1.InWorkObject = sketch1
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    sketches1 = sketch1.Parent
    bSTR1 = sketch1.Name
    selection1.Add(sketch1)
    visPropertySet1 = visPropertySet1.Parent
    bSTR2 = visPropertySet1.Name
    bSTR3 = visPropertySet1.Name
    visPropertySet1.SetShow(0)
    selection1.Clear()
    part1.InWorkObject = hybridShapeCurveExplicit1
    partDocument1 = catapp.ActiveDocument
    selection2 = partDocument1.Selection
    selection2.Add(hybridShapeCurveExplicit1)
    selection2.Copy()
    windows1 = catapp.Windows
    specsAndGeomWindow1 = windows1.Item("Data1.CATPart")
    specsAndGeomWindow1.Activate()
    partDocument2 = catapp.ActiveDocument
    selection3 = partDocument2.Selection
    selection3.Clear()
    part2 = partDocument2.Part
    bodies2 = part2.Bodies
    body2 = bodies2.Item("Body.2")
    selection3.Add(body2)
    selection3.Paste()
    hybridShapeFactory2 = part2.HybridShapeFactory
    parameters2 = part2.Parameters
    hybridShapeCurveExplicit2 = parameters2.Item("offset_sketch_hole")
    reference4 = part2.CreateReferenceFromObject(hybridShapeCurveExplicit2)
    hybridShapes1 = body2.HybridShapes
    hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
    reference5 = part2.CreateReferenceFromObject(hybridShapePlaneOffset1)
    hybridShapeProject2 = hybridShapeFactory2.AddNewProject(reference4, reference5)
    hybridShapeProject2.SolutionType = 0
    hybridShapeProject2.Normal = True
    hybridShapeProject2.SmoothingType = 0
    body2.InsertHybridShape(hybridShapeProject2)
    part2.InWorkObject = hybridShapeProject2
    part2.Update()
    reference6 = part2.CreateReferenceFromObject(hybridShapeProject2)
    hybridShapeCurveExplicit3 = hybridShapeFactory2.AddNewCurveDatum(reference6)
    body2.InsertHybridShape(hybridShapeCurveExplicit3)
    part2.InWorkObject = hybridShapeCurveExplicit3
    part2.Update()
    hybridShapeFactory2.DeleteObjectForDatum(reference6)
    shapeFactory1 = part2.ShapeFactory
    reference7 = part2.CreateReferenceFromName("")
    pocket1 = shapeFactory1.AddNewPocketFromRef(reference7, 20)
    reference8 = part2.CreateReferenceFromObject(hybridShapeCurveExplicit3)
    pocket1.SetProfileElement(reference8)
    reference9 = part2.CreateReferenceFromObject(hybridShapeCurveExplicit3)
    pocket1.SetProfileElement(reference9)
    limit1 = pocket1.FirstLimit
    limit1.LimitMode = 3
    hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
    reference10 = part2.CreateReferenceFromObject(hybridShapePlaneOffset2)
    limit1.LimitingElement = reference10
    part2.UpdateObject(pocket1)
    part2.Update()
    partDocument2 = catapp.ActiveDocument
    selection4 = partDocument2.Selection
    selection4.Clear()
    selection4.Add(hybridShapeCurveExplicit2)
    selection4.Delete()
    specsAndGeomWindow2 = windows1.Item(
        "op" + str(op_number) + "_bend_up_shaping_punch_down_" + str(pp_count) + ".CATPart")
    specsAndGeomWindow2.Activate()
    specsAndGeomWindow2.Close()
    partDocument1 = catapp.ActiveDocument
    partDocument1.Close()


def bend_up_shaping_cavity_hole_2(op_number, pp_count):  # 整平模組孔up
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(
        gvar.save_path + "op" + str(op_number) + "_bend_up_shaping_punch_up_" + str(pp_count) + ".CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length1 = parameters1.Item("offset_sketch_length")
    length1.Value = -int(gvar.strip_parameter_list[20])
    part1.Update()
    hybridShapeFactory1 = part1.HybridShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    sketches1 = body1.Sketches
    sketch1 = sketches1.Item("offset_sketch")
    reference1 = part1.CreateReferenceFromObject(sketch1)
    originElements1 = part1.OriginElements
    hybridShapePlaneExplicit1 = originElements1.PlaneXY
    reference2 = part1.CreateReferenceFromObject(hybridShapePlaneExplicit1)
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
    part1.InWorkObject.Name = "offset_sketch_hole"
    part1.Update()
    hybridShapeFactory1.DeleteObjectForDatum(reference3)
    part1.InWorkObject = sketch1
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    sketches1 = sketch1.Parent
    bSTR1 = sketch1.Name
    selection1.Add(sketch1)
    visPropertySet1 = visPropertySet1.Parent
    bSTR2 = visPropertySet1.Name
    bSTR3 = visPropertySet1.Name
    visPropertySet1.SetShow(0)
    selection1.Clear()
    part1.InWorkObject = hybridShapeCurveExplicit1
    partDocument1 = catapp.ActiveDocument
    selection2 = partDocument1.Selection
    selection2.Add(hybridShapeCurveExplicit1)
    selection2.Copy()
    windows1 = catapp.Windows
    specsAndGeomWindow1 = windows1.Item("Data1.CATPart")
    specsAndGeomWindow1.Activate()
    partDocument2 = catapp.ActiveDocument
    selection3 = partDocument2.Selection
    selection3.Clear()
    part2 = partDocument2.Part
    bodies2 = part2.Bodies
    body2 = bodies2.Item("Body.2")
    selection3.Add(body2)
    selection3.Paste()
    hybridShapeFactory2 = part2.HybridShapeFactory
    parameters2 = part2.Parameters
    hybridShapeCurveExplicit2 = parameters2.Item("offset_sketch_hole")
    reference4 = part2.CreateReferenceFromObject(hybridShapeCurveExplicit2)
    hybridShapes1 = body2.HybridShapes
    hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
    reference5 = part2.CreateReferenceFromObject(hybridShapePlaneOffset1)
    hybridShapeProject2 = hybridShapeFactory2.AddNewProject(reference4, reference5)
    hybridShapeProject2.SolutionType = 0
    hybridShapeProject2.Normal = True
    hybridShapeProject2.SmoothingType = 0
    body2.InsertHybridShape(hybridShapeProject2)
    part2.InWorkObject = hybridShapeProject2
    part2.Update()
    reference6 = part2.CreateReferenceFromObject(hybridShapeProject2)
    hybridShapeCurveExplicit3 = hybridShapeFactory2.AddNewCurveDatum(reference6)
    body2.InsertHybridShape(hybridShapeCurveExplicit3)
    part2.InWorkObject = hybridShapeCurveExplicit3
    part2.Update()
    hybridShapeFactory2.DeleteObjectForDatum(reference6)
    shapeFactory1 = part2.ShapeFactory
    reference7 = part2.CreateReferenceFromName("")
    pocket1 = shapeFactory1.AddNewPocketFromRef(reference7, 20)
    reference8 = part2.CreateReferenceFromObject(hybridShapeCurveExplicit3)
    pocket1.SetProfileElement(reference8)
    reference9 = part2.CreateReferenceFromObject(hybridShapeCurveExplicit3)
    pocket1.SetProfileElement(reference9)
    limit1 = pocket1.FirstLimit
    limit1.LimitMode = 3
    hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
    reference10 = part2.CreateReferenceFromObject(hybridShapePlaneOffset2)
    limit1.LimitingElement = reference10
    part2.UpdateObject(pocket1)
    part2.Update()
    partDocument2 = catapp.ActiveDocument
    selection4 = partDocument2.Selection
    selection4.Clear()
    selection4.Add(hybridShapeCurveExplicit2)
    selection4.Delete()
    specsAndGeomWindow2 = windows1.Item(
        "op" + str(op_number) + "_bend_up_shaping_punch_up_" + str(pp_count) + ".CATPart")
    specsAndGeomWindow2.Activate()
    specsAndGeomWindow2.Close()
    partDocument1 = catapp.ActiveDocument
    partDocument1.Close()


def punch_Reinforcement_Ecxavation(g, n, i, parameter_digital1):  # 補強入子
    length = [] * 11
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    formula_name1 = "cut_line_formula_3"
    line_name1 = "plate_line_" + str(g) + "_op" + str(n * 10) + "_Reinforcement_cut_line_" + str(i)
    line_name2 = "plate_line_" + str(g) + "_op" + str(n * 10) + "_Reinforcement_cut_Ecxavation_line_" + str(i)
    body_name1 = "Body.2"
    body_name2 = "Reinforcement_cut_punch"
    sketch_name1 = "Bolt_Sketch"
    element_name1 = "down_die_plate_up_plane"
    defs.FormulaChange(body_name2, formula_name1, line_name1, sketch_name1)
    # ---------改變參數↓-------------
    strParam1 = parameters1.Item("QR_type")
    strParam1.Value = "QR_B"
    # ---------改變參數↑-------------
    length[1] = part1.Parameters.Item("QR_wide")
    length[2] = part1.Parameters.Item("Reinforcement_Grow_wide_2")
    if length[1].Value < 10:
        length[2].Value = 10 - length[1].Value
        strParam1.Value = "Reinforcement_C"
        sketch_name2 = "C_pad_Excavation_Sketch"
    else:
        strParam1.Value = "QR_B"
        sketch_name2 = "B_pad_Excavation_Sketch"
    length[0] = part1.Parameters.Item("Reinforcement_Grow_wide_1")
    length[0].Value = 20
    length[0] = part1.Parameters.Item("Reinforcement_Grow_long_1")
    length[0].Value = 46
    length[0] = part1.Parameters.Item("Reinforcement_Grow_long_2")
    length[0].Value = 25
    length[0] = part1.Parameters.Item("gap")
    length[0].Value = parameter_digital1
    length[0] = part1.Parameters.Item("Bolt_long")
    if strParam1.Value == "QR_B":
        length[0].Value = 9
    else:
        length[0].Value = 9 + length[2].Value
    part1.Update()
    defs.F_projection(element_name1, sketch_name2, line_name2, body_name1, body_name2)
    defs.F_Excavation(element_name1, line_name2, body_name1)
    part1.Update()


def punch_d_cutting(g, n):  # 切斷沖頭_下
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    op_number = n * 10
    stop_plate_machining_explanation_shape = int()
    for i in range(1, int(gvar.StripDataList[27][g][n] + 1)):
        stop_plate_machining_explanation_shape = stop_plate_machining_explanation_shape + 1  # ------------------------------------加工說明
        # ----------cut_line_number(h) > 0 ↓------------------------
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        hybridShapeFactory1 = part1.HybridShapeFactory
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_d_cutting_" + str(i))
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
        part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_d_cutting_" + str(
            i) + "_project_line"  # 更改外形線名稱
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
        hybridShape3DCurveOffset1 = hybridShapeFactory1.AddNew3DCurveOffset(reference4, hybridShapeDirection1, 0, 1,
                                                                            0.5)  # 偏移距離
        hybridShape3DCurveOffset1.InvertDirection = False
        body1.InsertHybridShape(hybridShape3DCurveOffset1)
        part1.InWorkObject = hybridShape3DCurveOffset1
        part1.Update()
        reference5 = part1.CreateReferenceFromObject(hybridShape3DCurveOffset1)
        hybridShapeCurveExplicit3 = hybridShapeFactory1.AddNewCurveDatum(reference5)
        body1.InsertHybridShape(hybridShapeCurveExplicit3)
        part1.InWorkObject = hybridShapeCurveExplicit3
        part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_d_cutting_" + str(
            i) + "_offset_line"
        part1.Update()
        hybridShapeFactory1.DeleteObjectForDatum(reference5)
        shapeFactory1 = part1.ShapeFactory
        # ------------------------------------------------------
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
        part1.Update()


def punch_u_cutting(g, n):  # 切斷沖頭_上
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    op_number = n * 10
    stop_plate_machining_explanation_shape = int()
    for i in range(1, int(gvar.StripDataList[28][g][n]) + 1):
        stop_plate_machining_explanation_shape = stop_plate_machining_explanation_shape + 1  # ------------------------------------加工說明
        # ----------cut_line_number(h) > 0 ↓------------------------
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        hybridShapeFactory1 = part1.HybridShapeFactory
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_u_cutting_" + str(i))
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
        part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_u_cutting_" + str(
            i) + "_project_line"  # 更改外形線名稱
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
        hybridShape3DCurveOffset1 = hybridShapeFactory1.AddNew3DCurveOffset(reference4, hybridShapeDirection1, 0, 1,
                                                                            0.5)  # 偏移距離
        hybridShape3DCurveOffset1.InvertDirection = False
        body1.InsertHybridShape(hybridShape3DCurveOffset1)
        part1.InWorkObject = hybridShape3DCurveOffset1
        part1.Update()
        reference5 = part1.CreateReferenceFromObject(hybridShape3DCurveOffset1)
        hybridShapeCurveExplicit3 = hybridShapeFactory1.AddNewCurveDatum(reference5)
        body1.InsertHybridShape(hybridShapeCurveExplicit3)
        part1.InWorkObject = hybridShapeCurveExplicit3
        part1.InWorkObject.Name = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_u_cutting_" + str(
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
        part1.Update()
        # ------------------------------------------------------





def emboss_forming_punch_right(g, n):  # 打凸包沖頭_右
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    op_number = n * 10
    for i in range(1, int(gvar.StripDataList[22][g][n]) + 1):
        partDocument1 = documents1.Open(
            gvar.save_path & "op" + str(op_number) + "_emboss_forming_punch_right_0" + str(i) + ".CATPart")
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
        hybridShape3DCurveOffset1 = hybridShapes1.Item("demise_hole_right_offset")
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
        part1.InWorkObject.Name = "demise_hole_right_offset_project"
        part1.Update()
        hybridShapeFactory1.DeleteObjectforDatum(reference3)
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=demise_hole_right_offset_project*,all")
        selection1.VisProperties.SetShow(1)  # 1隱藏/0顯示
        selection1.Clear()
        # ------------------------------------------------------------↑offset
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        selection1 = partDocument1.Selection
        selection1.Clear()
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("demise_hole_right_offset_project")
        selection1.Add(hybridShapeCurveExplicit1)
        selection1.Copy()
        # ===================================================================
        windows1 = catapp.Windows
        specsAndGeomWindow1 = windows1.Item("Data1.CATPart")
        specsAndGeomWindow1.Activate()
        # ===================================================================
        # ---------------------------------------------
        # 在CATIA上切換各視窗
        partDocument2 = catapp.ActiveDocument
        selection2 = partDocument2.Selection
        selection2.Clear()
        part2 = partDocument2.Part
        bodies1 = part2.Bodies
        body1 = bodies1.Item("Body.2")
        part2.InWorkObject = body1
        selection2.Add(part2)
        selection2.Paste()
        part2.InWorkObject.Name = "demise_hole_right_offset_project_" + str(i)
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
        hybridShapeCurveExplicit1 = parameters1.Item("demise_hole_right_offset_project_" + str(i))
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


def QR_half_cut_punch(g, n):  # 半沖切
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    documents1 = catapp.Documents
    partDocument2 = documents1.Open(gvar.open_path + "QR_half_cut_punch_line.CATPart")
    defs.window_change(partDocument1, partDocument2)
    length = [] * 30
    # ======================================================================================================
    length[0] = part1.Parameters.Item("QR_half_punch_up_plane")
    if gvar.Mold_status == "開模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + int(gvar.strip_parameter_list[20]) + int(
            gvar.strip_parameter_list[17]) + int(gvar.strip_parameter_list[14]) + 28  # (upper_die_open_height)
    elif gvar.Mold_status == "閉模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + int(gvar.strip_parameter_list[20]) + int(
            gvar.strip_parameter_list[17]) + int(gvar.strip_parameter_list[14])
    # ======================================================================================================
    # ======================================================================================================
    length[1] = part1.Parameters.Item("QR_half_punch_height")
    for i in range(1, int(gvar.StripDataList[3][g][n]) + 1):
        die_rule_file_name = "沖頭切入深度"
        Row_string_serch = "精密級"  # ---------X
        Column_string_serch = "合金工具鋼"  # ---------------Y
        Thickness = float(gvar.strip_parameter_list[1])
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
        length[1].Value = Thickness + float(gvar.strip_parameter_list[1]) + int(gvar.strip_parameter_list[20]) + int(
            gvar.strip_parameter_list[17]) + int(gvar.strip_parameter_list[14]) + serch_result  # 沖頭高度
        length[2] = part1.Parameters.Item("QR_half_punch_Hanging_Desk_height")
        length[2].Value = int(gvar.strip_parameter_list[14])
        part1.Parameters.Item("half_cut_cut_line_formula_" + str(i)).OptionalRelation.Modify(
            "die\\plate_line_" + str(g) + "_op" + str(n * 10) + "_half_cut_line_" + str(i))  # 草圖置換
        part1.Update()
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        hybridShapeFactory1 = part1.HybridShapeFactory
        for B_n in range(60, 0, -1):
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
        sketch1 = sketches1.Item("stop_plate_cut_line_" + str(i))
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
        hybridShapeFactory1.DeleteObjectForDatum(reference3)
        shapeFactory1 = part1.ShapeFactory
        reference4 = part1.CreateReferenceFromName("")
        pocket1 = shapeFactory1.AddNewPocketFromRef(reference4, 100)
        reference5 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        pocket1.SetProfileElement(reference5)
        reference6 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        pocket1.SetProfileElement(reference6)
        limit1 = pocket1.FirstLimit
        limit1.LimitMode = 2
        part1.Update()


def quickly_remove_punch(data_type, data_number, part_name):  # 快拆沖頭
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    documents1 = catapp.Documents
    parameter_name = [] * 10
    stop_plate_machining_explanation_shape = int()
    for i in range(1, data_number + 1):
        stop_plate_machining_explanation_shape = stop_plate_machining_explanation_shape + 1  # ------------------------------------加工說明
        partDocument1 = documents1.Open(gvar.save_path + part_name + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        if data_type == "line":
            data_type = "cut_punch_stop_plate_demise_Sketch"
            parameter_name[1] = "gap"
        elif data_type == "surface":
            data_type = "forming_punch_stop_plate_demise_Sketch"
            parameter_name[1] = "forming_punch_demise_gap"
        length = [] * 4
        # ================================================
        length[1] = part1.Parameters.Item("line_formula_1_Xmax_length")
        length[2] = part1.Parameters.Item("formula_Sketch_Xmax_length")
        # ================================================
        length[3] = part1.Parameters.Item(parameter_name[1])
        length[3].Value = 0
        part1.Update()
        if length[2].Value < length[1].Value:
            length[3].Value = -length[3].Value
        # ================================================
        part1.UpdateObject(part1.Bodies.Item("Body.2"))
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        # ---------------------------------------------------↓project
        hybridShapeFactory1 = part1.HybridShapeFactory
        sketches1 = body1.Sketches
        sketch1 = sketches1.Item(data_type)
        reference1 = part1.CreateReferenceFromObject(sketch1)
        hybridShapes1 = body1.HybridShapes
        hybridShapePlaneOffset1 = hybridShapes1.Item("punch_up_plane")
        reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
        hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference1, reference2)
        hybridShapeProject1.SolutionType = 0
        hybridShapeProject1.Normal = True
        hybridShapeProject1.SmoothingType = 0
        body1.InsertHybridShape(hybridShapeProject1)
        part1.InWorkObject = hybridShapeProject1
        part1.Update()
        # --------------------------------------------------------------------------------↓打斷關聯
        reference3 = part1.CreateReferenceFromObject(hybridShapeProject1)
        hybridShapeCurveExplicit1 = hybridShapeFactory1.AddNewCurveDatum(reference3)
        body1.InsertHybridShape(hybridShapeCurveExplicit1)
        part1.InWorkObject = hybridShapeCurveExplicit1
        hybridShapeCurveExplicit1.Name = part_name + str(i) + "_demise_line_" + str(i)
        part1.Update()
        hybridShapeFactory1.DeleteObjectforDatum(reference3)
        # -----------------------------------------------------------↓隱藏
        partDocument1.Selection.Add(hybridShapeCurveExplicit1)
        visPropertySet1 = partDocument1.Selection.VisProperties
        visPropertySet1 = visPropertySet1.Parent
        visPropertySet1.SetShow(1)  # 1為隱藏,0為顯示
        # -----------------------------------------------------------↑
        # ----------------------------------------↓複製
        partDocument1 = catapp.ActiveDocument
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Add(hybridShapeCurveExplicit1)
        selection1.Copy()
        # ---------------------------------------------------------↓切換視窗
        windows1 = catapp.Windows
        specsAndGeomWindow1 = windows1.Item("Data1.CATPart")
        specsAndGeomWindow1.Activate()
        viewer3D1 = specsAndGeomWindow1.ActiveViewer
        viewpoint3D1 = viewer3D1.Viewpoint3D
        # ------------------------------------------↓貼上
        partDocument2 = catapp.ActiveDocument
        part1 = partDocument2.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        selection2 = partDocument2.Selection
        selection2.Clear()
        part2 = partDocument2.Part
        selection2.Add(part2)
        selection2.Paste()
        partDocument1.Close()
        # -------------------------------------------------------------------------------------↓挖除
        shapeFactory1 = part2.ShapeFactory
        reference4 = part2.CreateReferenceFromName("")
        pocket1 = shapeFactory1.AddNewPocketFromRef(reference4, 20)
        parameters1 = part2.Parameters
        hybridShapeCurveExplicit2 = parameters1.Item(part_name + str(i) + "_demise_line_" + str(i))
        reference5 = part2.CreateReferenceFromObject(hybridShapeCurveExplicit2)
        pocket1.SetProfileElement(reference5)
        pocket1.Name = part_name + str(i) + "_demise_Pocket_" + str(i)
        limit1 = pocket1.FirstLimit
        limit1.LimitMode = 3
        hybridShapes2 = body1.HybridShapes
        hybridShapePlaneOffset2 = hybridShapes2.Item("down_die_plate_up_plane")
        reference6 = part2.CreateReferenceFromObject(hybridShapePlaneOffset2)
        limit1.LimitingElement = reference6
        part2.Update()


def formula_change(body_name, line_name1, sketch_name,formula_name1):
    # ------------------------------------------------------------↓基本宣告
    catapp = win32.Dispatch('CATIA.Application')
    documents = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    bodies1 = part1.Bodies
    body1 = bodies1.Item(body_name[2])
    cut_cavity_machining_explanation_shape = int()
    # ------------------------------------------------------------↑
    cut_cavity_machining_explanation_shape = cut_cavity_machining_explanation_shape + 1  # --------------------------------------------------加工說明
    # ------------------------------------------------------------↓置換草圖
    parameters1.Item(formula_name1).OptionalRelation.Modify("die\\" + line_name1)  # 草圖置換
    # ------------------------------------------------------------↑
    # ------------------------------------------------------------↓宣告草圖
    sketches1 = body1.Sketches
    sketch1 = sketches1.Item(sketch_name[1])
    part1.UpdateObject(sketch1)  # 單步對草圖進行更新
    # ------------------------------------------------------------↑


def clash_change(xi, op_number):
    if xi != 0:
        catapp = win32.Dispatch('CATIA.Application')
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        bodies1 = part1.Bodies
        bodies2 = part1.Bodies
        body2 = bodies2.Item("Body.2")
        sketches1 = body2.Sketches
        sketch1 = sketches1.Item("Sketch.39")
        part1.InWorkObject = sketch1
        factory2D1 = sketch1.OpenEdition()
        # numeral As Integer  #--------------------------------------------定義
        # j = 20 #----------------------------------------------------變更條件
        numeral = op_number
        for i in range(2, 1 + xi):
            parameters2 = part1.Parameters  # -----------------------------------------
            reference1 = parameters2.Item("plate_line_1_op" + str(numeral) + "_cut_line_" + str(i))
            geometricElements1 = factory2D1.CreateProjections(reference1)
        hybridBodies2 = part1.HybridBodies
        hybridBody2 = hybridBodies2.Item("cut_line_assume1")
        part1.InWorkObject = hybridBody2
        part1.Update()


def A_punch_clash_change(xi, op_number):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    bodies2 = part1.Bodies
    body2 = bodies2.Item("Body.2")
    sketches1 = body2.Sketches
    sketch1 = sketches1.Item("insert_pocket_sketch")
    part1.InWorkObject = sketch1
    factory2D1 = sketch1.OpenEdition()
    numeral = op_number
    for i in range( 2 , 1+ xi):
        parameters2 = part1.Parameters #-----------------------------------------
        reference1 = parameters2.Item("plate_line_1_op" + str(numeral) + "_A_punch_" + str(i))
        geometricElements1 = factory2D1.CreateProjections(reference1)
    hybridBodies2 = part1.HybridBodies
    hybridBody2 = hybridBodies2.Item("cut_line_assume1")
    part1.InWorkObject = hybridBody2
    part1.Update()


def A_punch_clash_change_QR_Splint(xi, op_number):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    bodies2 = part1.Bodies
    body2 = bodies2.Item("Body.2")
    sketches1 = body2.Sketches
    sketch1 = sketches1.Item("insert_hole_sketch")
    part1.InWorkObject = sketch1
    factory2D1 = sketch1.OpenEdition()
    numeral = op_number
    for i in range (2 , 1+ xi):
        parameters1 = part1.Parameters #-----------------------------------------
        reference1 = parameters1.Item("plate_line_1_op" + str(numeral) + "_A_punch_" + str(i))
        geometricElements1 = factory2D1.CreateProjections(reference1)
        sketch1.CloseEdition()
        sketch2 = sketches1.Item("insert_pocket_sketch_1")
        part1.InWorkObject = sketch2
        factory2D2 = sketch2.OpenEdition()
        geometricElements2 = factory2D2.CreateProjections(reference1)
        point2D1 = factory2D2.CreatePoint(0, 0)
        length1 = parameters1.Item("Body.2\insert_pocket_sketch_1\Radius.22\Radius")
        d = length1.Value
        circle2D1 = factory2D2.CreateClosedCircle(0, 0, d)
        circle2D1.CenterPoint = point2D1
        geometry2D1 = geometricElements2.Item("Mark.1")
        geometry2D1.Construction = True
        constraints1 = sketch2.Constraints
        reference2 = part1.CreateReferenceFromObject(point2D1)
        reference3 = part1.CreateReferenceFromObject(geometry2D1)
        constraint1 = constraints1.AddBiEltCst(3, reference2, reference3)
        constraint1.mode = 0
        sketch2.CloseEdition()
        sketch3 = sketches1.Item("insert_pocket_sketch_2")
        part1.InWorkObject = sketch3
        factory2D3 = sketch3.OpenEdition()
        geometricElements3 = factory2D3.CreateProjections(reference1)
        point2D2 = factory2D3.CreatePoint(0, 0)
        length2 = parameters1.Item("Body.2\insert_pocket_2\insert_pocket_sketch_2\Radius.40\Radius")
        d = length2.Value
        circle2D2 = factory2D3.CreateClosedCircle(0, 0, d)
        circle2D2.CenterPoint = point2D2
        geometry2D2 = geometricElements3.Item("Mark.1")
        geometry2D2.Construction = True
        constraints2 = sketch3.Constraints
        reference4 = part1.CreateReferenceFromObject(point2D2)
        reference5 = part1.CreateReferenceFromObject(geometry2D2)
        constraint2 = constraints2.AddBiEltCst(3, reference4, reference5)
        constraint1.mode = 0
        sketch3.CloseEdition()
    hybridBodies2 = part1.HybridBodies
    hybridBody2 = hybridBodies2.Item("cut_line_assume")
    part1.InWorkObject = hybridBody2
    part1.Update()


def A_punch_clash_change_QR_Stripper(xi, op_number):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    bodies2 = part1.Bodies
    body2 = bodies2.Item("Body.2")
    sketches1 = body2.Sketches
    sketch1 = sketches1.Item("insert_hole_sketch_1")
    part1.InWorkObject = sketch1
    factory2D1 = sketch1.OpenEdition()
    #numeral As Integer  #--------------------------------------------定義
    #j = 20 #----------------------------------------------------變更條件
    #if j = 2 :
    numeral = op_number
    for i in range (2 , 1+ xi):
        parameters1 = part1.Parameters #-----------------------------------------
        reference1 = parameters1.Item("plate_line_1_op" + str(numeral) + "_A_punch_" + str(i))
        geometricElements1 = factory2D1.CreateProjections(reference1)
        sketch1.CloseEdition()
        hybridShapes1 = body2.HybridShapes
        hybridShape1 = hybridShapes1.Item("down_plane")
        element_Document = partDocument1
        element_body = body2
        Sketch_position = "Body"
        (element_sketch1)=defs.BuildSketch("A_punch_insert_punch", hybridShape1,element_Document,Sketch_position,element_body,None)     #(sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
        sketch2 = sketches1.Item("A_punch_insert_punch")
        part1.InWorkObject = sketch2
        factory2D2 = sketch2.OpenEdition()
        geometricElements2 = factory2D2.CreateProjections(reference1)
        point2D1 = factory2D2.CreatePoint(0, 0)
        length1 = parameters1.Item("Body.2\insert_hole_sketch_2\Radius.128\Radius")
        d = length1.Value
        circle2D1 = factory2D2.CreateClosedCircle(0, 0, d)
        circle2D1.CenterPoint = point2D1
        geometry2D1 = geometricElements2.Item("Mark.1")
        geometry2D1.Construction = True
        constraints1 = sketch2.Constraints
        reference2 = part1.CreateReferenceFromObject(point2D1)
        reference3 = part1.CreateReferenceFromObject(geometry2D1)
        constraint1 = constraints1.AddBiEltCst(3, reference2, reference3)
        constraint1.mode = 0
        sketch2.CloseEdition()
        shapeFactory1 = part1.ShapeFactory
        pocket1 = shapeFactory1.AddNewPocket(sketch2, 20)
        limit1 = pocket1.FirstLimit
        limit1.LimitMode = 5
        part1.Update()
        hybridShapeFactory1 = part1.HybridShapeFactory
        reference4 = part1.CreateReferenceFromObject(sketch2)
        hybridShapePointCenter1 = hybridShapeFactory1.AddNewPointCenter(reference4)
        body2.InsertHybridShape (hybridShapePointCenter1)
        part1.InWorkObject = hybridShapePointCenter1
        part1.Update()
        reference5 = part1.CreateReferenceFromObject(hybridShapePointCenter1)
        hybridShapePlaneOffset1 = hybridShapes1.Item("up_plane")
        reference6 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
        length1 = parameters1.Item("Body.2\insert_hole\HoleLimit.17\Depth")
        d = length1.Value
        hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference5, reference6, d)
        hole1.Type = 0
        hole1.AnchorMode = 0
        hole1.BottomType = 1
        limit1 = hole1.BottomLimit
        limit1.LimitMode = 0
        length1 = parameters1.Item("Body.2\insert_hole\Diameter")
        d = length1.Value
        length2 = hole1.Diameter
        length2.Value = d
        angle1 = hole1.BottomAngle
        angle1.Value = 120
        hole1.ThreadingMode = 1
        hole1.ThreadSide = 0
        part1.Update()
    hybridBodies2 = part1.HybridBodies
    hybridBody2 = hybridBodies2.Item("cut_line_assume")
    part1.InWorkObject = hybridBody2
    selection1 = partDocument1.Selection
    selection1.Search ("name=*Sketch*,all")
    selection1.VisProperties.SetShow (1)
    selection1.Search ("name=*Point*,all")
    selection1.VisProperties.SetShow (1)
    selection1.Clear()
    part1.Update()


def cut_cavity_clash_change(xi, op_number):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    bodies2 = part1.Bodies
    body2 = bodies2.Item("cut_cavity_insert")
    sketches1 = body2.Sketches
    sketch1 = sketches1.Item("Sketch.40")
    part1.InWorkObject = sketch1
    factory2D1 = sketch1.OpenEdition()
    # numeral As Integer  #--------------------------------------------定義
    #j = 20 #----------------------------------------------------變更條件
    #if j = 2 :
    numeral = op_number
    for i in range( 2 , 1+ xi):
        parameters2 = part1.Parameters #-----------------------------------------
        reference1 = parameters2.Item("plate_line_1_op" + str(numeral) + "_cut_line_" + str(i))
        geometricElements1 = factory2D1.CreateProjections(reference1)
    hybridBodies2 = part1.HybridBodies
    hybridBody2 = hybridBodies2.Item("cut_line_assume1")
    part1.InWorkObject = hybridBody2
    part1.Update()

def cut_plate_clash_change(xi, op_number):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    bodies2 = part1.Bodies
    body2 = bodies2.Item("cut_cavity_insert")
    sketches1 = body2.Sketches
    sketch1 = sketches1.Item("Sketch.40")
    part1.InWorkObject = sketch1
    factory2D1 = sketch1.OpenEdition()
    numeral = op_number
    for i in range( 1 , 1+ xi):
        parameters2 = part1.Parameters #-----------------------------------------
        reference1 = parameters2.Item("plate_line_1_op" + str(numeral) + "_cut_line_" + str(i))
        geometricElements1 = factory2D1.CreateProjections(reference1)
        selection1 = partDocument1.Selection
        selection1.Add (geometricElements1)
        time.sleep(0.5)
    part1.Update()
    hybridBodies2 = part1.HybridBodies
    hybridBody2 = hybridBodies2.Item("cut_line_assume1")
    part1.InWorkObject = hybridBody2
    part1.Update()

import win32com.client as win32
import global_var as gvar
import defs
import PunchDef
import time


def UpPlate(now_plate_line_number):
    total_op_number = int(gvar.strip_parameter_list[2])
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument2 = catapp.ActiveDocument
    partDocument1 = documents1.Open(gvar.open_path + "up_plate.CATPart")
    # ======================================
    defs.window_change(partDocument2, partDocument1)  # 在CATIA上切換各視窗
    # ======================================
    selection3 = partDocument2.Selection
    selection3.Clear()
    # -----------------------------------------------------------------------------------------------↓補強入子
    selection3.Search("Name=plate_line_" + str(now_plate_line_number) + "_op*_Reinforcement_cut_line_*,all")
    if selection3.Count > 0:
        partDocument1 = documents1.Open(gvar.open_path + "QR_punch_Reinforcement.CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("cut_line_formula_1")
        hybridShapeCurveExplicit1.Name = "cut_line_formula_3"
        # ======================================
        defs.window_change(partDocument2, partDocument1)  # 在CATIA上切換各視窗
        # ======================================
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
    # -----------------------------------------------------------------------------------------------↑補強入子
    # ======================================   up_plate_change
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    length = [None] * 7
    formula = [None] * 7
    length[0] = part1.Parameters.Item("pitch")
    length[0].Value = float(gvar.strip_parameter_list[4])
    # ======================================================================================================
    length[1] = part1.Parameters.Item("plate_height")
    g = now_plate_line_number
    point_conter = [0] * (g + 1)
    try:
        length[1].Value = float(gvar.strip_parameter_list[11])
    except:
        die_rule_file_name = "模板厚度選擇"
        excel_Sheet_name = "200T以下"
        Column_string_serch = "上墊板"
        Row_string_serch = "精密級"
        (serch_result) = defs.ExcelSearch(die_rule_file_name, excel_Sheet_name, Row_string_serch, Column_string_serch)
        length[1].Value = serch_result
        gg = serch_result
        gvar.strip_parameter_list[11] = gg
    length[4] = part1.Parameters.Item("plate_down_plane")
    if gvar.die_type == "module":
        plate_position = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[14]) + 0 + 0  # (back_stripper_plate_height + back_splint_height)
    else:
        plate_position = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14])
    if gvar.Mold_status == "開模":
        length[4].Value = plate_position + 28  # (upper_die_open_height)
    elif gvar.Mold_status == "閉模":
        length[4].Value = plate_position
    file_name = "Data1"
    body_name1 = "Body.2"
    hybridBody_name = "die"
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(file_name, body_name1,
                                                                                             hybridBody_name)
    if gvar.die_type == "module":
        M_plate_length = 35
        M_plate_wide = 176
        hybridShape1 = ElementBody.HybridShapes.Item("down_die_plate_up_plane")  # 宣告平面(下)
        part1.Update()
        defs.material_tpye_palte_sketch(hybridShape1, M_plate_length, M_plate_wide,
                                        100 + 50 * (now_plate_line_number - 1),
                                        112.5 + 12.5, ElementDocument, ElementBody,
                                        ElementHybridBody)  # (平面的宣告,模板長,模板寬,中心座標_X,中心座標_Y)
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify("die\\plate_size")  # 草圖置換
        part1.Update()
        M_upper_design(M_plate_length, M_plate_wide)
    else:
        part1.Parameters.Item("1_formula_1").OptionalRelation.Modify(
            "die\\number_" + str(g) + "_plate_line")  # 草圖置換
    part1.Update()
    for n in range(1, total_op_number + 1):
        op_number = 10 * n
        # ---------------------------------------------------------------------------------↓ 補強入子
        if round(gvar.StripDataList[14][g][n]) > 0:
            for for_counter in range(1, round(gvar.StripDataList[14][g][n]) + 1):
                (up_pad_Bolt_Hole) = punch_Reinforcement_Hole(g, n, for_counter)
                point_conter[g] = up_pad_Bolt_Hole
        # ---------------------------------------------------------------------------------↑ 補強入子
        if round(gvar.StripDataList[21][g][n]) > 0:  # 打凸包沖頭_左
            emboss_forming_punch_left(g, n)
        if round(gvar.StripDataList[22][g][n]) > 0:  # 打凸包沖頭_右
            emboss_forming_punch_right(g, n)
        if round(gvar.StripDataList[3][g][n]) > 0:  # 半沖切
            QR_half_cut_punch(g, n)
        if round(gvar.StripDataList[42][g][n]) > 0:  # 整平模組孔down
            for j in range(1, gvar.StripDataList[42][g][n] + 1):
                pp_count = j
                bend_up_shaping_cavity_hole_1(op_number, pp_count)
        if round(gvar.StripDataList[43][g][n]) > 0:  # 整平模組孔up
            for j in range(1, round(gvar.StripDataList[43][g][n]) + 1):
                pp_count = j
                bend_up_shaping_cavity_hole_2(op_number, pp_count)
        if round(gvar.StripDataList[73][g][n]) > 0:  # 整形沖頭
            for now_data_number in range(1, gvar.StripDataList[73][g][n]):
                F_bending_up_plate(g, op_number, now_data_number)
    selection1 = partDocument1.Selection
    selection1.Clear()
    selection1.Search("Name=bolt_line*,all")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    selection1.Search("Name=Sketch*,all")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    selection1.Search("Name=Point*,all")
    selection1.VisProperties.SetShow(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    part1.Update()
    product1.PartNumber = "up_plate_" + str(g)  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("up_plate_" + str(g))
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
    strParam4.ValuateFromString(gvar.strip_parameter_list[12])
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam5 = parameters5.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(gvar.strip_parameter_list[13])
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
    if now_plate_line_number == 1:
        selection1.Clear()
        selection1.Search("Name=plate_line_2*,all")
        try:
            selection1.Delete()
            selection1.Clear()
        except:
            pass
    elif now_plate_line_number == 2:
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
    partDocument1.SaveAs(gvar.save_path + "up_plate_" + str(g) + ".CATPart")  # 存檔的檔案名稱
    time.sleep(2)
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    partDocument1.Close()
    # ======================================   up_plate_change


def M_upper_design(M_plate_length, M_plate_wide, ElementSketch, ElementDocument, ElementBody,
                   ElementHybridbody):
    (ElementReference1) = defs.ExtremumPoint("X_min", "Y_min", "Z_max", 2, ElementSketch, ElementDocument, ElementBody,
                                             ElementHybridbody)  # 建立極值點   (極值方向1,極值方向2,極值方向3,需要幾個方向,在哪個元素上的極值)  element_Reference(1)為OUT
    ElementReference2 = ElementReference1
    ElementReference2.Name = "up_common_seat_plate_min"  # 建立最小點
    ElementPoint3 = ElementReference2
    for up_pin in range(1, 3):
        x_value = 35 / 2
        if up_pin == 1:
            y_value = 13
        else:
            y_value = M_plate_wide - 13
        (ElementPoint5) = defs.BuildPoint(ElementReference2, ElementDocument, ElementBody, ElementHybridbody,
                                          "Body")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
        ElementPoint5.X.Value = x_value
        ElementPoint5.Y.Value = y_value
        hybridShape1 = ElementBody.HybridShapes.Item("down_die_plate_up_plane")  # 宣告平面(下)
        hybridShape2 = ElementBody.HybridShapes.Item("down_die_plate_down_plane")  # 宣告平面(上)
        ElementReference11 = ElementPoint5
        ElementReference12 = hybridShape2
        (hole_pin) = defs.HoleSimpleD(12, 40, 0, ElementDocument, ElementBody, ElementReference11,
                                      ElementReference12)  # 直孔  (M,深度,out孔陳述式,方向) element_Reference(11)=point #element_Reference(12)=plane direction=>0=下 1=上
        if up_pin == 1:
            hole_pin.Name = "pin_hole_12_R"
        else:
            hole_pin.Name = "pin_hole_12_L"
    for up_pin in range(1, 3):
        x_value = 35 / 2
        if up_pin == 1:
            y_value = 64
        else:
            y_value = M_plate_wide - 64
        (ElementPoint5) = defs.BuildPoint(ElementReference2, ElementDocument, ElementBody, ElementHybridbody,
                                          "Body")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
        ElementPoint5.X.Value = x_value
        ElementPoint5.Y.Value = y_value
        ElementReference11 = (ElementPoint5)
        ElementReference12 = hybridShape2
        (hole_pin) = defs.HoleSimpleD(16.5, 40, 0, ElementDocument, ElementBody, ElementReference11,
                                      ElementReference12)  # 直孔  (M,深度,out孔陳述式,方向) element_Reference(11)=point #element_Reference(12)=plane direction=>0=下 1=上
    x_value = 35 / 2
    y_value = M_plate_wide / 2
    (ElementPoint5) = defs.BuildPoint(ElementReference2, ElementDocument, ElementBody, ElementHybridbody,
                                      "Body")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
    ElementPoint5.X.Value = x_value
    ElementPoint5.Y.Value = y_value
    ElementReference11 = (ElementPoint5)
    ElementReference12 = hybridShape2
    (hole_pin) = defs.HoleSimpleD(8, 40, 0, ElementDocument, ElementBody, ElementReference11,
                                  ElementReference12)  # 直孔  (M,深度,out孔陳述式,方向) element_Reference(11)=point #element_Reference(12)=plane direction=>0=下 1=上
    defs.BuildSketch("E_sketch", hybridShape1, ElementDocument, "Body", ElementBody,
                     ElementHybridbody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    for up_pin in range(1, 3):
        x_value = 17.5
        if up_pin == 1:
            y_value = 15
        else:
            y_value = M_plate_wide - 15
        (ElementPoint5) = defs.BuildPoint(ElementReference2, ElementDocument, ElementBody, ElementHybridbody,
                                          "Body")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
        ElementPoint5.X.Value = x_value
        ElementPoint5.Y.Value = y_value
        (ElementPoint1,element_line1, element_line2, element_line3, element_line4) = defs.SketchRectangle(ElementSketch, 35, 30, ElementDocument,
                                               ElementHybridbody)  # (草圖陳述句,長,寬)  element_point(1)=>output中心點之陳述句  [需經環境副程式跑過]
        ElementPoint3 = ElementPoint5
        ElementPoint4 = ElementPoint1
        defs.SketchBuildCallout(ElementSketch, "free", "Binding", 0, ElementDocument, ElementPoint3,
                                ElementPoint4)  # 建立標註(草圖陳述句,標註方向,式標OR拘束,改變OR讀取之數值,標註兩點)
    # ==============挖除元素
    part1 = ElementDocument.Part
    shapeFactory1 = part1.ShapeFactory
    reference1 = part1.CreateReferenceFromObject(ElementSketch)
    part1.InWorkObject = ElementBody
    pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
    # ==============挖除元素
    pocket1.DirectionOrientation = 0
    pocket1.FirstLimit.dimension.Value = 18
    defs.BuildSketch("E_sketch", hybridShape2, ElementDocument, "Body", ElementBody,
                     ElementHybridbody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    (ElementPoint1,element_line1, element_line2, element_line3, element_line4) = defs.SketchRectangle(ElementSketch, 35, 20, ElementDocument,
                                           ElementHybridbody)  # (草圖陳述句,長,寬)  element_point(1)=>output中心點之陳述句  [需經環境副程式跑過]
    (ElementPoint5) = defs.BuildPoint(ElementReference2, ElementDocument, ElementBody, ElementHybridbody,
                                      "Body")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
    ElementPoint5.X.Value = M_plate_length / 2
    ElementPoint5.Y.Value = M_plate_wide / 2
    ElementPoint3 = ElementPoint5
    ElementPoint4 = ElementPoint1
    defs.SketchBuildCallout(ElementSketch, "Horizontal", "Binding", 0, ElementDocument, ElementPoint3,
                            ElementPoint4)  # 建立標註(草圖陳述句,標註方向,式標OR拘束,改變OR讀取之數值,標註兩點)
    defs.SketchBuildCallout(ElementSketch, "Vertical", "Binding", 0, ElementDocument, ElementPoint3,
                            ElementPoint4)  # 建立標註(草圖陳述句,標註方向,式標OR拘束,改變OR讀取之數值,標註兩點)
    # ==============挖除元素
    part1 = ElementDocument.Part
    shapeFactory1 = part1.ShapeFactory
    reference1 = part1.CreateReferenceFromObject(ElementSketch)
    part1.InWorkObject = ElementBody
    pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
    # ==============挖除元素
    pocket1.DirectionOrientation = 1
    pocket1.FirstLimit.dimension.Value = 10
    defs.BuildSketch("E_sketch", hybridShape1, ElementDocument, "Body", ElementBody,
                     ElementHybridbody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    (ElementPoint1,element_line1, element_line2, element_line3, element_line4) = defs.SketchRectangle(ElementSketch, 35, 20, ElementDocument,
                                           ElementHybridbody)  # (草圖陳述句,長,寬)  element_point(1)=>output中心點之陳述句  [需經環境副程式跑過]
    (ElementPoint5) = defs.BuildPoint(ElementReference2, ElementDocument, ElementBody, ElementHybridbody,
                                      "Body")  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
    ElementPoint5.X.Value = M_plate_length / 2
    ElementPoint5.Y.Value = M_plate_wide / 2
    ElementPoint3 = ElementPoint5
    ElementPoint4 = ElementPoint1
    defs.SketchBuildCallout(ElementSketch, "Horizontal", "Binding", 0, ElementDocument, ElementPoint3,
                            ElementPoint4)  # 建立標註(草圖陳述句,標註方向,式標OR拘束,改變OR讀取之數值,標註兩點)
    defs.SketchBuildCallout(ElementSketch, "Vertical", "Binding", 0, ElementDocument, ElementPoint3,
                            ElementPoint4)  # 建立標註(草圖陳述句,標註方向,式標OR拘束,改變OR讀取之數值,標註兩點)
    # ==============挖除元素
    part1 = ElementDocument.Part
    shapeFactory1 = part1.ShapeFactory
    reference1 = part1.CreateReferenceFromObject(ElementSketch)
    part1.InWorkObject = ElementBody
    pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
    # ==============挖除元素
    pocket1.DirectionOrientation = 0
    pocket1.FirstLimit.dimension.Value = 9


def punch_Reinforcement_Hole(g, n, i):
    length = [None] * 11
    body_name = [""] * 3
    sketch_name = [""] * 3
    point_name = [""] * 3
    formula_name1 = "cut_line_formula_3"
    line_name1 = "plate_line_" + str(g) + "_op" + str(n * 10) + "_Reinforcement_cut_line_" + str(i)
    body_name[1] = "Body.2"
    body_name[2] = "Reinforcement_cut_punch"
    sketch_name[1] = "Bolt_Sketch"
    element_name1 = "down_die_plate_up_plane"
    point_name[2] = " "
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    PunchDef.formula_change(body_name, line_name1, sketch_name,formula_name1)
    # ---------改變參數↓-------------
    strParam1 = parameters1.Item("QR_type")
    strParam1.Value = "QR_B"
    # ---------改變參數↑-------------
    length[1] = part1.Parameters.Item("QR_wide")
    length[2] = part1.Parameters.Item("Reinforcement_Grow_wide_2")
    if length[1].Value < 10:
        length[2].Value = 10 - length[1].Value
        strParam1.Value = "Reinforcement_C"
        sketch_name[2] = "C_pad_Excavation_Sketch"
    else:
        strParam1.Value = "QR_B"
        sketch_name[2] = "B_pad_Excavation_Sketch"
    length[0] = part1.Parameters.Item("Reinforcement_Grow_wide_1")
    length[0].Value = 20
    length[0] = part1.Parameters.Item("Reinforcement_Grow_long_1")
    length[0].Value = 46
    length[0] = part1.Parameters.Item("Reinforcement_Grow_long_2")
    length[0].Value = 25
    length[0] = part1.Parameters.Item("QR_punch_up_plane")
    if gvar.Mold_status == "開模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14]) + 28  # (upper_die_open_height)
    elif gvar.Mold_status == "閉模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14])
    length[0] = part1.Parameters.Item("Bolt_long")
    if strParam1.Value == "QR_B":
        length[0].Value = 9
    else:
        length[0].Value = 9 + length[2].Value
    part1.Update()
    point_name[1] = "Bolt_Point_B_1"
    up_pad_Bolt_Hole = int()
    (point_name[2], up_pad_Bolt_Hole) = point_build(g, point_name[1], 12, body_name, up_pad_Bolt_Hole)
    (point_name[2], up_pad_Bolt_Hole) = point_build(g, point_name[1], 20, body_name, up_pad_Bolt_Hole)
    F_Hole_up(n, i, 1, 16, element_name1, point_name[2], body_name)
    point_name[1] = "Bolt_Point_B_2"
    (point_name[2], up_pad_Bolt_Hole) = point_build(g, point_name[1], 12, body_name, up_pad_Bolt_Hole)
    (point_name[2], up_pad_Bolt_Hole) = point_build(g, point_name[1], 20, body_name, up_pad_Bolt_Hole)
    F_Hole_up(n, i, 2, 16, element_name1, point_name[2], body_name)
    return up_pad_Bolt_Hole


def point_build(g, point, high, body_name, up_pad_Bolt_Hole_Input):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    parameters1 = part1.Parameters
    bodies1 = part1.Bodies
    body1 = bodies1.Item(body_name[1])
    body2 = bodies1.Item(body_name[2])
    # ------------------------------------------↓搜尋現有pad_Bolt_Hole點的螺栓數量
    selection1.Clear()
    selection1.Search("Name=up_pad_" + str(g) + "_Bolt_point_*,all")
    point_conter = selection1.Count
    selection1.Clear()
    # ------------------------------------------↑搜尋現有pad_Bolt_Hole點的螺栓數量
    # ------------------------------------------------------------↓   零建檔建立點之語法"HybridShapes"包涵在body裡
    hybridShapes2 = body2.HybridShapes
    hybridShapePointCoord1 = hybridShapes2.Item(point)
    reference2 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
    # ------------------------------------------------------------↑
    # ------------------------------------------------------------↓   建立點
    hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, high)  # (x,y,z)
    hybridShapePointCoord1.PtRef = reference2  # 基礎點 = reference2
    body1.InsertHybridShape(hybridShapePointCoord1)  # 建立點
    part1.Update()
    # ------------------------------------------------------------↑
    reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
    point_conter = point_conter + 1
    up_pad_Bolt_Hole = up_pad_Bolt_Hole_Input + point_conter
    # ------------------------------------------------------------↓   建立點(打斷關連的)
    hybridShapePointExplicit1 = hybridShapeFactory1.AddNewPointDatum(reference3)  # 將點置於元素3
    body1.InsertHybridShape(hybridShapePointExplicit1)
    part1.InWorkObject = hybridShapePointExplicit1
    hybridShapePointExplicit1.Name = "up_pad_" + str(g) + "_Bolt_point_" + point_conter  # 點改名字
    point_name = "up_pad_" + str(g) + "_Bolt_point_" + point_conter  # 點名字輸出
    part1.Update()
    hybridShapeFactory1.DeleteObjectForDatum(reference3)  # 刪除元素3
    selection1.Add(hybridShapePointExplicit1)
    visPropertySet1 = visPropertySet1.Parent
    visPropertySet1.SetShow(1)
    selection1.Clear()
    # ------------------------------------------------------------↑
    return point_name, up_pad_Bolt_Hole


def F_Hole_up(n, i, hole_digital, hole_high, plane, point_name, body_name):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    shapeFactory1 = part1.ShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item(body_name[1])
    hybridShapeFactory1 = part1.HybridShapeFactory
    parameters1 = part1.Parameters
    # ------------------------------------------------------------↓   在已有平面建立條件             (方法二)
    hybridShapes1 = body1.HybridShapes
    hybridShapePlaneOffset1 = hybridShapes1.Item(plane)
    reference1 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
    # ------------------------------------------------------------↑
    hybridShapePointExplicit1 = parameters1.Item(point_name)
    reference2 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
    hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference2, reference1, 15)
    hole1.Name = ("Hole_OP" + (n * 10) + "_pick_" + (i) + "_" + (hole_digital))  # 孔改名字
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
    # =============================================↓   牙深設定 孔深>牙深(否則錯誤)
    length2 = limit2.dimension  # 孔深
    length2.Value = hole_high
    # =============================================↓   牙深設定 孔深>牙深(否則錯誤)
    length3 = hole1.ThreadDepth  # 牙深
    length3.Value = length2.Value - 2
    # =============================================↑
    sketch1 = hole1.sketch
    selection1.Add(sketch1)
    visPropertySet1 = visPropertySet1.Parent
    visPropertySet1.SetShow(1)
    selection1.Clear()
    hole1.Reverse()
    part1.UpdateObject(hole1)


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
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapes1 = body1.HybridShapes
        hybridShapePointCoord1 = hybridShapes1.Item("demise_hole_left_offset")
        reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
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
        part1.InWorkObject.Name = ("hang_bolt_end_point_left_project_" + str(i))
        hybridShapeFactory1.DeleteObjectForDatum(reference3)
        part1.Update()
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=hang_bolt_end_point_left_project_*,all")
        selection1.VisProperties.SetShow(1)  # 1隱藏/0顯示
        selection1.Clear()
        # ------------------------------------------------------------↑offset
        partDocument1 = catapp.ActiveDocument
        selection1 = partDocument1.Selection
        selection1.Clear()
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("hang_bolt_end_point_left_project_" + str(i))
        selection1.Add(hybridShapeCurveExplicit1)
        selection1.Copy()
        # ===================================================================
        windows1 = catapp.Windows
        specsAndGeomWindow1 = windows1.Item("Data1.CATPart")
        specsAndGeomWindow1.Activate()
        # ===================================================================
        partDocument2 = catapp.ActiveDocument
        part2 = partDocument2.Part
        bodies2 = part2.Bodies
        body2 = bodies2.Add()
        body2.Name = "remove_body_" + str(i)
        part2.Update()
        selection2 = partDocument2.Selection
        selection2.Clear()
        selection2.Add(body2)
        selection2.Paste()
        partDocument1.Close()
        # ---------------------------------------------
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        hybridShapeFactory1 = part1.HybridShapeFactory
        parameters1 = part1.Parameters
        hybridShapePointExplicit2 = parameters1.Item("hang_bolt_end_point_left_project_" + str(i))
        reference1 = part1.CreateReferenceFromObject(hybridShapePointExplicit2)
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        hybridShapes1 = body1.HybridShapes
        hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
        reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
        shapeFactory1 = part2.ShapeFactory
        hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 10)
        hole1.Type = 0  # (catSimpleHole)
        hole1.AnchorMode = 0  # (catExtremPointHoleAnchor)
        limit1 = hole1.BottomLimit
        limit1.LimitMode = 0  # (catOffsetLimit)
        hole1.ThreadingMode = 1  # (catSmoothHoleThreading)
        hole1.ThreadSide = 0  # (catRightThreadSide)
        hole1.BottomType = 2  # (catTrimmedHoleBottom)
        limit1.LimitMode = 3  # (catUpToPlaneLimit)
        hybridShapePlaneOffset3 = hybridShapes1.Item("down_die_plate_down_plane")
        reference6 = part1.CreateReferenceFromObject(hybridShapePlaneOffset3)
        limit1.LimitingElement = reference6
        hole1.ThreadingMode = 0
        hole1.CreateStandardThreadDesignTable(1)  # (catHoleMetricThickPitch)
        strParam1 = hole1.HoleThreadDescription
        strParam1.Value = "M8"
        hole1.Reverse()
        part2.Update()
        part2.InWorkObject = body1
        remove1 = shapeFactory1.AddNewRemove(body2)
        part2.Update()
    selection1 = partDocument2.Selection
    selection1.Clear()
    selection1.Search("Name=Sketch*,all")
    selection1.VisProperties.SetShow(1)  # '1隱藏/0顯示
    selection1.Clear()


def emboss_forming_punch_right(g, n):  # 打凸包沖頭_右
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    now_plate_line_number = g
    op_number = n * 10
    for i in range(1, round(gvar.StripDataList[22][g][n]) + 1):
        partDocument1 = documents1.Open(
            gvar.save_path + "op" + str(op_number) + "_emboss_forming_punch_right_0" + str(i) + ".CATPart")  # 開啟挖孔input
        part1 = partDocument1.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapes1 = body1.HybridShapes
        hybridShapePointCoord1 = hybridShapes1.Item("hang_bolt_end_point_right")
        reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
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
        hybridShapePointExplicit1 = hybridShapeFactory1.AddNewPointDatum(reference3)
        body1.InsertHybridShape(hybridShapePointExplicit1)
        part1.InWorkObject = hybridShapePointExplicit1
        part1.InWorkObject.Name = ("hang_bolt_end_point_right_project_" + str(i))
        hybridShapeFactory1.DeleteObjectforDatum(reference3)
        part1.Update()
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=hang_bolt_end_point_right_project_*,all")
        selection1.VisProperties.SetShow(1)  # 1隱藏/0顯示
        selection1.Clear()
        # --------------------------------------------------------------------------------------------↓複製
        partDocument1 = catapp.ActiveDocument
        selection1 = partDocument1.Selection
        selection1.Clear()
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        hybridShapePointExplicit1 = parameters1.Item("hang_bolt_end_point_right_project_" + str(i))
        selection1.Add(hybridShapePointExplicit1)
        selection1.Copy()
        # --------------------------------------------------------------------------------------------↑複製
        # ---------------------------------------------↓切換視窗
        windows1 = catapp.Windows
        specsAndGeomWindow1 = windows1.Item("Data1.CATPart")
        specsAndGeomWindow1.Activate()
        # ---------------------------------------------↑切換視窗
        # ---------------------------------------------↓搜尋remove_body數
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=remove_body_*,all")
        nn = selection1.Count
        selection1.Clear()
        # ---------------------------------------------↑搜尋remove_body數
        partDocument2 = catapp.ActiveDocument
        part2 = partDocument2.Part
        bodies2 = part2.Bodies
        body2 = bodies2.Add()
        body2.Name = "remove_body_" + str(i + nn)
        part2.Update()
        # ---------------------------------------------↓關閉沖頭視窗
        partDocument2 = catapp.ActiveDocument
        selection2 = partDocument2.Selection
        selection2.Clear()
        selection2.Add(body2)
        selection2.Paste()
        partDocument1.Close()
        # ---------------------------------------------↑關閉沖頭視窗
        shapeFactory1 = part2.ShapeFactory
        parameters1 = part2.Parameters
        hybridShapePointExplicit2 = parameters1.Item("hang_bolt_end_point_right_project_" + str(i))
        reference4 = part2.CreateReferenceFromObject(hybridShapePointExplicit2)
        body3 = bodies2.Item("Body.2")
        hybridShapes2 = body3.HybridShapes
        hybridShapePlaneOffset2 = hybridShapes2.Item("down_die_plate_up_plane")
        reference5 = part2.CreateReferenceFromObject(hybridShapePlaneOffset2)
        hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference4, reference5, 10)
        hole1.Type = 0  # (catSimpleHole)
        hole1.AnchorMode = 0  # (catExtremPointHoleAnchor)
        limit1 = hole1.BottomLimit
        limit1.LimitMode = 0  # (catOffsetLimit)
        hole1.ThreadingMode = 1  # (catSmoothHoleThreading)
        hole1.ThreadSide = 0  # (catRightThreadSide)
        hole1.BottomType = 2  # (catTrimmedHoleBottom)
        limit1.LimitMode = 3  # (catUpToPlaneLimit)
        hybridShapePlaneOffset3 = hybridShapes2.Item("down_die_plate_down_plane")
        reference6 = part2.CreateReferenceFromObject(hybridShapePlaneOffset3)
        limit1.LimitingElement = reference6
        hole1.ThreadingMode = 0
        hole1.CreateStandardThreadDesignTable(1)  # (catHoleMetricThickPitch)
        strParam1 = hole1.HoleThreadDescription
        strParam1.Value = "M8"
        hole1.Reverse()
        part2.Update()
        part2.InWorkObject = body3
        remove1 = shapeFactory1.AddNewRemove(body2)
        part2.Update()
    selection1 = partDocument2.Selection
    selection1.Clear()
    selection1.Search("Name=Sketch*,all")
    selection1.VisProperties.SetShow(1)  # 1隱藏/0顯示
    selection1.Clear()
    part1.Update()


def QR_half_cut_punch(g, n):  # 半衝切
    catapp = win32.Dispatch('CATIA.Application')
    partDocument2 = catapp.ActiveDocument
    product1 = partDocument2.getItem("Part1")
    part1 = partDocument2.Part
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.open_path + "QR_half_cut_punch_line.CATPart")
    # ======================================
    defs.window_change(partDocument2, partDocument1)  # 在CATIA上切換各視窗
    length = [None] * 99
    # ======================================================================================================
    length[0] = part1.Parameters.Item("QR_half_punch_up_plane")
    now_plate_line_number = g
    if gvar.Mold_status == "開模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14]) + 28  # (upper_die_open_height)
    elif gvar.Mold_status == "閉模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14])
    # ======================================================================================================
    length[1] = part1.Parameters.Item("QR_half_punch_height")
    op_number = n * 10
    for i in range(1, round(gvar.StripDataList[3][g][n]) + 1):
        try:
            length[1].Value = QR_punch_height
        except:
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
            elif Thickness >= 3.5:
                excel_Sheet_name = "3.5以上"
            (serch_result) = defs.ExcelSearch(die_rule_file_name, excel_Sheet_name, Row_string_serch,
                                              Column_string_serch)
            length[1].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
                gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14]) + serch_result  # 沖頭高度
        length[2] = part1.Parameters.Item("QR_half_punch_Hanging_Desk_height")
        length[2].Value = float(gvar.strip_parameter_list[14])
        part1.Parameters.Item("half_cut_cut_line_formula_" + str(i)).OptionalRelation.Modify(
            "die\plate_line_" + str(g) + "_op" + str(op_number) + "_half_cut_line_" + str(i))  # 草圖置換
        part1.Update()
        for B_n in range(60, 0, -1):
            selection1 = partDocument1.Selection
            selection1.Clear()
            selection1.Search("Name=Body." + str(B_n) + "*,all")
            Body_n = selection1.Count
            selection1.Clear()
            if Body_n > 0:
                body_number = B_n
                break
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, 0)
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body." + str(body_number))
        hybridShapes1 = body1.HybridShapes
        hybridShapePointCoord2 = hybridShapes1.Item("Bolt_point_" + str(i))
        reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord2)
        hybridShapePointCoord1.PtRef = reference1
        body1 = bodies1.Item("Body.2")
        hybridShapePointExplicit1 = hybridShapeFactory1.AddNewPointDatum(reference1)
        body1.InsertHybridShape(hybridShapePointExplicit1)
        part1.InWorkObject = hybridShapePointExplicit1
        reference1 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
        hybridShapes1 = body1.HybridShapes
        hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
        reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
        shapeFactory1 = part1.ShapeFactory
        hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 40)
        hole1.Type = 0  # (catSimpleHole)
        hole1.AnchorMode = 0  # (catExtremPointHoleAnchor)
        limit1 = hole1.BottomLimit
        limit1.LimitMode = 0  # (catOffsetLimit)
        hole1.ThreadingMode = 1  # (catSmoothHoleThreading)
        hole1.ThreadSide = 0  # (catRightThreadSide)
        hole1.BottomType = 1  # (catVHoleBottom)
        length2 = limit1.dimension
        length2.Value = 24
        hole1.ThreadingMode = 0
        hole1.CreateStandardThreadDesignTable(1)  # (catHoleMetricThickPitch)
        strParam1 = hole1.HoleThreadDescription
        strParam1.Value = "M8"
        length3 = hole1.ThreadDepth
        length3.Value = 16
        hole1.Reverse()
        part1.UpdateObject(hole1)
    selection1 = partDocument1.Selection
    selection1.Clear()
    selection1.Search("Name=Sketch.*,all")
    selection1.VisProperties.Show(1)  # 1為隱藏,0為顯示
    selection1.Clear()
    selection1.Search("Name=Point.*,all")
    selection1.VisProperties.Show(1)  # 1為隱藏,0為顯示
    selection1.Clear()


def bend_up_shaping_cavity_hole_1(op_number, pp_count):  # 整平模組孔down
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(
        gvar.save_path + "op" + str(op_number) + "_bend_up_shaping_punch_down_" + str(pp_count) + ".CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length1 = parameters1.Item("offset_sketch_length")
    length1.Value = 0
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
    limit1.LimitMode = 3  # (catUpToPlaneLimit)
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


def bend_up_shaping_cavity_hole_2(op_number, pp_count):  # 整平模組孔down
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(
        gvar.save_path + "op" + str(op_number) + "_bend_up_shaping_punch_up_" + str(pp_count) + ".CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length1 = parameters1.Item("offset_sketch_length")
    length1.Value = 0
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
    limit1.LimitMode = 3  # (catUpToPlaneLimit)
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


def F_bending_up_plate(g, op_number, now_data_number):
    now_plate_line_number = g
    OP = op_number
    i = now_data_number
    EX_file_name = "op" + str(OP) + "_bending_punch_" + str(i)
    plate_name = "Data1"
    deep = 20
    M = "M10"
    direction = 1
    plane_name = "down_die_plate_up_plane"
    bolt_point_name = "Bolt_Start_point"
    body_name2 = "bending_punch"
    body_name3 = "Body.2"
    parameter_name = "bending_cavity_parameter"
    # ===========================FUN_bolt_hole
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.save_path + EX_file_name + ".CATPart")
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(EX_file_name, body_name2,
                                                                                             "die")  # 環境設定
    part1 = ElementDocument.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    hybridShape1 = ElementBody.HybridShapes.Item(bolt_point_name)  # 宣告Bolt_point
    reference2 = part1.CreateReferenceFromObject(hybridShape1)
    hybridShapePointExplicit1 = hybridShapeFactory1.AddNewPointDatum(reference2)
    ElementBody.InsertHybridShape(hybridShapePointExplicit1)
    part1.InWorkObject = hybridShapePointExplicit1
    part1.Update()
    hybridShapePointExplicit1.Name = (bolt_point_name + "_" + str(op_number) + "_" + str(now_data_number))
    selection1 = ElementDocument.Selection
    selection1.Clear()
    selection1.Add(hybridShapePointExplicit1)
    selection1.Copy()
    if plate_name == "Data1":
        partDocument2 = documents1.Item(plate_name + ".CATPart")
    else:
        partDocument2 = documents1.Open(gvar.save_path + plate_name + ".CATPart")
    part1 = ElementDocument.Part
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(plate_name, body_name3,
                                                                                             "die")  # 環境設定
    selection1 = ElementDocument.Selection
    selection1.Clear()
    selection1.Add(ElementBody)
    selection1.Copy()
    partDocument1.Close()
    ElementReference11 = ElementBody.HybridShapes.Item(
        bolt_point_name + "_" + str(op_number) + "_" + str(now_data_number))  # 宣告Bolt_point
    ElementReference12 = ElementBody.HybridShapes.Item(plane_name)  # 宣告Bolt_point
    visPropertySet1 = selection1.VisProperties
    selection1.Clear()
    selection1.Add(ElementReference11)
    visPropertySet1 = visPropertySet1.Parent
    visPropertySet1.SetShow(1)
    selection1.Clear()
    (hole) = defs.hole_simple_M(M, deep, direction, ElementDocument, ElementBody, ElementReference11,
                                ElementReference12)  # 直孔  (M,深度,out孔陳述式,方向) element_Reference(11)=point #element_Reference(12)=plane direction=>0=下 1=上
    if plate_name != "Data1":
        partDocument2.save()
        partDocument2.Close()
    # ================FUN_bolt_hole(EX_file_name, plate_name, deep, bolt_point_name, plane_name, M, direction)

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
# UpPlate(1)

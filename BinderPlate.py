import win32com.client as win32
import global_var as gvar
import time


def BinderPlateSystem(now_plate_line_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    op_number = 0
    total_op_number = int(gvar.strip_parameter_list[2])
    for n in range(1, 1 + total_op_number):
        op_number = 10 * n
        if op_number == 30 or op_number == 50:
            Binder_Plate_other(op_number)
        elif op_number == 40:
            Binder_Plate(op_number)
    # --------------↓刪除不需要的Data↓--------------
    partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
    selection1 = partDocument1.Selection
    selection1.Clear()
    part1 = partDocument1.Part
    hybridBodies1 = part1.HybridBodies
    hybridBody1 = hybridBodies1.Item("die")
    hybridShapes1 = hybridBody1.HybridShapes
    try:
        hybridShapeExtremum1 = hybridShapes1.Item("Extremum.5")
        selection1.Add(hybridShapeExtremum1)
        selection1.Delete()
    except:
        pass
    partDocument1.save()

    # --------------↑刪除不需要的Data↑--------------


def Binder_Plate_other(op_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
    # --------------↓刪除不需要的Data↓--------------
    if op_number == 50:
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=Data_*,all")
        selection1.Delete()
    # --------------↑刪除不需要的Data↑--------------
    part1 = partDocument1.Part
    relations1 = part1.Relations
    parameters1 = part1.Parameters  # 參數指令起手宣告
    # -----------------------------------------------------------------------------建立測量參數
    length1 = parameters1.CreateDimension("", "LENGTH", 0)  # build parameter
    length2 = parameters1.CreateDimension("", "LENGTH", 0)
    length3 = parameters1.CreateDimension("", "LENGTH", 0)
    length4 = parameters1.CreateDimension("", "LENGTH", 0)
    length5 = parameters1.CreateDimension("", "LENGTH", 0)
    length6 = parameters1.CreateDimension("", "LENGTH", 0)
    length1.rename("Data_Binder_Plate_other_outside")
    length2.rename("Data_Binder_Plate_other_inside")
    length3.rename("Data_Binder_Plate_other_thickness")
    length4.rename("Data_Binder_Plate_other_distance")
    length5.rename("Data_Binder_Plate_other_distance_H")
    length6.rename("Data_Binder_Plate_other_height")
    # -----------------------------------------------------------------------------開始測量
    formula1 = relations1.Createformula("Data_Binder_Plate_other_outside", "", length1, "length()")
    formula2 = relations1.Createformula("Data_Binder_Plate_other_inside", "", length2, "length()")
    formula3 = relations1.Createformula("Data_Binder_Plate_other_thickness", "", length3, "length()")
    formula4 = relations1.Createformula("Data_Binder_Plate_other_distance", "", length4, "length()")
    formula5 = relations1.Createformula("Data_Binder_Plate_other_distance_H", "", length5, "length()")
    formula6 = relations1.Createformula("Data_Binder_Plate_other_height", "", length6, "length()")
    # -----------------------------------------------------------------------------鍵槽
    if op_number == 30:
        S_point_distance_parameter = "die\\open_curve_1_1_A"
        E_point_distance_parameter = "die\\Extremum.5"
        formula1.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject(formula1)  # 單步更新 formula
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\plate_centor_point"
        formula4.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject(formula4)  # 單步更新 formula
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\Extremum.5"
        formula5.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject(formula5)  # 單步更新 formula
        # -----------------------------------------------------------------------------中心軸
    elif op_number == 50:
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\open_curve_1_1_A"
        formula1.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject(formula1)  # 單步更新 formula
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\plate_centor_point"
        formula4.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject(formula4)  # 單步更新 formula
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\open_curve_1_1_A"
        formula5.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject(formula5)  # 單步更新 formula
    # -----------------------------------------------------------------------------開始測量
    ##--------------------------------------------------------------讀取數據
    Data_Binder_Plate_other_outside = part1.Parameters.Item("Data_Binder_Plate_other_outside")
    Data_Binder_Plate_other_outside = Data_Binder_Plate_other_outside.Value
    Data_Binder_Plate_other_inside = part1.Parameters.Item("Data_Binder_Plate_other_inside")
    Data_Binder_Plate_other_inside = Data_Binder_Plate_other_inside.Value
    Data_Binder_Plate_other_thickness = part1.Parameters.Item("Data_Binder_Plate_other_thickness")
    Data_Binder_Plate_other_thickness = Data_Binder_Plate_other_thickness.Value
    Data_Binder_Plate_other_distance = part1.Parameters.Item("Data_Binder_Plate_other_distance")
    Data_Binder_Plate_other_distance = Data_Binder_Plate_other_distance.Value
    Data_Binder_Plate_other_distance_H = part1.Parameters.Item("Data_Binder_Plate_other_distance_H")
    Data_Binder_Plate_other_distance_H = Data_Binder_Plate_other_distance_H.Value
    Data_Binder_Plate_other_height = part1.Parameters.Item("Data_Binder_Plate_other_height")
    Data_Binder_Plate_other_height = Data_Binder_Plate_other_height.Value
    # 壓板標準件尺寸判斷↓
    if op_number == 30 or op_number == 50:
        if op_number == 30:
            Data_Binder_Plate_other_outside = int(Data_Binder_Plate_other_outside) / 2
        else:
            Data_Binder_Plate_other_outside = int(Data_Binder_Plate_other_outside)
        if Data_Binder_Plate_other_outside < 5:
            # ------------------------------------------------------------------------------------------------外徑-內徑之差值
            GAP_BP = (Data_Binder_Plate_other_outside + 1) - 2.1 - 0.5
            # ------------------------------------------------------------------------------------------------外徑-內徑之差值
            Data_Binder_Plate_other_inside = Data_Binder_Plate_other_outside - GAP_BP
            Data_Binder_Plate_other_thickness = 3
        elif Data_Binder_Plate_other_outside >= 5 and Data_Binder_Plate_other_outside < 6:
            Data_Binder_Plate_other_outside = 5
            Data_Binder_Plate_other_inside = 2.1
            Data_Binder_Plate_other_thickness = 3
        elif Data_Binder_Plate_other_outside >= 6 and Data_Binder_Plate_other_outside < 7:
            Data_Binder_Plate_other_outside = 6
            Data_Binder_Plate_other_inside = 2.6
            Data_Binder_Plate_other_thickness = 3
        elif Data_Binder_Plate_other_outside >= 7:
            Data_Binder_Plate_other_outside = 7
            Data_Binder_Plate_other_inside = 3.1
            Data_Binder_Plate_other_thickness = 3
    # 壓板標準件尺寸判斷↑
    partDocument1.save()
    partDocument1.Close()  # 關閉檔案
    # ==================================================================================================製作挖槽壓板
    time.sleep(2)
    partDocument1 = documents1.Open(gvar.open_path + "Binder_Plate_other.CATPart")
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    # ------------------------------------------------設定條件
    if op_number == 30:
        # -----------------------------------------------------------------------------鍵槽
        Binder_Plate_other_outside = part1.Parameters.Item("Binder_Plate_other_outside")
        Binder_Plate_other_outside.Value = Data_Binder_Plate_other_outside
        Binder_Plate_other_inside = part1.Parameters.Item("Binder_Plate_other_inside")
        Binder_Plate_other_inside.Value = Data_Binder_Plate_other_inside
        Binder_Plate_other_thickness = part1.Parameters.Item("Binder_Plate_other_thickness")
        Binder_Plate_other_thickness.Value = Data_Binder_Plate_other_thickness
        Binder_Plate_other_distance = part1.Parameters.Item("Binder_Plate_other_distance")
        Binder_Plate_other_distance.Value = Data_Binder_Plate_other_distance - float(gvar.strip_parameter_list[4])
        Binder_Plate_other_distance_H = part1.Parameters.Item("Binder_Plate_other_distance_H")
        Binder_Plate_other_distance_H.Value = Data_Binder_Plate_other_distance_H - Data_Binder_Plate_other_outside + 1
        # 壓板高度
        Binder_Plate_other_height = part1.Parameters.Item("Binder_Plate_other_height")
        Binder_Plate_other_height.Value = float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[1])
    # -----------------------------------------------------------------------------中心軸
    elif op_number == 50:
        Binder_Plate_other_outside = part1.Parameters.Item("Binder_Plate_other_outside")
        Binder_Plate_other_outside.Value = Data_Binder_Plate_other_outside
        Binder_Plate_other_inside = part1.Parameters.Item("Binder_Plate_other_inside")
        Binder_Plate_other_inside.Value = Data_Binder_Plate_other_inside
        Binder_Plate_other_thickness = part1.Parameters.Item("Binder_Plate_other_thickness")
        Binder_Plate_other_thickness.Value = Data_Binder_Plate_other_thickness
        Binder_Plate_other_distance = part1.Parameters.Item("Binder_Plate_other_distance")
        Binder_Plate_other_distance.Value = Data_Binder_Plate_other_distance + float(gvar.strip_parameter_list[4])
        Binder_Plate_other_distance_H = part1.Parameters.Item("Binder_Plate_other_distance_H")
        Binder_Plate_other_distance_H.Value = Data_Binder_Plate_other_distance_H + Data_Binder_Plate_other_outside - 1
        Binder_Plate_other_height = part1.Parameters.Item("Binder_Plate_other_height")
        Binder_Plate_other_height.Value = float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[1])
    part1.Update()
    product1.PartNumber = "Binder_Plate_" + str(op_number) + ""
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Binder_Plate_" + str(op_number) + "")
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
    strParam4.ValuateFromString(gvar.strip_parameter_list[33])
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam5 = parameters5.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString('')
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
    # ---------------------------------------------------存檔
    partDocument1.SaveAs(gvar.save_path + "Binder_Plate_" + str(op_number) + ".CATPart")  # 存檔的檔案名稱
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    time.sleep(2)
    partDocument1.Close()  # 關閉檔案


def Binder_Plate(op_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
    # ---------------------------------------------------------------------------開起Data
    # --------------↓刪除不需要的Data↓--------------
    selection1 = partDocument1.Selection
    selection1.Clear()
    selection1.Search("Name=Data_Binder_Plate_*,all")
    if selection1.Count > 1 :
        selection1.Delete()
    # --------------↑刪除不需要的Data↑--------------
    part1 = partDocument1.Part
    relations1 = part1.Relations
    parameters1 = part1.Parameters  # 參數指令起手宣告
    # -----------------------------------------------------------------------------建立測量參數
    length1 = parameters1.CreateDimension("", "LENGTH", 0)  # build parameter
    length2 = parameters1.CreateDimension("", "LENGTH", 0)
    length3 = parameters1.CreateDimension("", "LENGTH", 0)
    length4 = parameters1.CreateDimension("", "LENGTH", 0)
    length5 = parameters1.CreateDimension("", "LENGTH", 0)
    length6 = parameters1.CreateDimension("", "LENGTH", 0)
    length7 = parameters1.CreateDimension("", "LENGTH", 0)
    length8 = parameters1.CreateDimension("", "LENGTH", 0)
    length1.rename("Data_Binder_Plate_outside")
    length2.rename("Data_Binder_Plate_length")
    length3.rename("Data_Binder_Plate_wide")
    length4.rename("Data_Binder_Plate_thickness")
    length5.rename("Data_Binder_Plate_number")
    length6.rename("Data_Binder_Plate_distance")
    length7.rename("Data_Binder_Plate_height")
    length8.rename("Data_Binder_Plate_Bolt_hole")
    # -----------------------------------------------------------------------------建立測量參數
    # -----------------------------------------------------------------------------開始測量
    formula1 = relations1.Createformula("Data_Binder_Plate_outside", "", length1, "length()")  # 中心外徑
    formula2 = relations1.Createformula("Data_Binder_Plate_length", "", length2, "length()")  # 下料距離
    formula3 = relations1.Createformula("Data_Binder_Plate_wide", "", length3, "length()")  # 靴齒部間隙
    formula4 = relations1.Createformula("Data_Binder_Plate_thickness", "", length4, "length()")  # 壓板厚度
    formula5 = relations1.Createformula("Data_Binder_Plate_number", "", length5, "length()")  # 壓板數量
    formula6 = relations1.Createformula("Data_Binder_Plate_distance", "", length6, "length()")  # 工站間距
    formula7 = relations1.Createformula("Data_Binder_Plate_height", "", length7, "length()")  # 壓板高度
    formula8 = relations1.Createformula("Data_Binder_Plate_Bolt_hole", "", length8, "length()")  # 壓板螺栓孔
    # 中心外徑
    S_point_distance_parameter = "die\\zero_original_point"
    E_point_distance_parameter = "die\\open_curve_center_point_1"
    formula1.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
    part1.UpdateObject(formula1)  # 單步更新 formula
    # 下料距離
    S_point_distance_parameter = "die\\zero_original_point"
    E_point_distance_parameter = "die\\Contour_circle_line_Ymax"
    formula2.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
    part1.UpdateObject(formula2)  # 單步更新 formula
    # 靴齒部間隙
    S_point_distance_parameter = "die\\open_curve_2_1"
    E_point_distance_parameter = "die\\open_curve_2_2"
    formula3.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
    part1.UpdateObject(formula3)  # 單步更新 formula
    # 工站間距
    S_point_distance_parameter = "die\\zero_original_point"
    E_point_distance_parameter = "die\\plate_centor_point"
    formula6.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
    part1.UpdateObject(formula6)  # 單步更新 formula
    # -----------------------------------------------------------------------------開始測量
    # --------------------------------------------------------------讀取數據
    Data_Binder_Plate_outside = part1.Parameters.Item("Data_Binder_Plate_outside")
    Data_Binder_Plate_outside = Data_Binder_Plate_outside.Value
    Data_Binder_Plate_length = part1.Parameters.Item("Data_Binder_Plate_length")
    Data_Binder_Plate_length = Data_Binder_Plate_length.Value
    Data_Binder_Plate_wide = part1.Parameters.Item("Data_Binder_Plate_wide")
    Data_Binder_Plate_wide = Data_Binder_Plate_wide.Value
    Data_Binder_Plate_thickness = part1.Parameters.Item("Data_Binder_Plate_thickness")
    Data_Binder_Plate_thickness = Data_Binder_Plate_thickness.Value
    Data_Binder_Plate_number = part1.Parameters.Item("Data_Binder_Plate_number")
    Data_Binder_Plate_number = Data_Binder_Plate_number.Value
    Data_Binder_Plate_distance = part1.Parameters.Item("Data_Binder_Plate_distance")
    Data_Binder_Plate_distance = Data_Binder_Plate_distance.Value
    Data_Binder_Plate_height = part1.Parameters.Item("Data_Binder_Plate_height")
    Data_Binder_Plate_height = Data_Binder_Plate_height.Value
    Data_Binder_Plate_Bolt_hole = part1.Parameters.Item("Data_Binder_Plate_Bolt_hole")
    Data_Binder_Plate_Bolt_hole = Data_Binder_Plate_Bolt_hole.Value
    # --------------------------------------------------------------讀取數據
    partDocument1.save()
    time.sleep(2)
    partDocument1.Close()  # 關閉檔案
    # ==================================================================================================製作沖頭壓板
    partDocument1 = documents1.Open(gvar.open_path + "Binder_Plate_40.CATPart")  # 開啟壓板母檔
    time.sleep(2)
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    # ------------------------------------------------設定條件
    Binder_Plate_outside = part1.Parameters.Item("Binder_Plate_outside")
    Data_Binder_Plate_outside = int(Data_Binder_Plate_outside) + 1
    Binder_Plate_outside.Value = Data_Binder_Plate_outside + 1
    # 固定條長度(中心到下料)=(外圈圓半徑-1)
    Binder_Plate_length = part1.Parameters.Item("Binder_Plate_length")
    Data_Binder_Plate_length = int(Data_Binder_Plate_length) + 1
    Binder_Plate_length.Value = Data_Binder_Plate_length - 1
    # 固定條寬度=靴齒部到靴齒部距離
    Binder_Plate_wide = part1.Parameters.Item("Binder_Plate_wide")
    Binder_Plate_wide.Value = Data_Binder_Plate_wide - 1
    # 固定片厚度
    Binder_Plate_thickness = part1.Parameters.Item("Binder_Plate_thickness")
    Binder_Plate_thickness.Value = 3
    # 固定條數量
    Binder_Plate_number = part1.Parameters.Item("Binder_Plate_number")
    Binder_Plate_number.Value = 4
    # 工站間距
    Binder_Plate_distance = part1.Parameters.Item("Binder_Plate_distance")
    Binder_Plate_distance.Value = Data_Binder_Plate_distance
    # 壓板高度
    Binder_Plate_height = part1.Parameters.Item("Binder_Plate_height")
    Binder_Plate_height.Value = float(gvar.strip_parameter_list[20]) + float(gvar.strip_parameter_list[17]) + float(
        gvar.strip_parameter_list[1])
    # 壓板螺栓孔
    Binder_Plate_Bolt_hole = part1.Parameters.Item("Binder_Plate_Bolt_hole")
    Binder_Plate_Bolt_hole.Value = 5
    part1.Update()
    product1.PartNumber = "Binder_Plate_" + str(op_number) + ""
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Binder_Plate_" + str(op_number) + "")
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
    strParam4.ValuateFromString(gvar.strip_parameter_list[33])
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam5 = parameters5.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString("")
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
    # ---------------------------------------------------存檔
    partDocument1.SaveAs(gvar.save_path + "Binder_Plate_" + str(op_number) + ".CATPart")  # 存檔的檔案名稱
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    partDocument1.Close()  # 關閉檔案

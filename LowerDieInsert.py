import win32com.client as win32
import global_var as gvar
import defs
import InsertDef
import time
import PunchDef
import math


def LowerDieInsert(now_plate_line_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    g = now_plate_line_number
    total_op_number = int(gvar.strip_parameter_list[2])
    cut_cavity_insert_machining_explanation_shape = 0  # ------------------------------------------------------------加工說明
    open_name1 = "QR_allotype_insert"
    (insert_interferance_count) = InsertDef.InsertInterferance((gvar.StripDataList[39]), "_allotype_cut_line_",
                                                               open_name1,
                                                               now_plate_line_number)
    for now_op_number in range(1, 1 + total_op_number):
        n = now_op_number
        op_number = 10 * n
        # ==============================================================================================================基本模組↓
        if gvar.StripDataList[38][g][n] > 0:
            xi = int()
            if gvar.StripDataList[38][g][n] > 1:
                xi = gvar.StripDataList[38][g][n]
            for i in range(1, 1 + 1):
                Row = "op" + str(op_number) + " " + str(i)
                (h1) = defs.ExcelSearch("rule", "Shape", Row, "形狀")
                # ===================================有無干涉 要哪種入子
                insert_interferance_no_delete = 0
                insert_interferance_decide_Excavation = 0
                for ii in range(1, 1 + total_op_number):
                    for qq in range(1, 1 + 10):
                        if insert_interferance_count[ii][qq][1] == n and insert_interferance_count[ii][qq][2] == i:
                            insert_interferance_now = ii
                            insert_interferance_decide_Excavation = 1
                # ====================================
                partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
                time.sleep(1)
                if h1 == 0:
                    partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert.CATPart")
                elif h1 == 1:
                    partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_1.CATPart")
                elif h1 == 2:
                    partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_2.CATPart")
                elif h1 == 3:
                    partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_3.CATPart")
                # ======================================
                defs.window_change(partDocument1, partDocument2)
                # ======================================
                # ===================================================基本模組================================================
                if h1 == 0:
                    partDocument1 = catapp.ActiveDocument
                    part1 = partDocument1.Part
                    length = [None] * 99
                    # ------------------------------------------------------------↓   參數宣告
                    length[11] = part1.Parameters.Item("cut_cavity_insert_height")
                    g = now_plate_line_number
                    length[11].Value = float(gvar.strip_parameter_list[26])
                    length[12] = part1.Parameters.Item("die_open_height")
                    length[12].Value = 0  # (die_open_height)
                    length[13] = part1.Parameters.Item("insert_line")  # 多出來的
                    length[13].Value = 5  # 多出來的
                    length[33] = part1.Parameters.Item("gap")
                    length[33].Value = 0.01  # (lower_die_cavity_plate_space)  # 間隙
                    # ------------------------------------------------------------↑
                    # ------------------------------------------------------------↓   判斷是否對稱名稱
                    selection3 = partDocument1.Selection
                    selection3.Clear()
                    selection3.Search(
                        "Name=plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_symmetric_" + str(i))
                    # ------------------------------------------------------------↑
                    # ------------------------------------------------------------↓   判斷是否對稱名稱
                    if selection3.Count > 0:
                        part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                            "die\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_symmetric_" + str(
                                i))  # 草圖置換
                    else:
                        part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                            "die\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i))  # 草圖置換
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
                    part1.Update()
                    if xi > 1:
                        PunchDef.clash_change(xi, op_number)
                        xi = 0
                    # ------------------------------------------------------------↓   整數化
                    q = length[34].Value
                    R = length[35].Value
                    length[36].Value = math.ceil(q)
                    length[37].Value = math.ceil(R)
                    part1.Update()
                    # ------------------------------------------------------------↑
                    # =========================================干涉
                    body_name1 = "cut_cavity_insert"
                    part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                        "die\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i))  # 草圖置換
                    part1.Update()
                    part1.UpdateObject(part1.Bodies.Item("Body.2"))
                    product1 = partDocument1.getItem("Part1")
                    X = 0  # 名稱命名
                    if i >= 10:
                        X = ""  # 名稱命名
                    part1.Update()
                    product1.PartNumber = "op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(i)  # 樹枝圖名稱
                    # ====↓設定性質↓=====================================
                    partDocument1 = catapp.ActiveDocument
                    product1 = partDocument1.getItem("op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(i))
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
                    strParam8.ValuateFromString("L1 :1- 線割, 單 %%P 0.01")
                    product1 = product1.ReferenceProduct
                    parameters9 = product1.UserRefProperties
                    strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
                    strParam9.ValuateFromString("A : (下模螺絲)")
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
                    strParam19.ValuateFromString("AP: (A沖沖孔)")
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
                        selection1.Search("Name=plate_line_1*,all")
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
                    partDocument1.SaveAs(
                        gvar.save_path + "op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(
                            i) + ".CATPart")  # 存檔的檔案名稱
                    time.sleep(1)
                    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
                    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
                    # ---------使用迴圈，建立關連↑-------------
                    part1.UpdateObject(part1.Bodies.Item("Body.2"))
                    partDocument1.Close()
                # =========================================可修改模組1=============================================
                if h1 == 1:
                    catapp = win32.Dispatch('CATIA.Application')
                    partDocument1 = catapp.ActiveDocument
                    part1 = partDocument1.Part
                    length = [None] * 99
                    length[0] = part1.Parameters.Item("Y_min_to_side")
                    length[1] = part1.Parameters.Item("X_max_to_side")
                    length[2] = part1.Parameters.Item("Y_max_to_side")
                    length[3] = part1.Parameters.Item("X_min_to_side")
                    length[4] = part1.Parameters.Item("Chamfer_1")
                    length[5] = part1.Parameters.Item("Chamfer_2")
                    length[6] = part1.Parameters.Item("Chamfer_4")
                    length[7] = part1.Parameters.Item("Chamfer_3")
                    length[8] = part1.Parameters.Item("cut_cavity_insert_height")
                    g = now_plate_line_number
                    length[8].Value = float(gvar.strip_parameter_list[26])
                    length[9] = part1.Parameters.Item("die_open_height")
                    length[9].Value = 0.  # (die_open_height)
                    length[10] = part1.Parameters.Item("insert_line")  # 多出來的
                    length[10].Value = 5
                    length[33] = part1.Parameters.Item("gap")
                    length[33].Value = 0.01  # (lower_die_cavity_plate_space)  # 間隙
                    part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                        "die\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i))  # 草圖置換
                    part1.Update()
                    length[34] = part1.Parameters.Item("x_to_x")
                    length[35] = part1.Parameters.Item("y_to_y")
                    length[36] = part1.Parameters.Item("int_x")
                    length[37] = part1.Parameters.Item("int_y")
                    q = length[34].Value
                    R = length[35].Value
                    length[36].Value = math.ceil(q)
                    length[37].Value = math.ceil(R)
                    part1.Update()
                    part1.UpdateObject(part1.Bodies.Item("Body.2"))
                    product1 = partDocument1.getItem("Part1")
                    X = 0  # 名稱命名
                    if i >= 10:
                        X = ""  # 名稱命名
                    product1.PartNumber = "op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(i)  # 樹枝圖名稱
                    # ====↓設定性質↓=====================================
                    partDocument1 = catapp.ActiveDocument
                    product1 = partDocument1.getItem("op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(i))
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
                    strParam8.ValuateFromString("L1 :1- 線割, 單 %%P 0.01")
                    product1 = product1.ReferenceProduct
                    parameters9 = product1.UserRefProperties
                    strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
                    strParam9.ValuateFromString("A : (下模螺絲)")
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
                    strParam19.ValuateFromString("AP: (A沖沖孔)")
                    product1 = product1.ReferenceProduct
                    # ====↑設定性質↑=====================================
                    part1.Update()
                    # --------------↓刪除不需要的Data↓--------------
                    selection1 = partDocument1.Selection
                    if now_plate_line_number == 1:  # --------模板1
                        selection1.Clear()
                        selection1.Search("Name=plate_line_2*,all")
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
                    partDocument1.SaveAs(gvar.save_path + "op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(
                        i) + ".CATPart")  # 存檔的檔案名稱
                    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
                    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
                    # ---------使用迴圈，建立關連↑-------------
                    part1.UpdateObject(part1.Bodies.Item("Body.2"))
                try:
                    partDocument1.Close()
                except:
                    pass
                # =========================================可修改模組2=============================================
                if h1 == 2:
                    catapp = win32.Dispatch('CATIA.Application')
                    partDocument1 = catapp.ActiveDocument
                    part1 = partDocument1.Part
                    length = [None] * 99
                    length[18] = part1.Parameters.Item("Y_min_to_side")
                    length[19] = part1.Parameters.Item("X_max_to_side")
                    length[20] = part1.Parameters.Item("Y_max_to_side")
                    length[21] = part1.Parameters.Item("X_min_to_side")
                    length[22] = part1.Parameters.Item("Chamfer_1")
                    length[23] = part1.Parameters.Item("Chamfer_2")
                    length[24] = part1.Parameters.Item("Chamfer_4")
                    length[25] = part1.Parameters.Item("Chamfer_3")
                    length[26] = part1.Parameters.Item("cut_cavity_insert_height")
                    g = now_plate_line_number
                    length[26].Value = float(gvar.strip_parameter_list[26])
                    length[27] = part1.Parameters.Item("die_open_height")
                    length[27].Value = 0  # (die_open_height)
                    length[28] = part1.Parameters.Item("insert_line")  # 多出來的
                    length[28].Value = 5
                    length[33] = part1.Parameters.Item("gap")
                    length[33].Value = 0.01  # (lower_die_cavity_plate_space)  # 間隙
                    part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                        "die\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i))  # 草圖置換
                    part1.Update()
                    length[34] = part1.Parameters.Item("x_to_x")
                    length[35] = part1.Parameters.Item("y_to_y")
                    length[36] = part1.Parameters.Item("int_x")
                    length[37] = part1.Parameters.Item("int_y")
                    q = length[34].Value
                    R = length[35].Value
                    length[36].Value = math.ceil(q)
                    length[37].Value = math.ceil(R)
                    part1.Update()
                    part1.UpdateObject(part1.Bodies.Item("Body.2"))
                    product1 = partDocument1.getItem("Part1")
                    X = 0  # 名稱命名
                    if i >= 10:
                        X = ""  # 名稱命名
                    product1.PartNumber = "op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(i)  # 樹枝圖名稱
                    # ====↓設定性質↓=====================================
                    partDocument1 = catapp.ActiveDocument
                    product1 = partDocument1.getItem("op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(i))
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
                    strParam8.ValuateFromString("L1 :1- 線割, 單 %%P 0.01")
                    product1 = product1.ReferenceProduct
                    parameters9 = product1.UserRefProperties
                    strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
                    strParam9.ValuateFromString("A : (下模螺絲)")
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
                    strParam19.ValuateFromString("AP: (A沖沖孔)")
                    product1 = product1.ReferenceProduct
                    # ====↑設定性質↑=====================================
                    part1.Update()
                    # --------------↓刪除不需要的Data↓--------------
                    selection1 = partDocument1.Selection
                    if now_plate_line_number == 1:  # --------模板1
                        selection1.Clear()
                        selection1.Search("Name=plate_line_2*,all")
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
                    partDocument1.SaveAs(
                        gvar.save_path + "op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(
                            i) + ".CATPart")  # 存檔的檔案名稱
                    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
                    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber
                    # ---------使用迴圈，建立關連↑-------------
                    part1.UpdateObject(part1.Bodies.Item("Body.2"))
                    partDocument1.Close()
                # =========================================可修改模組3=============================================
                if h1 == 3:
                    catapp = win32.Dispatch('CATIA.Application')
                    partDocument1 = catapp.ActiveDocument
                    part1 = partDocument1.Part
                    length = [None] * 99
                    length[14] = part1.Parameters.Item("cut_cavity_insert_height")
                    length[14].Value = float(gvar.strip_parameter_list[26])
                    length[15] = part1.Parameters.Item("die_open_height")
                    length[15].Value = 0  # (die_open_height)
                    length[16] = part1.Parameters.Item("insert_line")  # 多出來的
                    length[16].Value = 5
                    length[33] = part1.Parameters.Item("gap")
                    length[33].Value = 0.01  # (lower_die_cavity_plate_space)  # 間隙
                    part1.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                        "die\plate_line_" + str(g) + "_op" + str(op_number) + "_cut_line_" + str(i))  # 草圖置換
                    part1.Update()
                    length[34] = part1.Parameters.Item("x_to_x")
                    length[35] = part1.Parameters.Item("y_to_y")
                    length[36] = part1.Parameters.Item("int_x")
                    length[37] = part1.Parameters.Item("int_y")
                    q = length[34].Value
                    R = length[35].Value
                    length[36].Value = math.ceil(q)
                    length[37].Value = math.ceil(R)
                    part1.Update()
                    part1.UpdateObject(part1.Bodies.Item("Body.2"))
                    product1 = partDocument1.getItem("Part1")
                    X = 0  # 名稱命名
                    if i >= 10:
                        X = ""  # 名稱命名
                    product1.PartNumber = "op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(i)  # 樹枝圖名稱
                    # ====↓設定性質↓=====================================
                    partDocument1 = catapp.ActiveDocument
                    product1 = partDocument1.getItem("op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(i))
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
                    strParam8.ValuateFromString("L1 :1- 線割, 單 %%P 0.01")
                    product1 = product1.ReferenceProduct
                    parameters9 = product1.UserRefProperties
                    strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
                    strParam9.ValuateFromString("A : (下模螺絲)")
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
                    strParam19.ValuateFromString("AP: (A沖沖孔)")
                    product1 = product1.ReferenceProduct
                    # ====↑設定性質↑=====================================
                    part1.Update
                    # --------------↓刪除不需要的Data↓--------------

                    ##selection1 As Selection
                    selection1 = partDocument1.Selection
                    if now_plate_line_number == 1:  # --------模板1
                        selection1.Clear()
                        selection1.Search("Name=plate_line_2*,all")
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
                    partDocument1.SaveAs(
                        gvar.save_path + "op" + str(op_number) + "_cut_cavity_insert_" + str(X) + str(
                            i) + ".CATPart")  # 存檔的檔案名稱
                    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
                    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber
                    # ---------使用迴圈，建立關連↑-------------
                    part1.UpdateObject(part1.Bodies.Item("Body.2"))
                    partDocument1.Close()
        if gvar.StripDataList[27][g][n] > 0:
            for j in range(1, 1 + gvar.StripDataList[27][g][n]):
                # --------------------------------------------------------------------------------------------------------↓對稱入子
                insert_line_name = [''] * 3
                point_name = [""] * 7
                X = 0  # 名稱命名
                if j >= 10:
                    X = ""  # 名稱命名
                # ------------------------------------------------------------↓參數名稱
                insert_name = "op" + str(op_number) + "_cutting_cavity_d_left_insert_" + str(X) + str(j)
                insert_line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_d_cutting_" + str(j)
                insert_line_name[2] = "plate_line_" + str(g) + "_op" + str(
                    op_number) + "_cut_punch_d_insert_surface_" + str(j)
                point_name[1] = "cut_base_point_X_min"
                point_name[2] = "op" + str(op_number) + "_d_cut_X_min_point"
                point_name[3] = "cut_base_point_X_max"
                point_name[4] = "op" + str(op_number) + "_d_cut_X_max_point"
                point_name[5] = "cut_base_point_Z"
                point_name[6] = "op" + str(op_number) + "_d_cut_Z_point"
                # ------------------------------------------------------------↑
                left_insert(g, n, j, insert_name, insert_line_name)  # 左邊入子
                partDocument2 = documents1.Open(gvar.open_path + "Data1.CATPart")
                partDocument1 = documents1.Open(gvar.open_path + "cut_cavity_insert_shear.CATPart")
                # ======================================
                defs.window_change(partDocument2, partDocument1)
                # ======================================
                partDocument1 = catapp.ActiveDocument
                part1 = partDocument1.Part
                bodies1 = part1.Bodies
                body1 = bodies1.Item("Body.2")
                body1.Name = "cut_cavity_insert_shear"
                parameters1 = part1.Parameters
                strParam1 = parameters1.Item("Type_shear")
                sketches1 = body1.Sketches
                sketch1 = sketches1.Item("sketch_Coordinate")
                length[30] = part1.Parameters.Item("cut_cavity_insert_height")
                length[30].Value = float(gvar.strip_parameter_list[26])
                length[31] = part1.Parameters.Item("die_open_height")
                length[31].Value = 0  # (die_open_height)
                strParam1.Value = "B_down"
                part1.Parameters.Item("cut_line_formula_2").OptionalRelation.Modify(
                    "die\\" + insert_line_name[1])  # 草圖置換
                part1.Parameters.Item("shear_Surface").OptionalRelation.Modify("die\\" + insert_line_name[2])  # 草圖置換
                part1.UpdateObject(sketch1)
                length[34] = part1.Parameters.Item("x_to_x_shear")
                length[35] = part1.Parameters.Item("y_to_y_shear")
                length[36] = part1.Parameters.Item("int_x_shear")
                length[37] = part1.Parameters.Item("int_y_shear")
                length[36].Value = int(length[34].Value)
                length[37].Value = int(length[35].Value)
                # ------------------------------------------------------------↓     改變點位置
                for i in range(1, 1 + 3):
                    hybridShapes1 = body1.HybridShapes
                    hybridShapePointCoord1 = hybridShapes1.Item(point_name[i * 2 - 1])
                    hybridShapePointExplicit1 = parameters1.Item(point_name[i * 2])
                    reference1 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
                    hybridShapePointCoord1.PtRef = reference1
                    hybridShapePointCoord1.X.Value = 0
                    hybridShapePointCoord1.Y.Value = 0
                    hybridShapePointCoord1.Z.Value = 0
                # ------------------------------------------------------------↑
                part1.Update()
                # ---------------------------------------------------------------------------------------------------------------↑ 對稱入子
                product1 = partDocument1.getItem("Part1")
                X = 0  # 名稱命名
                if j >= 10:
                    X = ""  # 名稱命名
                product1.PartNumber = "op" + str(op_number) + "_cutting_cavity_d_insert_" + str(X) + str(j)  # 樹枝圖名稱
                # ====↓設定性質↓=====================================
                partDocument1 = catapp.ActiveDocument
                product1 = partDocument1.getItem("op" + str(op_number) + "_cutting_cavity_d_insert_" + str(X) + str(j))
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
                strParam8.ValuateFromString("L1 :1- 線割, 單 %%P 0.01")
                product1 = product1.ReferenceProduct
                parameters9 = product1.UserRefProperties
                strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
                strParam9.ValuateFromString("A :4-M11 攻穿, 正面沉頭 %%C 15 深15mm(下模螺絲)")
                product1 = product1.ReferenceProduct
                parameters14 = product1.UserRefProperties
                strParam12 = parameters14.CreateString("HP", "")  # 合銷孔
                strParam12.ValuateFromString("")
                product1 = product1.ReferenceProduct
                parameters15 = product1.UserRefProperties
                strParam13 = parameters15.CreateString("B", "")  # B型引導沖孔
                strParam13.ValuateFromString
                ""
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
                    selection1.Search(
                        "Name=plate_line_2*,all")
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
                        selection1.Search(
                            "Name=*_op" + str(o) + "0_*,all")
                        if selection1.Count != 0:
                            selection1.Delete()
                        selection1.Clear()
                if now_plate_line_number == 2:  # --------模板2
                    selection1.Clear()
                    selection1 = partDocument1.Selection
                    selection1.Search("Name=plate_line_1*,all")
                    selection1.Delete()
                    selection1.Clear()
                    for o in range(1, 1 + (now_op_number - 1)):
                        selection1.Clear()
                        selection1.Search()
                        "Name=*_op" + str(o) + "0_*,all"
                        if selection1.Count != 0:
                            selection1.Delete()
                        selection1.Clear()
                    for o in range((now_op_number + 1), 1 + total_op_number):
                        selection1.Clear()
                        selection1.Search()
                        "Name=*_op" + str(o) + "0_*,all"
                        if selection1.Count != 0:
                            selection1.Delete()
                        selection1.Clear()
                # --------------↑刪除不需要的Data↑--------------
                part1.Update()
                time.sleep(2)
                partDocument1.SaveAs(
                    gvar.save_path + "op" + str(op_number) + "_cutting_cavity_d_insert_" + str(X) + str(
                        j) + ".CATPart")  # 存檔的檔案名稱
                gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
                gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
                # ---------使用迴圈，建立關連↑-------------
                part1.Update()
                partDocument1.Close()

        if gvar.StripDataList[28][g][n] > 0:
            for k in range(1, 1 + gvar.StripDataList[28][g][n]):
                insert_line_name = [''] * 3
                point_name = [""] * 7
                # ---------------------------------------------------------------------------------------------------------------------------------------------------↓對稱入子
                # 數字二位數化,1~10改為01~10
                X = 0  # 名稱命名
                if k >= 10:
                    X = ""
                    # ------------------------------------------------------------↓參數名稱
                insert_name = "op" + str(op_number) + "_cutting_cavity_u_left_insert_" + str(X) + str(k)
                insert_line_name[1] = "plate_line_" + str(g) + "_op" + str(op_number) + "_cut_punch_u_cutting_" + str(k)
                insert_line_name[2] = "plate_line_" + str(g) + "_op" + str(
                    op_number) + "_cut_punch_u_insert_surface_" + str(k)
                point_name[1] = "cut_base_point_X_min"
                point_name[2] = "op" + str(op_number) + "_u_cut_X_min_point"
                point_name[3] = "cut_base_point_X_max"
                point_name[4] = "op" + str(op_number) + "_u_cut_X_max_point"
                point_name[5] = "cut_base_point_Z"
                point_name[6] = "op" + str(op_number) + "_u_cut_Z_point"
                # ------------------------------------------------------------↑
                left_insert(g, n, k, insert_name, insert_line_name)  # 左邊入子
                partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
                partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_shear.CATPart")
                # ======================================
                defs.window_change(partDocument1, partDocument2)
                # ======================================
                partDocument1 = catapp.ActiveDocument
                part1 = partDocument1.Part
                bodies1 = part1.Bodies
                body1 = bodies1.Item("Body.2")
                body1.Name = "cut_cavity_insert_shear"
                parameters1 = part1.Parameters
                strParam1 = parameters1.Item("Type_shear")
                sketches1 = body1.Sketches
                sketch1 = sketches1.Item("sketch_Coordinate")
                length[30] = part1.Parameters.Item("cut_cavity_insert_height")
                length[30].Value = float(gvar.strip_parameter_list[26])
                length[31] = part1.Parameters.Item("die_open_height")
                length[31].Value = 0  # (die_open_height)
                strParam1.Value = "A_up"
                part1.Parameters.Item("cut_line_formula_2").OptionalRelation.Modify(
                    "die\\" + insert_line_name[1])  # 草圖置換
                part1.Parameters.Item("shear_Surface").OptionalRelation.Modify("die\\" + insert_line_name[2])  # 草圖置換
                part1.UpdateObject(sketch1)
                length[34] = part1.Parameters.Item("x_to_x_shear")
                length[35] = part1.Parameters.Item("y_to_y_shear")
                length[36] = part1.Parameters.Item("int_x_shear")
                length[37] = part1.Parameters.Item("int_y_shear")
                length[36].Value = int(length[34].Value)
                length[37].Value = int(length[35].Value)
                # ------------------------------------------------------------↓     改變點位置
                for i in range(1, 1 + 3):
                    hybridShapes1 = body1.HybridShapes
                    hybridShapePointCoord1 = hybridShapes1.Item(point_name[i * 2 - 1])
                    hybridShapePointExplicit1 = parameters1.Item(point_name[i * 2])
                    reference1 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
                    hybridShapePointCoord1.PtRef = reference1
                    hybridShapePointCoord1.X.Value = 0
                    hybridShapePointCoord1.Y.Value = 0
                    hybridShapePointCoord1.Z.Value = 0
                # ------------------------------------------------------------↑
                part1.Update()
                product1 = partDocument1.getItem("Part1")
                product1.PartNumber = "op" + str(op_number) + "_cutting_cavity_u_insert_" + str(X) + str(k)  # 樹枝圖名稱
                # ====↓設定性質↓=====================================
                ##partDocument1 As PartDocument
                partDocument1 = catapp.ActiveDocument
                product1 = partDocument1.getItem(
                    "op" + str(op_number) + "_cutting_cavity_u_left_insert_" + str(X) + str(k))
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
                strParam8.ValuateFromString("L1 :1- 線割, 單 %%P 0.01")
                product1 = product1.ReferenceProduct
                parameters9 = product1.UserRefProperties
                strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
                strParam9.ValuateFromString("A :4-M11 攻穿, 正面沉頭 %%C 15 深15mm(下模螺絲)")
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
                time.sleep(2)
                partDocument1.SaveAs(
                    gvar.save_path + "op" + str(op_number) + "_cutting_cavity_u_left_insert_" + str(X) + str(
                        k) + ".CATPart")  # 存檔的檔案名稱
                gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
                gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
                # ---------使用迴圈，建立關連↑-------------
                part1.Update()
                partDocument1.Close()
                # ---------------------------------------------------------------------------------------------------------------------------------------------------↑ 對稱入子
        # ====================================================================================↓其他形式入子
        if gvar.StripDataList[4][g][n] > 0:  # 補強入子
            # for for_counter in range(1, int(gvar.StripDataList[4][g][n]) + 1):
            #     parameter_digital1 = 0
            #     Reinforcement_Ecxavation(g, n, for_counter, parameter_digital1)
            pass  # 未使用
        # ----------------------------------------------------------------------------------------------↓快拆沖頭
        if gvar.StripDataList[29][g][n] > 0:  # '沖切沖頭_右
            # data_type = "line"
            # data_number = int(gvar.StripDataList[29][g][n])
            # part_name = "op" + str(op_number) + "_right_quickly_remove_cut_punch_"
            # PunchDef.quickly_remove_punch(data_type, data_number, part_name)
            pass  # 未使用
        if gvar.StripDataList[30][g][n] > 0:  # 沖切沖頭_左
            # data_type = "line"
            # data_number = int(gvar.StripDataList[30][g][n])
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
        if gvar.StripDataList[33][g][n] > 0:  # 折彎沖頭_右
            # data_type = "surface"
            # data_number = gvar.StripDataList[33][g][n]
            # part_name = "op" + str(op_number) + "_right_quickly_remove_bending_punch_"
            # PunchDef.quickly_remove_punch(data_type, data_number, part_name)
            pass  # 未使用
        if gvar.StripDataList[34][g][n] > 0:  # 折彎沖頭_左
            # data_type = "surface"
            # data_number = gvar.StripDataList[34][g][n]
            # part_name = "op" + str(op_number) + "_left_quickly_remove_bending_punch_"
            # PunchDef.quickly_remove_punch(data_type, data_number, part_name)
            pass  # 未使用
        if gvar.StripDataList[35][g][n] > 0:  # 折彎沖頭_上
            # data_type = "surface"
            # data_number = gvar.StripDataList[35][g][n]
            # part_name = "op" + str(op_number) + "_up_quickly_remove_bending_punch_"
            # PunchDef.quickly_remove_punch(data_type, data_number, part_name)
            pass  # 未使用
        if gvar.StripDataList[36][g][n] > 0:  # 折彎沖頭_下
            # data_type = "surface"
            # data_number = gvar.StripDataList[36][g][n]
            # part_name = "op" + str(op_number) + "_down_quickly_remove_bending_punch_"
            # PunchDef.quickly_remove_punch(data_type, data_number, part_name)
            pass  # 未使用
        # ----------------------------------------------------------------------------------------------↑快拆沖頭
        if gvar.StripDataList[37][g][n] > 0:
            plate_line_A_punch(now_plate_line_number, now_op_number)
        if gvar.StripDataList[53][g][n] > 0:
            # unnomal_cut_line_T
            pass  # 未使用
        if gvar.StripDataList[54][g][n] > 0:
            # unnomal_cut_line_I
            pass  # 未使用
        if gvar.StripDataList[55][g][n] > 0:
            # unnomal_cut_line_M
            pass  # 未使用
        X = "0"  # 名稱命名
        # --------------------------------------------------------------------------------------------異型沖頭
        if gvar.StripDataList[39][g][n] > 0:
            # for now_data_number in range (1, 1 + plate_line_allotype_cut_line_number(g, n)):
            file_name = "cut_cavity_insert"
            element_name0 = "_allotype_cut_line_"
            # cut_allotype_insert
            pass  # 未使用
        # --------------------------------------------------------------------------------------------異型沖頭
        # --------------------------------------------------------------------------------------------打凸包沖頭_左
        if gvar.StripDataList[21][g][n] > 0:
            # if gvar.StripDataList[67][g][n] > 0:
            #     for now_data_number in range (1, 1 + gvar.StripDataList[67][g][n]):
            #         if now_data_number >= 10:
            #             X = ""  # 名稱命名
            #         emboss_forming_punch_direction = "left"
            #         file_name = "op" + str(op_number) + "_left_emboss_forming_down_punch_" + str(X) + str(now_data_number) + ".CATPart"
            #         down_shoulder_insert(file_name)
            pass  # 未使用
        # --------------------------------------------------------------------------------------------打凸包沖頭_左
        # --------------------------------------------------------------------------------------------打凸包沖頭_右
        if gvar.StripDataList[22][g][n] > 0:
            # if gvar.StripDataList[67][g][n] > 0:
            #     for now_data_number in range (1, 1 + gvar.StripDataList[67][g][n]):
            #         if now_data_number >= 10:
            #             X = ""  # 名稱命名
            #         emboss_forming_punch_direction = "right"
            #         file_name = "op" + op_number + "_right_emboss_forming_down_punch_" + X + now_data_number + ".CATPart"
            #         Call
            #         down_shoulder_insert(file_name)
            pass  # 未使用
        # --------------------------------------------------------------------------------------------打凸包沖頭_右
        # ====================================================================================其他形式入子
        # --------------------------------------------------------------------------------------------整形模穴
        if gvar.StripDataList[74][g][n] > 0:
            # for now_data_number = (1, 1 + gvar.StripDataList[74][g][n])
            # F_bending_cavity
            pass  # 未使用


def left_insert(g, n, j, insert_name, insert_line_name):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
    partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_shear.CATPart")
    defs.window_change(partDocument1, partDocument2)
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    sketches1 = body1.Sketches
    sketch1 = sketches1.Item("cut_line_formula_2_Sketch")
    body1.Name = "cut_cavity_insert_shear"
    parameters1 = part1.Parameters
    parameters1.Item("cut_line_formula_2").OptionalRelation.Modify("die\\" + insert_line_name[1])  # 草圖置換
    parameters1.Item("shear_Surface").OptionalRelation.Modify("die\\" + insert_line_name[2])  # 草圖置換
    part1.UpdateObject(sketch1)
    length30 = part1.Parameters.Item("cut_cavity_insert_height")
    length30.Value = 30
    length31 = parameters1.Item("die_open_height")
    length31.Value = 0
    strParam1 = parameters1.Item("Type_shear")
    strParam1.Value = "C_left"
    length34 = parameters1.Item("x_to_x_shear")
    length35 = parameters1.Item("y_to_y_shear")
    length36 = parameters1.Item("int_x_shear")
    length37 = parameters1.Item("int_y_shear")
    length36.Value = int(length34.Value)
    length37.Value = int(length35.Value)
    part1.Update()
    product1 = partDocument1.getItem("Part1")
    # 數字二位數化,1~10改為01~10
    X = 0  # 名稱命名
    if j >= 10:
        X = ''  # 名稱命名
    product1.PartNumber = insert_name  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem(insert_name)
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
    strParam8.ValuateFromString("L1 :1- 線割, 單 %%P 0.01")
    product1 = product1.ReferenceProduct
    parameters9 = product1.UserRefProperties
    strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
    strParam9.ValuateFromString("A :4-M11 攻穿, 正面沉頭 %%C 15 深15mm(下模螺絲)")
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
    if g == 1:  # --------模板1
        selection1.Clear()
        selection1.Search("Name=plate_line_2*,all")
        if selection1.Count != 0:
            selection1.Delete()
        selection1.Clear()
        for o in range(1, g):
            selection1.Clear()
            selection1.Search("Name=*_op" + str(o) + "0_*,all")
            if selection1.Count != 0:
                selection1.Delete()
            selection1.Clear()
        for o in range((n + 1), gvar.strip_parameter_list[2] + 1):
            selection1.Clear()
            selection1.Search("Name=*_op" + str(o) + "0_*,all")
            if selection1.Count != 0:
                selection1.Delete()
            selection1.Clear()
    if g == 2:  # --------模板2
        selection1.Clear()
        selection1 = partDocument1.Selection
        selection1.Search("Name=plate_line_1*,all")
        selection1.Delete()
        selection1.Clear()
        for o in range(1, g):
            selection1.Clear()
            selection1.Search("Name=*_op" + str(o) + "0_*,all")
            if selection1.Count != 0:
                selection1.Delete()
            selection1.Clear()
        for o in range((n + 1), gvar.strip_parameter_list[2] + 1):
            selection1.Clear()
            selection1.Search("Name=*_op" + str(o) + "0_*,all")
            if selection1.Count != 0:
                selection1.Delete()
            selection1.Clear()
    # --------------↑刪除不需要的Data↑--------------
    part1.Update()
    time.sleep(2)
    partDocument1.SaveAs(gvar.save_path + insert_name + ".CATPart")
    # 存檔的檔案名稱
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    # ---------使用迴圈，建立關連↑-------------
    part1.Update()
    partDocument1.Close()


def plate_line_A_punch(now_plate_line_number, now_op_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    length = [None] * 99
    cut_cavity_insert_machining_explanation_shape = 0  # ------------------------------------------------------------加工說明
    # ===================================================特別模組↓================================================
    cut_cavity_insert_machining_explanation_shape = 0  # ------------------------------------------------------------加工說明
    g = now_plate_line_number
    n = now_op_number
    op_number = 10 * n
    total_op_number = int(gvar.strip_parameter_list[2])
    # for i in range(1, 1 + gvar.StripDataList[37][g][n]):
    i = 1
    Row = "op" + str(op_number) + " " + str(i)
    (h1) = defs.ExcelSearch("rule", "Shape", Row, "形狀")
    partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
    if h1 == 0:
        partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert.CATPart")
    elif h1 == 1:
        partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_1.CATPart")
    elif h1 == 2:
        partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_2.CATPart")
    elif h1 == 3:
        partDocument2 = documents1.Open(gvar.open_path + "cut_cavity_insert_3.CATPart")
    # ======================================
    defs.window_change(partDocument1, partDocument2)
    # ======================================
    # ===================================================基本模組================================================
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    length = [None] * 99
    # ---------------------------------------------------------------------------------------------------------------------------------------------------↓對稱入子
    # ------------------------------------------------------------↓   參數宣告
    length[11] = part1.Parameters.Item("cut_cavity_insert_height")
    g = now_plate_line_number
    length[11].Value = float(gvar.strip_parameter_list[26])
    length[12] = part1.Parameters.Item("die_open_height")
    length[12].Value = 0  # (die_open_height)
    length[13] = part1.Parameters.Item("insert_line")  # 多出來的
    length[13].Value = 5  # 多出來的
    length[33] = part1.Parameters.Item("gap")
    length[33].Value = 0.01  # (lower_die_cavity_plate_space)  # 間隙
    # ------------------------------------------------------------↓   判斷是否對稱名稱
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
    # ------------------------------------------------------------↓   判斷是否對稱名稱
    boolParam1 = part1.Parameters.Item("symmetry_switch")
    if selection3.Count > 0:
        boolParam1.Value = True
    else:
        boolParam1.Value = False
    # ------------------------------------------------------------↑
    part1.Update()
    # ------------------------------------------------------------↓   參數宣告
    length[34] = part1.Parameters.Item("x_to_x")
    length[35] = part1.Parameters.Item("y_to_y")
    length[36] = part1.Parameters.Item("int_x")
    length[37] = part1.Parameters.Item("int_y")
    if gvar.StripDataList[37][g][n] > 1:
        xi = gvar.StripDataList[37][g][n]
        PunchDef.A_punch_clash_change(xi, op_number)
        # i = i + (xi - 1)
        xi = 0
    # ------------------------------------------------------------↓   先放大入子以免更換時  孔>入子塊
    length[36].Value = 500
    length[37].Value = 500
    # ------------------------------------------------------------↑
    part1.Update()
    # ------------------------------------------------------------↓   整數化
    q = length[34].Value
    R = length[35].Value
    length[36].Value = math.ceil(q)
    length[37].Value = math.ceil(R)
    # ------------------------------------------------------------↑
    part1.Update()
    # ---------------------------------------------------------------------------------------------------------------------------------------------------↑對稱入子
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    product1 = partDocument1.getItem("Part1")
    # 數字二位數化,1~10改為01~10
    X = 0  # 名稱命名
    if i >= 10:
        X = ""  # 名稱命名
    product1.PartNumber = "op" + str(op_number) + "_A_punch_insert_" + str(X) + str(i)  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("op" + str(op_number) + "_A_punch_insert_" + str(X) + str(i))
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
    strParam8.ValuateFromString("L1 :1- 線割, 單 %%P 0.01")
    product1 = product1.ReferenceProduct
    parameters9 = product1.UserRefProperties
    strParam9 = parameters9.CreateString("A", "")  # 螺栓孔
    strParam9.ValuateFromString("A : (下模螺絲)")
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
    strParam19.ValuateFromString("AP: (A沖沖孔)")
    product1 = product1.ReferenceProduct
    # ====↑設定性質↑=====================================
    part1.Update()
    # --------------↓刪除不需要的Data↓--------------
    selection1 = partDocument1.Selection
    if now_plate_line_number == 1:  # --------模板1
        selection1.Clear()
        selection1.Search(
            "Name=plate_line_2*,all")
        if selection1.Count != 0:
            selection1.Delete()
        selection1.Clear()
        for o in range(1, 1 + (now_op_number - 1)):
            selection1.Clear()
            selection1.Search()
            "Name=*_op" + str(o) + "0_*,all"
            if selection1.Count != 0:
                selection1.Delete()
            selection1.Clear()
        for o in range((now_op_number + 1), 1 + total_op_number):
            selection1.Clear()
            selection1.Search(
                "Name=*_op" + str(o) + "0_*,all")
            if selection1.Count != 0:
                selection1.Delete()
            selection1.Clear()
    if now_plate_line_number == 2:  # --------模板2
        selection1.Clear()
        selection1 = partDocument1.Selection
        selection1.Search("Name=plate_line_1*,all")
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
    partDocument1.SaveAs(
        gvar.save_path + product1.PartNumber + ".CATPart")  # 存檔的檔案名稱
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    partDocument1.Close()
    # 原本for到這邊

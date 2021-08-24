import win32com.client as win32
import global_var as gvar
import defs
import InsertDef
import time
import PunchDef


def SplintInsert(now_plate_line_number,A_punch_H):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    g = now_plate_line_number
    open_name1 = "QR_allotype_insert"
    total_op_number = int(gvar.strip_parameter_list[2])
    (insert_interferance_count) = InsertDef.InsertInterferance(gvar.StripDataList[39], "_allotype_cut_line_",
                                                               open_name1,
                                                               now_plate_line_number)
    for now_op_number in range(1, 1 + total_op_number):
        n = now_op_number
        op_number = 10 * n
        if gvar.StripDataList[39][g][n] > 0:
            # QR_allotype(now_plate_line_number,now_op_number,total_op_number,insert_interferance_count)
            pass  # 未使用
        if gvar.StripDataList[37][g][n] > 0:
            partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            partDocument2 = documents1.Open(gvar.open_path + "QR_Splint.CATPart")
            # 在catapp上切換各視窗
            # ======================================
            defs.window_change(partDocument1, partDocument2)
            # ======================================
            A_punch_QR_Splint_change(now_plate_line_number, now_op_number, A_punch_H)


def A_punch_QR_Splint_change(now_plate_line_number, now_op_number, A_punch_H):
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    # ---------使用迴圈，建立關連↓-------------
    # ---------決定沖頭高度↓-------------
    length = [None] * 5
    # ======================================================================================================
    length[0] = part1.Parameters.Item("plate_down_plane")
    g = now_plate_line_number
    if gvar.Mold_status == "開模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[14]) + 0  # (upper_die_open_height)
    elif gvar.Mold_status == "閉模":
        length[0].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[14])
    # ======================================================================================================
    length[1] = part1.Parameters.Item("plate_height")
    length[1].Value = gvar.strip_parameter_list[14]
    length[2] = part1.Parameters.Item("D")
    length[2].Value = gvar.strip_parameter_list[23]
    length[3] = part1.Parameters.Item("H")
    length[3].Value = A_punch_H  # A punch 的沉頭直徑
    part1.Update()
    # ---------決定沖頭高度↑-------------
    g = now_plate_line_number
    n = now_op_number
    op_number = n * 10
    # for i in range(1, 1 + gvar.StripDataList[37][g][n]):
    i = 1
    part1.Parameters.Item("cut_line_assume_1").OptionalRelation.Modify(
        "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i))  # 草圖置換
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    product1 = partDocument1.getItem("QR_Splint")
    if gvar.StripDataList[37][g][n] > 1:
        xi = gvar.StripDataList[37][g][n]
        PunchDef.A_punch_clash_change_QR_Splint(xi, op_number)
        # i = i + (xi - 1)
    # 數字二位數化,1~10改為01~10
    X = 0  # 名稱命名
    if i >= 10:
        X = ""  # 名稱命名
    product1.PartNumber = "op" + str(op_number) + "_A_punch_QR_Splint_insert_" + str(X) + str(i)  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("op" + str(op_number) + "_A_punch_QR_Splint_insert_" + str(X) + str(i))
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
    strParam19 = parameters21.CreateString("AP", "")  # A沖沖孔7
    strParam19.ValuateFromString("")
    product1 = product1.ReferenceProduct
    # ====↑設定性質↑=====================================
    time.sleep(2)
    partDocument1.SaveAs(gvar.save_path + product1.PartNumber + ".CATPart")  # 存檔的檔案名稱
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    #原本FOR到這邊
    # ---------使用迴圈，建立關連↑-------------
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    partDocument1.Close()

import win32com.client as win32
import global_var as gvar
import defs
import InsertDef
import PunchDef
import time


def StripperInsert(now_plate_line_number, A_punch_H):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    g = now_plate_line_number
    total_op_number = int(gvar.strip_parameter_list[2])
    for now_op_number in range(1, 1 + total_op_number):
        n = now_op_number
        op_number = 10 * n
        # ---------------------------------------------------------------------------------------↓補強入子
        if gvar.StripDataList[4][g][n] > 0:
            # for for_counter = 1, 1 + plate_line_Reinforcement_cut_line(g, n)
            #     stripper_Reinforcement_insert(g, n, for_counter)
            pass  # 未使用
        # ---------------------------------------------------------------------------------------↑補強入子
        # ----------------------------------------------------------------------------------------------↓快拆沖頭
        if gvar.StripDataList[29][g][n] > 0:  # 沖切沖頭_右
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # data_type = "line"
            # data_number = gvar.StripDataList[29][g][n]
            # part_name = "quickly_remove_punch_stripper_insert_right"
            # product_name = "right_quickly_remove_cut_punch"
            # modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_right_quickly_remove_cut_line_"
            # quickly_remove_punch
            # quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[30][g][n] > 0:  # 沖切沖頭_左
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # data_type = "line"
            # data_number =gvar.StripDataList[30][g][n]
            # part_name = "quickly_remove_punch_stripper_insert_left"
            # product_name = "left_quickly_remove_cut_punch"
            # modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_left_quickly_remove_cut_line_"
            # quickly_remove_punch
            # quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[31][g][n] > 0:  # 沖切沖頭_上
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # data_type = "line"
            # data_number = gvar.StripDataList[31][g][n]
            # part_name = "quickly_remove_punch_stripper_insert_up"
            # product_name = "up_quickly_remove_cut_punch"
            # modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_up_quickly_remove_cut_line_"
            # quickly_remove_punch
            # quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[32][g][n] > 0:  # 沖切沖頭_下
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # data_type = "line"
            # data_number = gvar.StripDataList[32][g][n]
            # part_name = "quickly_remove_punch_stripper_insert_down"
            # product_name = "down_quickly_remove_cut_punch"
            # modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_down_quickly_remove_cut_line_"
            # quickly_remove_punch
            # quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[33][g][n] > 0:  # 成形沖頭_右
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # data_type = "surface"
            # data_number = gvar.StripDataList[33][g][n]
            # part_name = "quickly_remove_punch_stripper_insert_right"
            # product_name = "right_quickly_remove_bending_punch"
            # modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_right_quickly_remove_bending_surface_"
            # quickly_remove_punch
            # quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[34][g][n] > 0:  # 成形沖頭_左
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # data_type = "surface"
            # data_number = gvar.StripDataList[34][g][n]
            # part_name = "quickly_remove_punch_stripper_insert_left"
            # product_name = "left_quickly_remove_bending_punch"
            # modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_left_quickly_remove_bending_surface_"
            # quickly_remove_punch
            # quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[35][g][n] > 0:  # 成形沖頭_上
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # data_type = "surface"
            # data_number = gvar.StripDataList[35][g][n]
            # part_name = "quickly_remove_punch_stripper_insert_up"
            # product_name = "up_quickly_remove_bending_punch"
            # modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_up_quickly_remove_bending_surface_"
            # quickly_remove_punch
            # quickly_remove_punch_change
            pass  # 未使用
        if gvar.StripDataList[36][g][n] > 0:  # 成形沖頭_下
            # partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # data_type = "surface"
            # data_number = gvar.StripDataList[36][g][n]
            # part_name = "quickly_remove_punch_stripper_insert_down"
            # product_name = "down_quickly_remove_bending_punch"
            # modify_name = "plate_line_" + str(g) + "_op" + str(op_number) + "_down_quickly_remove_bending_surface_"
            # quickly_remove_punch
            # quickly_remove_punch_change
            pass  # 未使用
        # ----------------------------------------------------------------------------------------------↑快拆沖頭
        if gvar.StripDataList[39][g][n] > 0:  # 快拆異形沖頭
            # QR_allotype
            pass  # 未使用
        if gvar.StripDataList[37][g][n] > 0:  # A沖
            partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            # ===========================================================================(A_punch_QR_Stripper)
            partDocument2 = documents1.Open(gvar.open_path + "QR_Stripper.CATPart")
            defs.window_change(partDocument1, partDocument2)
            # ===========================================================================(A_punch_QR_Stripper)
            A_punch_QR_Stripper_change(now_plate_line_number, now_op_number, A_punch_H)


def A_punch_QR_Stripper_change(now_plate_line_number, now_op_number, A_punch_H):
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
        length[0].Value = float(gvar.strip_parameter_list[1]) + 0  # (upper_die_open_height)
    if gvar.Mold_status == "閉模":
        length[0].Value = float(gvar.strip_parameter_list[1])
    # ======================================================================================================
    # ======================================================================================================
    length[1] = part1.Parameters.Item("plate_height")
    length[1].Value = gvar.strip_parameter_list[20]
    length[2] = part1.Parameters.Item("D")
    length[2].Value = gvar.strip_parameter_list[23]
    length[3] = part1.Parameters.Item("H")
    length[3].Value = A_punch_H  # A punch 的沉頭直徑
    # ---------決定沖頭高度↑-------------
    g = now_plate_line_number
    n = now_op_number
    op_number = n * 10
    # for i in range(1 , 1 + gvar.StripDataList[37][g][n]):
    i = 1
    part1.Parameters.Item("cut_line_assume_1").OptionalRelation.Modify(
        "die\\plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i))  # 草圖置換
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    product1 = partDocument1.getItem("QR_Stripper")
    # 數字二位數化,1~10改為01~10
    X = 0  # 名稱命名
    if i >= 10:
        X = ""  # 名稱命名
    product1.PartNumber = "op" + str(op_number) + "_A_punch_QR_Stripper_insert_" + str(X) + str(i)  # 樹枝圖名稱
    if gvar.StripDataList[37][g][n] > 1:
        xi = gvar.StripDataList[37][g][n]
        PunchDef.A_punch_clash_change_QR_Stripper(xi, op_number)
        # i = i + (xi - 1)
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("op" + str(op_number) + "_A_punch_QR_Stripper_insert_" + str(X) + str(i))
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
    strParam19 = parameters21.CreateString("AP", "")  # A沖沖孔
    strParam19.ValuateFromString("")
    product1 = product1.ReferenceProduct
    # ====↑設定性質↑=====================================
    time.sleep(2)
    partDocument1.SaveAs(gvar.save_path + product1.PartNumber + ".CATPart")  # 存檔的檔案名稱
    gvar.all_part_number = gvar.all_part_number + 1  # 2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber  # 將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    # ---------使用迴圈，建立關連↑-------------
    # 原本FOR到這裡
    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    partDocument1.Close()

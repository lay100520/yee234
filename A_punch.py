import win32com.client as win32
import defs
import global_var as gvar
import time

def APunchMaking(now_plate_line_number):
    g = now_plate_line_number
    total_op_number = int(gvar.strip_parameter_list[2])
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    A_Punch_Module = str('SJAS')
    for    now_op_number in range (1,1+    total_op_number):
        n = now_op_number
        op_number = 10 * n
        # '--------------------------------------------------------------------------------------------A沖
        if  gvar.StripDataList[37][g][n] > 0:
            partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
            time.sleep(1)
            partDocument2 = documents1.Open(gvar.open_path + A_Punch_Module + ".CATPart")
            # '在CATIA上切換各視窗
            # '======================================
            defs.window_change(partDocument1,partDocument2)
            # '======================================
            (A_punch_H)=A_punch_change(now_plate_line_number,now_op_number)
    return A_punch_H

def A_punch_change(now_plate_line_number,now_op_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    #----------------------------------------------------------  A punch 數值修改和抓取
    part1 = partDocument1.Part
    length=[None]*5
    A_Punch_Module = str('SJAS')
    length[0] = part1.Parameters.Item("D")
    length[0].Value = float(gvar.strip_parameter_list[23])
    length[1] = part1.Parameters.Item("L")
    length[1].Value = float(gvar.strip_parameter_list[1]) + float(gvar.strip_parameter_list[20]) + float(
            gvar.strip_parameter_list[17]) + float(gvar.strip_parameter_list[14])+5 #(Punch_eat)
    length[2] = part1.Parameters.Item("H")
    length[3] = part1.Parameters.Item("hight")
    length[3].Value = -5 #(-Punch_eat)
    part1.UpdateObject (part1.Bodies.Item("Body.2"))
    A_punch_D = length[0].Value  #A punch 的直徑
    A_punch_H = length[2].Value  #A punch 的沉頭直徑
    #--------------------------------------------------------
    g = now_plate_line_number
    n = now_op_number
    op_number = n*10
    for i in range( 1 ,1+ gvar.StripDataList[37][g][n]):
        part1.Parameters.Item("cut_line_assume_1").OptionalRelation.Modify ("die\\plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i)) #草圖置換
        part1.UpdateObject (part1.Bodies.Item("Body.2"))
        product1 = partDocument1.getItem(A_Punch_Module)
        #數字二位數化,1~10改為01~10
        X = 0 #名稱命名
        if i >= 10 :
            X = ""  #名稱命名
        product1.PartNumber = "op" + str(op_number) + "_" + A_Punch_Module + "_" + str(length[0].Value) + "_" + str(A_punch_D) + "_" + str(X) + str(i)  #樹枝圖名稱
        #====↓設定性質↓=====================================
        partDocument1 = catapp.ActiveDocument
        product1 = partDocument1.getItem(A_Punch_Module + "_" + str(length[0].Value) + "_" + str(A_punch_D))
        parameters1 = product1.UserRefProperties
        strParam1 = parameters1.CreateString("NO.", "")
        strParam1.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters2 = product1.UserRefProperties
        strParam2 = parameters2.CreateString("Part Name", "")
        strParam2.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters3 = product1.UserRefProperties
        strParam3 = parameters3.CreateString("Size", "")
        strParam3.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters4 = product1.UserRefProperties
        strParam4 = parameters4.CreateString("Material_Data", "")
        strParam4.ValuateFromString (A_Punch_Module)
        product1 = product1.ReferenceProduct
        parameters5 = product1.UserRefProperties
        strParam5 = parameters5.CreateString("Heat Treatment", "")
        strParam5.ValuateFromString (A_Punch_Module)
        product1 = product1.ReferenceProduct
        parameters6 = product1.UserRefProperties
        strParam6 = parameters6.CreateString("Quantity", "")
        strParam6.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters7 = product1.UserRefProperties
        strParam7 = parameters7.CreateString("Page", "")
        strParam7.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters8 = product1.UserRefProperties
        strParam8 = parameters8.CreateString("L1", "") #形狀孔
        strParam8.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters9 = product1.UserRefProperties
        strParam9 = parameters9.CreateString("A", "") #螺栓孔
        strParam9.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters14 = product1.UserRefProperties
        strParam12 = parameters14.CreateString("HP", "") #合銷孔
        strParam12.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters15 = product1.UserRefProperties
        strParam13 = parameters15.CreateString("B", "") #B型引導沖孔
        strParam13.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters16 = product1.UserRefProperties
        strParam14 = parameters16.CreateString("BP", "") #B沖沖孔
        strParam14.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters17 = product1.UserRefProperties
        strParam15 = parameters17.CreateString("TS", "") #浮升引導
        strParam15.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters18 = product1.UserRefProperties
        strParam16 = parameters18.CreateString("IG", "") #內導柱
        strParam16.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters19 = product1.UserRefProperties
        strParam17 = parameters19.CreateString("F", "") #外導柱
        strParam17.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters20 = product1.UserRefProperties
        strParam18 = parameters20.CreateString("CS", "") #等高套筒
        strParam18.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        parameters21 = product1.UserRefProperties
        strParam19 = parameters21.CreateString("AP", "") #A沖沖孔7
        strParam19.ValuateFromString ("")
        product1 = product1.ReferenceProduct
        #====↑設定性質↑=====================================
        partDocument1.SaveAs (gvar.save_path + "op" + str(op_number) + "_" + A_Punch_Module + "_" + str(length[0].Value) + "_" + str(A_punch_D) + "_" + str(X) + str(i) + ".CATPart" )#存檔的檔案名稱
        time.sleep(1)
        gvar.all_part_number = gvar.all_part_number + 1 #2D出圖陣列號碼累加
        gvar.all_part_name[gvar.all_part_number] = product1.PartNumber #將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    #---------使用迴圈，建立關連↑-------------
    part1.UpdateObject (part1.Bodies.Item("Body.2"))
    partDocument1.Close()
    return A_punch_H
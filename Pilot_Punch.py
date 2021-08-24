import win32com.client as win32
import global_var as gvar
import defs
import time

Pilot_Punch_Diameter = gvar.strip_parameter_list[23]
stripper_plate_height = gvar.strip_parameter_list[20]
Thickness = gvar.strip_parameter_list[1]
Pilot_Punch_data = [[0.0] * 3 for i in range(10)]
Pilot_Punch_Material = gvar.strip_parameter_list[24]
Pilot_Punch_Heat_treatment = gvar.strip_parameter_list[25]


def Pilot_Punch():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.open_path + "Pilot_Punch.CATPart")
    part1 = partDocument1.Part
    length = [None] * 20
    formula = [None] * 20
    parameter = [None] * 20
    # ======================================================================================================
    length[1] = part1.Parameters.Item("Pilot_Punch_D")  # 引導沖直徑
    length[1].Value = float(Pilot_Punch_Diameter)  # + Val(Pilot_Punch_Diameter)
    # ======================================================================================================
    # ======================================================================================================
    length[2] = part1.Parameters.Item("straight_L")  # 引導沖孔長度
    length[2].Value = float(stripper_plate_height) + 1.5 * float(Thickness)
    # ======================================================================================================
    # ======================================================================================================
    length[3] = part1.Parameters.Item("sink_D")  # 沉頭孔直徑
    length[3].Value = int(float(Pilot_Punch_Diameter) + 3)
    # ======================================================================================================
    Pilot_Punch_data[1][1] = length[1].Value  # 引導沖頭直徑 D
    Pilot_Punch_data[2][1] = float(stripper_plate_height)  # 引導沖頭孔深度 L
    Pilot_Punch_data[3][1] = length[2].Value + 10  # 引導沖頭總長度 L#
    Pilot_Punch_data[4][1] = length[3].Value
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1.Update()
    product1.PartNumber = "Pilot_Punch_" + str(Pilot_Punch_data[1][1]) + "Dx" + str(
        Pilot_Punch_data[3][1]) + "L"  # 樹枝圖名稱
    # ====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Pilot_Punch")
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
    strParam4.ValuateFromString(Pilot_Punch_Material)
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam5 = parameters5.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(Pilot_Punch_Heat_treatment)
    product1 = product1.ReferenceProduct
    parameters6 = product1.UserRefProperties
    strParam6 = parameters6.CreateString("Quantity", "")
    strParam6.ValuateFromString("")
    product1 = product1.ReferenceProduct
    parameters7 = product1.UserRefProperties
    strParam7 = parameters7.CreateString("Page", "")
    strParam7.ValuateFromString("")
    product1 = product1.ReferenceProduct
    # ====↑設定性質↑=====================================
    part1.Update()
    partDocument1.SaveAs(gvar.save_path + "Pilot_Punch_" + str(Pilot_Punch_data[1][1]) + "Dx" + str(
        Pilot_Punch_data[3][1]) + "L.CATPart")  # 存檔的檔案名稱
    # partDocument1.SaveAs save_path + "pad_lower_0" + i + ".CATPart" #存檔的檔案名稱
    # all_part_number = all_part_number + 1 #2D出圖陣列號碼累加
    # all_part_name(all_part_number) = product1.PartNumber #將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    part1.Update()
    partDocument1.Close()
    return Pilot_Punch_data
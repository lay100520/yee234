import win32com.client as win32
import global_var as gvar
import defs
import time

def UpperDieSet(lower_die_set_length,lower_die_set_width):
    total_op_number = int(gvar.strip_parameter_list[2])
    out_Guide_Material = "MYJP"
    out_Guide_Diameter = 32
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    partDocument2 = documents1.Open(gvar.open_path + "Data1.CATPart")
    #======================================
    defs.window_change(partDocument1,partDocument2)    #在CATIA上切換各視窗
    #======================================    # upper_die_set_change.CATMain
    if gvar.die_type == "module" :
        common_plate_size = 40
    else:
        common_plate_size = 0
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    selection1 = partDocument1.Selection
    if gvar.die_type == "module" :
        selection1.Search ("Name=Hole.*,all")
        selection1.Delete()
        selection1.Search ("Name=Pillar_center_*__hole_*,all")
        selection1.Delete()
    length=[None]*21
    formula=[None]*21
    parameter=[None]*21
    #======================================================================================================
    length[1] = part1.Parameters.Item("plate_height")
    g = 1
    if gvar.strip_parameter_list[5] != " " :
        length[1].Value = gvar.strip_parameter_list[5]
    else:
        die_rule_file_name = "模板厚度選擇"
        excel_Sheet_name = "200T以下"
        Column_string_serch = "上模座"
        Row_string_serch = "精密級"
        (search_result)=defs.ExcelSearch(die_rule_file_name,excel_Sheet_name,Row_string_serch,Column_string_serch)
        length[1].Value = search_result
        gvar.strip_parameter_list[5] = search_result
    #======================================================================================================
    length[4] = part1.Parameters.Item("plate_up_plane")
    if gvar.die_type == "module" :
        float(gvar.strip_parameter_list[1])+float(gvar.strip_parameter_list[20])+float(gvar.strip_parameter_list[14])+float(gvar.strip_parameter_list[11])
        plate_position = float(gvar.strip_parameter_list[1])+float(gvar.strip_parameter_list[20])+float(gvar.strip_parameter_list[14])+float(gvar.strip_parameter_list[11]) + 0+0+0#(back_splint_height + common_seat_height+ back_stripper_plate_height)
    else:
        plate_position = float(gvar.strip_parameter_list[1])+float(gvar.strip_parameter_list[20])+float(gvar.strip_parameter_list[14])+float(gvar.strip_parameter_list[11]) + float(gvar.strip_parameter_list[17])
    if gvar.Mold_status == "開模" :
        length[4].Value = plate_position + 28#(upper_die_open_height)
    elif gvar.Mold_status == "閉模" :
        length[4].Value = plate_position
    #======================================================================================================
    length[5] = part1.Parameters.Item("center_point_distance_X")
    length[5].Value = 0
    #======================================================================================================
    length[6] = part1.Parameters.Item("outer_guiding_post_diameter") #外導柱直徑
    length[6].Value = out_Guide_Diameter
    #======================================================================================================
    if float(gvar.strip_parameter_list[8]) != 0 :
    #======================================================================================================
        length[7] = part1.Parameters.Item("U_spacing") #U溝間距
        length[7].Value = float(gvar.strip_parameter_list[8])
        length[8] = part1.Parameters.Item("U_depth") #U溝深度
        length[8].Value = float(gvar.strip_parameter_list[9])
        length[9] = part1.Parameters.Item("U_width") #U溝寬度
        length[9].Value = float(gvar.strip_parameter_list[10])
    #======================================================================================================
    parameters1 = part1.Parameters
    strParam20 = parameters1.Item("outer_guiding_post_bolt")     #外導柱固定螺栓直徑
    length[10] = part1.Parameters.Item("outer_guiding_post_pin") #外導柱固定合銷直徑
    if out_Guide_Material == "MYJP":
        if out_Guide_Diameter == 20:
            strParam20.Value = "M6"
            length[10].Value = 6
        elif out_Guide_Diameter == 25:
            strParam20.Value = "M8"
            length[10].Value = 8
        elif out_Guide_Diameter == 32:
            strParam20.Value = "M10"
            length[10].Value = 8
        elif out_Guide_Diameter == 38:
            strParam20.Value = "M10"
            length[10].Value = 10
        elif out_Guide_Diameter == 50:
            strParam20.Value = "M12"
            length[10].Value = 10
    #----------------------------------------------------------
    if out_Guide_Material == "MYKP":
        if out_Guide_Diameter == 20 :
            strParam20.Value = "M8"
            length[10].Value = 8
        elif out_Guide_Diameter == 25:
            strParam20.Value = "M8"
            length[10].Value = 8
        elif out_Guide_Diameter == 32 :
            strParam20.Value = "M10"
            length[10].Value = 8
        elif out_Guide_Diameter == 38 :
            strParam20.Value = "M10"
            length[10].Value = 10
        elif out_Guide_Diameter == 50  :
            strParam20.Value = "M12"
            length[10].Value = 10
    #----------------------------------------------------------
    if out_Guide_Material == "DANLY":
        if out_Guide_Diameter == 25 :
            strParam20.Value = "M6"
        elif out_Guide_Diameter == 32  :
            strParam20.Value = "M8"
        elif out_Guide_Diameter == 40 :
            strParam20.Value = "M8"
        elif out_Guide_Diameter == 50  :
            strParam20.Value = "M8"
        elif out_Guide_Diameter == 63 :
            strParam20.Value = "M8"
        elif out_Guide_Diameter == 80   :
            strParam20.Value = "M8"
        if  out_Guide_Diameter > 32 :
            # Call upper_die_set_DANLY_MODEL_1
            pass  # 未使用                  #螺栓孔形式變更
        elif  out_Guide_Diameter <= 32 :
            # Call upper_die_set_DANLY_MODEL_2
            pass  # 未使用                  #螺栓孔形式變更
    #======================================================================================================
    part1.Parameters.Item("1_formula_1").OptionalRelation.Modify ("die\\upper_die_seat_line") #草圖置換
    part1.Update()
    specsAndGeomWindow1 = catapp.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    part1.Update()
    #======================================================================================================
    length[11] = part1.Parameters.Item("lower_up_die_set_X")
    upper_die_set_length = lower_die_set_length + length[11].Value
    #======================================================================================================
    length[12] = part1.Parameters.Item("lower_up_die_set_Y")
    upper_die_set_width = lower_die_set_width + length[12].Value
    #======================================================================================================
    if gvar.die_type == "module" :
        #-------------------------------------------------------------------孔座標(X,Y)
        guild_position_X=[0.0]*31
        guild_position_Y=[0.0]*31
        side_direction_X = 35
        side_direction_Y = 60
        guild_position_X[1] = side_direction_X
        guild_position_X[2] = side_direction_X
        guild_position_X[3] = upper_die_set_length - side_direction_X
        guild_position_X[4] = upper_die_set_length - side_direction_X
        guild_position_Y[1] = side_direction_Y
        guild_position_Y[2] = upper_die_set_width - side_direction_Y
        guild_position_Y[3] = upper_die_set_width - side_direction_Y
        guild_position_Y[4] = side_direction_Y
        #-------------------------------------------------------------------孔座標(X,Y)
        part_file_name = "upper_die_set"
        body_name_1 = "PartBody"
        hybridBody_name = "die"
        (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(part_file_name, body_name_1,
                                                                                                 hybridBody_name)  # 環境設定(檔案名,body名,依據群組名)(全域變數改)
        hybridShape1 = ElementHybridBody.HybridShapes.Item("upper_die_seat_line") #宣告元素(當下模板的曲線)
        (element_Reference1)=defs.ExtremumPoint("X_min", "Y_min", "Z_max", 2, hybridShape1,ElementDocument,ElementBody,ElementHybridBody)   #建立極值點   (極值方向1,極值方向2,極值方向3,需要幾個方向,在哪個元素上的極值)  element_Reference(1)為OUT
        element_Reference2= element_Reference1
        element_Reference2.Name = "upper_die_seat_plate_min" #建立最小點
        hybridShape2 = ElementBody.HybridShapes.Item("down_plane")  #宣告平面
        element_Reference12 = hybridShape2
        element_Reference10 = element_Reference1
        for L_Hole_N in range (1 ,5):
        #----------------------------------------------------------------------------------------------------------------guild_hole_1
            (element_point5)=defs.BuildPoint(element_Reference10,ElementDocument,ElementBody,ElementHybridBody,SketchPosition="Hybridbody")   #建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
            element_point5.X.Value = guild_position_X[L_Hole_N]
            element_point5.Y.Value = guild_position_Y[L_Hole_N]
            element_point5.Name = "guild_point_1_" + str(L_Hole_N)
            element_Reference11 = element_point5
            (hole_pin)=defs.HoleSimpleD(45, 32,0,ElementDocument,ElementBody,element_Reference11,element_Reference12) #直孔  (M,upper_die_set_height,out孔陳述式,方向) element_Reference(11)=point #element_Reference(12)=plane direction=>0=下 1=上
            hole_pin.Name = "upper_guild_Hole_1_" + str(L_Hole_N)
        #----------------------------------------------------------------------------------------------------------------guild_hole_1
    #------------------------------------------------------------------------------------------------------模組型_導柱孔
    product1.PartNumber = "upper_die_set" #樹枝圖名稱
    #====↓設定性質↓=====================================
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.getItem("upper_die_set")
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
    strParam4.ValuateFromString (gvar.strip_parameter_list[6])
    product1 = product1.ReferenceProduct
    parameters5 = product1.UserRefProperties
    strParam5 = parameters5.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString (gvar.strip_parameter_list[7])
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
    strParam19 = parameters21.CreateString("AP", "") #A沖沖孔
    strParam19.ValuateFromString ("")
    product1 = product1.ReferenceProduct
    #====↑設定性質↑=====================================
    part1.Update()
    part1.Update()
    partDocument1.SaveAs (gvar.save_path + "upper_die_set.CATPart") #存檔的檔案名稱
    gvar.all_part_number = gvar.all_part_number + 1 #2D出圖陣列號碼累加
    gvar.all_part_name[gvar.all_part_number] = product1.PartNumber #將樹枝圖名稱存至陣列,以利後續出圖時直接引用此陣列即可
    partDocument1.Close()
    #======================================    # upper_die_set_change.CATMain
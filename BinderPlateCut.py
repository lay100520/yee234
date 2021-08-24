import win32com.client as win32
import defs
import global_var as gvar
import time
def BinderPlateCut(now_plate_line_number):
    g = now_plate_line_number
    total_op_number = int(gvar.strip_parameter_list[2])
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    g = now_plate_line_number
    for now_op_number in range( 1, 1+ total_op_number):
        n = now_op_number
        op_number = 10 * n
        if gvar.StripDataList[38][g][n] > 0:
            for now_data_number in range( 1, 1+ gvar.StripDataList[38][g][n]):
                interferance_pad_name = "_cut_punch_"
                interferance_line_name = "_cut_line_"
                open_file_name = "cut_punch"
                # --------------------------------------------------------------------------------------------一個沖頭
                if gvar.StripDataList[38][g][n] == 1:
                    if op_number == 30 or op_number == 50:
                        S_Binder_Plate_other_cut(op_number)
                        partDocument1 = documents1.Open(gvar.save_path + "op" + str(op_number) + interferance_pad_name + "0" + str(now_data_number) + ".CATPart")
                        Punch_input(op_number)
                        Boolean_Remove()  # 進行布林移除挖槽
                        partDocument1.save()
                        partDocument1.Close()
                # --------------------------------------------------------------------------------------------四個沖頭
                if gvar.StripDataList[38][g][n] == 4:
                    if op_number == 40:
                        if now_data_number == 1:
                            S_Binder_Plate_cut(op_number)
                        partDocument1 = documents1.Open(
                        gvar.save_path + "op" + str(op_number) + interferance_pad_name + "0" + str(now_data_number) + ".CATPart")
                        Punch_input(op_number)
                        Boolean_Remove()  # 進行布林移除挖槽
                        partDocument1.save()
                        partDocument1.Close()
        partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
        # --------------↓刪除不需要的Data↓--------------
        selection1 = partDocument1.Selection
        selection2 = partDocument1.Selection
        selection3 = partDocument1.Selection
        selection4 = partDocument1.Selection
        selection5 = partDocument1.Selection
        selection6 = partDocument1.Selection
        selection7 = partDocument1.Selection
        if op_number == 30 or op_number == 40 or op_number == 50:
            selection1.Clear()
            selection1.Search(    "Name=Data_*,all")
            if selection1.Count > 0:
                selection1.Delete()
        # --------------↑刪除不需要的Data↑--------------

def S_Binder_Plate_other_cut(op_number):
    if op_number == 30 :
        creat_reference()
    #--------------------------------------------------
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    #---------------------------------------------------------------------------開起Data
    part1 = partDocument1.Part
    relations1 = part1.Relations
    parameters1 = part1.Parameters                   #參數指令起手宣告
    #-----------------------------------------------------------------------------建立測量參數
    length1 = parameters1.CreateDimension("", "LENGTH", 0)    #build parameter
    length2 = parameters1.CreateDimension("", "LENGTH", 0)
    length3 = parameters1.CreateDimension("", "LENGTH", 0)
    length4 = parameters1.CreateDimension("", "LENGTH", 0)
    length5 = parameters1.CreateDimension("", "LENGTH", 0)
    length1.rename ("Data_Binder_Plate_other_outside")
    length2.rename ("Data_Binder_Plate_other_thickness")
    length3.rename ("Data_Binder_Plate_other_distance")
    length4.rename ("Data_Binder_Plate_other_distance_H")
    length5.rename ("Data_Binder_Plate_other_height")
    #-----------------------------------------------------------------------------建立測量參數
    #-----------------------------------------------------------------------------開始測量
    formula1 = relations1.Createformula("Data_Binder_Plate_other_outside", "", length1, "length()")
    formula2 = relations1.Createformula("Data_Binder_Plate_other_thickness", "", length2, "length()")
    formula3 = relations1.Createformula("Data_Binder_Plate_other_distance", "", length3, "length()")
    formula4 = relations1.Createformula("Data_Binder_Plate_other_distance_H", "", length4, "length()")
    formula5 = relations1.Createformula("Data_Binder_Plate_other_height", "", length5, "length()")
    #-----------------------------------------------------------------------------鍵槽
    if op_number == 30 :
        S_point_distance_parameter = "die\\open_curve_1_1_A"
        E_point_distance_parameter = "die\\Extremum.5"
        formula1.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject (formula1) #單步更新 formula
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\plate_centor_point"
        formula3.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject (formula3) #單步更新 formula
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\Extremum.5"
        formula4.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject (formula4) #單步更新 formula
    #-----------------------------------------------------------------------------中心軸
    elif op_number == 50 :
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\open_curve_1_1_A"
        formula1.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject (formula1) #單步更新 formula
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\plate_centor_point"
        formula3.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject (formula3) #單步更新 formula
        S_point_distance_parameter = "die\\zero_original_point"
        E_point_distance_parameter = "die\\open_curve_1_1_A"
        formula4.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
        part1.UpdateObject (formula4) #單步更新 formula
    #-----------------------------------------------------------------------------開始測量
    ##--------------------------------------------------------------讀取數據
    Data_Binder_Plate_other_outside = part1.Parameters.Item("Data_Binder_Plate_other_outside")
    Data_Binder_Plate_other_outside = Data_Binder_Plate_other_outside.Value
    Data_Binder_Plate_other_thickness = part1.Parameters.Item("Data_Binder_Plate_other_thickness")
    Data_Binder_Plate_other_thickness = Data_Binder_Plate_other_thickness.Value
    Data_Binder_Plate_other_distance = part1.Parameters.Item("Data_Binder_Plate_other_distance")
    Data_Binder_Plate_other_distance = Data_Binder_Plate_other_distance.Value
    Data_Binder_Plate_other_distance_H = part1.Parameters.Item("Data_Binder_Plate_other_distance_H")
    Data_Binder_Plate_other_distance_H = Data_Binder_Plate_other_distance_H.Value
    Data_Binder_Plate_other_height = part1.Parameters.Item("Data_Binder_Plate_other_height")
    Data_Binder_Plate_other_height = Data_Binder_Plate_other_height.Value
    #壓板標準件尺寸判斷↓
    if op_number == 30 or op_number == 50 :
        if op_number == 30 :
            Data_Binder_Plate_other_outside = Data_Binder_Plate_other_outside / 2
        if Data_Binder_Plate_other_outside < 5 :
            Data_Binder_Plate_other_thickness = 6
        elif Data_Binder_Plate_other_outside >= 5 and Data_Binder_Plate_other_outside < 6 :
            Data_Binder_Plate_other_outside = 5
            Data_Binder_Plate_other_thickness = 6
        elif Data_Binder_Plate_other_outside >= 6 and Data_Binder_Plate_other_outside < 7 :
            Data_Binder_Plate_other_outside = 6
            Data_Binder_Plate_other_thickness = 6
        elif Data_Binder_Plate_other_outside >= 7 :
            Data_Binder_Plate_other_outside = 7
            Data_Binder_Plate_other_thickness = 6
    #壓板標準件尺寸判斷↑
    partDocument1.save()
    partDocument1.Close() #關閉檔案
    #==================================================================================================製作挖槽壓板
    partDocument1 = documents1.Open(gvar.open_path + "Binder_Plate_other_cut.CATPart")
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    #------------------------------------------------設定條件
    if op_number == 30 :
    #-----------------------------------------------------------------------------鍵槽
        Binder_Plate_other_outside = part1.Parameters.Item("Binder_Plate_other_outside")
        Binder_Plate_other_outside.Value = Data_Binder_Plate_other_outside
        Binder_Plate_other_thickness = part1.Parameters.Item("Binder_Plate_other_thickness")
        Binder_Plate_other_thickness.Value = Data_Binder_Plate_other_thickness
        Binder_Plate_other_distance = part1.Parameters.Item("Binder_Plate_other_distance")
        Binder_Plate_other_distance.Value = Data_Binder_Plate_other_distance - float(gvar.strip_parameter_list[4])
        Binder_Plate_other_distance_H = part1.Parameters.Item("Binder_Plate_other_distance_H")
        Binder_Plate_other_distance_H.Value = Data_Binder_Plate_other_distance_H + 1
        #壓板高度
        Binder_Plate_other_height = part1.Parameters.Item("Binder_Plate_other_height")
        Binder_Plate_other_height.Value = float(gvar.strip_parameter_list[17])+float(gvar.strip_parameter_list[20])+float(gvar.strip_parameter_list[1])
    #-----------------------------------------------------------------------------中心軸
    elif op_number == 50 :
        Binder_Plate_other_outside = part1.Parameters.Item("Binder_Plate_other_outside")
        Binder_Plate_other_outside.Value = Data_Binder_Plate_other_outside
        Binder_Plate_other_thickness = part1.Parameters.Item("Binder_Plate_other_thickness")
        Binder_Plate_other_thickness.Value = Data_Binder_Plate_other_thickness
        Binder_Plate_other_distance = part1.Parameters.Item("Binder_Plate_other_distance")
        Binder_Plate_other_distance.Value = Data_Binder_Plate_other_distance + float(gvar.strip_parameter_list[4])
        Binder_Plate_other_distance_H = part1.Parameters.Item("Binder_Plate_other_distance_H")
        Binder_Plate_other_distance_H.Value = Data_Binder_Plate_other_distance_H + Data_Binder_Plate_other_outside - 1
        #壓板高度
        Binder_Plate_other_height = part1.Parameters.Item("Binder_Plate_other_height")
        Binder_Plate_other_height.Value = float(gvar.strip_parameter_list[17])+float(gvar.strip_parameter_list[20])+float(gvar.strip_parameter_list[1])
    part1.Update()
    #---------------------------------------------------存檔
    partDocument1.SaveAs( gvar.open_path + "Temporary\\Binder_Plate_Temporary_" + str(op_number) + ".CATPart" )#存檔的檔案名稱
    partDocument1.Close() #關閉檔案

def creat_reference():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    hybridBodies1 = part1.HybridBodies
    hybridBody1 = hybridBodies1.Item("die")
    hybridBodies1 = hybridBody1.Parent
    bSTR1 = hybridBody1.Name
    selection1.Add (hybridBody1)
    visPropertySet1 = visPropertySet1.Parent
    bSTR2 = visPropertySet1.Name
    bSTR3 = visPropertySet1.Name
    visPropertySet1.SetShow (0)
    selection1.Clear()
    hybridShapeFactory1 = part1.HybridShapeFactory
    hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
    parameters1 = part1.Parameters
    hybridShapeCurveExplicit1 = parameters1.Item("finish_open_curve_1_1")
    reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
    hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 0)
    hybridBody1.AppendHybridShape (hybridShapeExtremum1)
    part1.InWorkObject = hybridShapeExtremum1
    part1.Update()
    partDocument1.save()


def Punch_input(op_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    time.sleep(2)
    partDocument2 = documents1.Open(gvar.open_path + "Temporary\\Binder_Plate_Temporary_" + str(op_number) + ".CATPart")
    #======================================
    defs.window_change(partDocument1,partDocument2)   #在CATIA上切換各視窗
    #======================================

def Boolean_Remove():
    #進行布林移除挖槽
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    shapeFactory1 = part1.ShapeFactory
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    body2 = bodies1.Item("Body.3")
    #--------------------------------------------------開始布林移除
    part1.InWorkObject = body1
    shapeFactory1.AddNewRemove (body2)
    part1.Update()

def S_Binder_Plate_cut(op_number):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    #---------------------------------------------------------------------------開啟Data
    part1 = partDocument1.Part
    relations1 = part1.Relations
    parameters1 = part1.Parameters                   #參數指令起手宣告
    #-----------------------------------------------------------------------------建立測量參數
    length1 = parameters1.CreateDimension("", "LENGTH", 0)    #build parameter
    length2 = parameters1.CreateDimension("", "LENGTH", 0)
    length3 = parameters1.CreateDimension("", "LENGTH", 0)
    length4 = parameters1.CreateDimension("", "LENGTH", 0)
    length5 = parameters1.CreateDimension("", "LENGTH", 0)
    length6 = parameters1.CreateDimension("", "LENGTH", 0)
    length7 = parameters1.CreateDimension("", "LENGTH", 0)
    length1.rename ("Data_Binder_Plate_outside")
    length2.rename ("Data_Binder_Plate_length")
    length3.rename ("Data_Binder_Plate_wide")
    length4.rename ("Data_Binder_Plate_thickness")
    length5.rename ("Data_Binder_Plate_number")
    length6.rename ("Data_Binder_Plate_distance")
    length7.rename ("Data_Binder_Plate_height")
    #-----------------------------------------------------------------------------建立測量參數
    #-----------------------------------------------------------------------------開始測量
    formula1 = relations1.CreateFormula("Data_Binder_Plate_outside", "", length1, "length()")   #中心外徑
    formula2 = relations1.CreateFormula("Data_Binder_Plate_length", "", length2, "length()")        #下料距離
    formula3 = relations1.CreateFormula("Data_Binder_Plate_wide", "", length3, "length()")      #靴齒部間隙
    formula4 = relations1.CreateFormula("Data_Binder_Plate_thickness", "", length4, "length()") #壓板厚度
    formula5 = relations1.CreateFormula("Data_Binder_Plate_number", "", length5, "length()") #壓板數量
    formula6 = relations1.CreateFormula("Data_Binder_Plate_distance", "", length6, "length()") #工站間距
    formula7 = relations1.CreateFormula("Data_Binder_Plate_height", "", length7, "length()") #壓板高度
    #中心外徑
    S_point_distance_parameter = "die\\zero_original_point"
    E_point_distance_parameter = "die\\open_curve_center_point_1"
    formula1.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
    part1.UpdateObject (formula1) #單步更新 formula
    #下料距離
    S_point_distance_parameter = "die\\zero_original_point"
    E_point_distance_parameter = "die\\Contour_circle_line_Ymax"
    formula2.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
    part1.UpdateObject (formula2) #單步更新 formula
    #靴齒部間隙
    S_point_distance_parameter = "die\\open_curve_2_1"
    E_point_distance_parameter = "die\\open_curve_2_2"
    formula3.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
    part1.UpdateObject (formula3) #單步更新 formula
    #工站間距
    S_point_distance_parameter = "die\\zero_original_point"
    E_point_distance_parameter = "die\\plate_centor_point"
    formula6.Modify ("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
    part1.UpdateObject (formula6) #單步更新 formula
    #-----------------------------------------------------------------------------開始測量
    ##--------------------------------------------------------------讀取數據
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
    partDocument1.save()
    partDocument1.Close() #關閉檔案
    #==================================================================================================製作挖槽壓板
    partDocument1 = documents1.Open(gvar.open_path + "Binder_Plate_40.CATPart")
    product1 = partDocument1.getItem("Part1")
    part1 = partDocument1.Part
    #------------------------------------------------設定條件
    #外圍圈半徑=(中心外徑+1)
    Binder_Plate_outside = part1.Parameters.Item("Binder_Plate_outside")
    Data_Binder_Plate_outside = int(Data_Binder_Plate_outside) + 1
    Binder_Plate_outside.Value = Data_Binder_Plate_outside + 1
    #固定條長度(中心到下料)=(外圈圓半徑-1)
    Binder_Plate_length = part1.Parameters.Item("Binder_Plate_length")
    Data_Binder_Plate_length = int(Data_Binder_Plate_length) + 1
    Binder_Plate_length.Value = Data_Binder_Plate_length - 1
    #固定條寬度=靴齒部到靴齒部距離
    Binder_Plate_wide = part1.Parameters.Item("Binder_Plate_wide")
    Binder_Plate_wide.Value = Data_Binder_Plate_wide + 2
    #固定片厚度
    Binder_Plate_thickness = part1.Parameters.Item("Binder_Plate_thickness")
    Binder_Plate_thickness.Value = 3.5
    #固定條數量
    Binder_Plate_number = part1.Parameters.Item("Binder_Plate_number")
    Binder_Plate_number.Value = 4
    #工站間距
    Binder_Plate_distance = part1.Parameters.Item("Binder_Plate_distance")
    Binder_Plate_distance.Value = Data_Binder_Plate_distance
    #壓板高度
    Binder_Plate_height = part1.Parameters.Item("Binder_Plate_height")
    Binder_Plate_height.Value = float(gvar.strip_parameter_list[17])+float(gvar.strip_parameter_list[20])+float(gvar.strip_parameter_list[1])
    part1.Update()
    #---------------------------------------------------存檔
    partDocument1.SaveAs(gvar.open_path + "Temporary\\Binder_Plate_Temporary_" + str(op_number) + ".CATPart") #存檔的檔案名稱
    partDocument1.Close() #關閉檔案


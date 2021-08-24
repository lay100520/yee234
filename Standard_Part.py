import win32com.client as win32
import global_var as gvar
import os
import defs
import time
import sys

outer_Guiding_data = [[0] * 9 for i in range(9)]
stripper_pin_data = [[''] * 3 for i in range(3)]
plate_bolt = ""
plate_pin = ""
outer_Guiding_bolt = ""
outer_Guiding_pin = ""
stripper_pin_stop_bolt = [0.0] * 3
stripper_pin_spring = [0] * 3
Pilot_Pin_Material = "MSTH"
Inner_Guiding_Post_Material = "SGPH"
Inner_Guiding_Post_Diameter = 20
Under_Inner_Guiding_Post_Material = "SGFZ"
SBT_Diameter = 16
SBT_Length = 60
out_Guide_Material = "MYJP"
out_Guide_Diameter = 32
out_Guide_Length = 90
SBT_Material = "SBT"
Pilot_Punch_Stripper_punch_Diameter = 2
Pilot_Punch_Stripper_punch_Length = 10
Pilot_Punch_Stripper_punch_Material = "LP"


def Standard_Part(Bolt_data, CB_data, Pin_data, Inner_Guiding_data, SBT_data, SBT_CB_data):
    catapp = win32.Dispatch('CATIA.Application')
    JUMP = False
    # ========================================================================================================================================↓模板鎖固螺栓改變長度參數
    for i in range(1, 1 + 3):
        b = Bolt_data[3][i] - Bolt_data[2][i]  # 螺栓孔總長-螺栓沉頭孔深
        # GoTo WWW
        # #------------------------判斷螺栓
        # ##MM1
        # if length3.Value < 15 :
        # MM1 = 5
        # End if
        # if length3.Value < 22 and length3.Value > 15 :
        # MM1 = 6
        # End if
        # if length3.Value < 27 and length3.Value > 22 :
        # MM1 = 8
        # End if
        # if length3.Value < 35 and length3.Value > 27 :
        # MM1 = 10
        # End if
        # if length3.Value > 35 :
        # MM1 = 12
        # End if
        # #-----------------------判斷數量
        # #form7.Text3.Text = MM1
        # WWW:
        a = str()
        if Bolt_data[1][i] == 25:
            a = "24"
        elif Bolt_data[1][i] == 21:
            a = "20"
        elif Bolt_data[1][i] == 17:
            a = "16"
        elif Bolt_data[1][i] == 13:
            a = "12"
        elif Bolt_data[1][i] == 11:
            a = "10"
        elif Bolt_data[1][i] == 9:
            a = "8"
        elif Bolt_data[1][i] == 7:
            a = "6"
        elif Bolt_data[1][i] == 5.5:
            a = "5"
        elif Bolt_data[1][i] == 4.5:
            a = "4"
        elif Bolt_data[1][i] == 3.5:
            a = "3"
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.standard_path + "\\Bolt\\CB_" + a + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        # --------------------------------------------------------------------↓螺栓長度
        strParam1 = parameters1.Item("CB_M_L")
        iSize = strParam1.GetEnumerateValuesSize()  # STRING 裡面選擇數量
        myArray = [''] * (iSize)
        # strParam1.GetEnumerateValues(myArray)  # 抓取 STRING 的數值放入矩陣之中
        for ii in range(0, iSize):
            excelname = 'StandardPart'
            (BoltValue) = defs.ExcelSearch(excelname, "CB", (ii + 1), 'CB_' + a)
            myArray[ii] = BoltValue
        plate_height = [0] * 4
        plate_height_1 = [0] * 4
        plate_height_2 = [0] * 4
        if i == 1:
            b = int((2 * int(a)) + float(gvar.strip_parameter_list[11]) + float(gvar.strip_parameter_list[14]) -
                    Bolt_data[2][i])
        elif i == 2:
            b = int((2 * int(a)) + float(gvar.strip_parameter_list[17]) - Bolt_data[2][i])
        elif i == 3:
            b = int((2 * int(a)) + float(gvar.strip_parameter_list[29]) + float(gvar.strip_parameter_list[26]) -
                    Bolt_data[2][i])
        while JUMP == False:
            for j in range(0, iSize):
                plate_bolt = a + "-" + str(b)
                if myArray[j] == plate_bolt:
                    JUMP = True
                    break
            b -= 1
            if b < 0:
                print('Error-111')
                sys.exit()
        b += 1
        JUMP = False
        strParam1.Value = plate_bolt
        CB_data[1][i] = "CB_" + plate_bolt
        part1.Update()
        # --------------------------------------------------------------------↑螺栓長度
        # -------------判斷螺栓長度-------------
        # b = -int(-b)
        # n = Right(b, 1)
        # if n <= 5 :
        # b = b - n
        # Elseif n > 5 :
        # b = b + (10 - n)
        # End if
        # CB_data[1][i] = b #螺栓長度
        # -------------判斷螺栓長度-------------
        # -------------------------------------------搜尋現有螺栓尺寸
        file_name = os.listdir(gvar.save_path)
        if "Bolt_CB_" + plate_bolt + ".CATPart" not in file_name:
            partDocument1 = catapp.ActiveDocument
            product1 = partDocument1.getItem("CB_" + a)
            product1.PartNumber = "Bolt_CB_" + plate_bolt  # 樹狀圖名稱
            partDocument1.SaveAs(gvar.save_path + "Bolt_CB_" + plate_bolt + ".CATPart")  # 存檔名稱/存檔
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
    # -------------------------------------------搜尋現有螺栓尺寸
    # ----↓將所有零件名稱存至文字檔(.txt)↓----
    # all_part_number 為儲存所有output零件的總數量,ex:零件有170個,就存"170"這個值
    # all_part_name(i) 為儲存所有output的零件名稱
    # Open save_path + "all_output_part_name.txt" for Output As #1 #記事本儲存路徑
    # Print #1, Trim$(all_part_number) #記事本第一行存all_part_number的總數量
    # for i = 1 , 1+ all_part_number
    # Print #1, all_part_name(i) #第一行之後依序存零件名稱
    # Next
    # Close() #1
    # -----↑將所有零件名稱存至文字檔(.txt)↑----
    # ========================================================================================================================================↑模板鎖固螺栓改變長度參數
    # ========================================================================================================================================↓模板定位Pin改變長度參數
    for i in range(1, 4):
        a = Pin_data[1][i]  # -------------------Pin直徑
        b = int(Pin_data[2][i])  # -------------------Pin長
        if Pin_data[1][i] == 5:
            if i == 2:
                a = "LP" + "_5"
            else:
                a = Pilot_Pin_Material + "_5"
        elif Pin_data[1][i] == 6:
            if i == 2:
                a = "LP" + "_6"
            else:
                a = Pilot_Pin_Material + "_6"
        elif Pin_data[1][i] == 8:
            if i == 2:
                a = "LP" + "_8"
            else:
                a = Pilot_Pin_Material + "_8"
        elif Pin_data[1][i] == 10:
            if i == 2:
                a = "LP_10"
            else:
                a = Pilot_Pin_Material + "_10"
        elif Pin_data[1][i] == 12:
            if i == 2:
                a = "LP" + "_12"
            else:
                a = Pilot_Pin_Material + "_12"
        elif Pin_data[1][i] == 13:
            if i == 2:
                a = "LP" + "_13"
            else:
                a = Pilot_Pin_Material + "_13"
        documents1 = catapp.Documents
        if i == 2:
            partDocument1 = documents1.Open(gvar.standard_path + "\\Stripper_pin\\" + a + ".CATPart")
            c = "Stripper_pin_"
        else:
            partDocument1 = documents1.Open(gvar.standard_path + "\\Pin\\" + a + ".CATPart")
            c = "Pin_"
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        # ------------------------------------------------------------↓pin長度
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        if i == 2:
            strParam1 = parameters1.Item("LP_D_L")
            excel_sheet_name = 'LP'
        else:
            strParam1 = parameters1.Item(Pilot_Pin_Material + "_D_L")
            excel_sheet_name = Pilot_Pin_Material
        iSize = strParam1.GetEnumerateValuesSize()  # STRING 裡面選擇數量
        myArray = [''] * (iSize)
        # strParam1.GetEnumerateValues(myArray)  # 抓取 STRING 的數值放入矩陣之中
        for ii in range(0, iSize):
            excelname = 'StandardPart'
            (BoltValue) = defs.ExcelSearch(excelname, excel_sheet_name, (ii + 1), a)
            myArray[ii] = BoltValue
        while JUMP == False:
            for k in range(0, iSize):
                plate_pin = str(Pin_data[1][i]) + "-" + str(b)
                if myArray[k] == plate_pin:
                    JUMP = True
                    break
            b = b - 1
        b += 1
        JUMP = False
        strParam1.Value = plate_pin
        Pin_data[2][i] = a + "-" + str(b)
        part1.Update()
        # ------------------------------------------------------------↑pin長度
        # -------------------------------------------搜尋現有pin尺寸
        file_name = os.listdir(gvar.save_path)
        if c + a + "-" + str(b) + ".CATPart" not in file_name:
            partDocument1 = catapp.ActiveDocument
            product1 = partDocument1.getItem(a)
            product1.PartNumber = c + a + "-" + str(b)  # 樹狀圖名稱
            partDocument1.SaveAs(gvar.save_path + c + a + "-" + str(b) + ".CATPart")  # 存檔名稱/存檔
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
    # ========================================================================================================================================↑模板定位Pin改變長度參數
    # ========================================================================================================================================↓內導柱改變長度參數
    for i in range(1, 1 + 1):  # ------------------------------同一副模具需要用到幾種尺寸的內導柱(設變數 往後便於修改規則時用)
        b = int(Inner_Guiding_data[2][i])  # -------------------內導柱長
        a = (Inner_Guiding_Post_Material + "_" + str(Inner_Guiding_Post_Diameter))
        # GoTo QQQQQ
        # if Inner_Guiding_data[1][i] = 8 :
        # a = Inner_Guiding_Post_Material + "_8"
        # End if
        # if Inner_Guiding_data[1][i] = 10 :
        # a = Inner_Guiding_Post_Material + "_10"
        # End if
        # if Inner_Guiding_data[1][i] = 13 :
        # a = Inner_Guiding_Post_Material + "_13"
        # End if
        # if Inner_Guiding_data[1][i] = 16 :
        # a = Inner_Guiding_Post_Material + "_16"
        # End if
        # if Inner_Guiding_data[1][i] = 20 :
        # a = Inner_Guiding_Post_Material + "_20"
        # End if
        # if Inner_Guiding_data[1][i] = 25 :
        # a = Inner_Guiding_Post_Material + "_25"
        # End if
        # QQQQQ:
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.standard_path + "\\Inner_Guiding_Post\\" + a + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        length1 = parameters1.Item(a + "_L")  # (Inner_Guiding_Post_Material)
        length1.Value = 100  # Val(Inner_Guiding_Post_Length)
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        partDocument1.SaveAs(
            gvar.save_path + "SGPH" + "_" + str(int(Inner_Guiding_data[1][i])) + "-" + str(b) + ".CATPart")
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
    # ========================================================================================================================================↑內導柱改變長度參數
    # ========================================================================================================================================↓內導套改變長度參數
    for i in range(1, 3):  # ------------------------------同一副模具需要用到幾種尺寸的內導柱(設變數 往後便於修改規則時用)
        a = (Under_Inner_Guiding_Post_Material + "_" + str(Inner_Guiding_Post_Diameter))  # --型號_直徑
        # b = form19.Combo4 + "_" + form19.Text7     #-------------------內導套高
        if i == 1:
            b = 20
        elif i == 2:
            b = 25
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.standard_path + "\\Inner_Guiding_Post\\" + a + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        length1 = parameters1.Item(Under_Inner_Guiding_Post_Material + "_d_L")
        length1.Value = str(int(Inner_Guiding_data[1][1])) + "-" + str(b)
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        partDocument1.SaveAs(
            gvar.save_path + Under_Inner_Guiding_Post_Material + "_" + str(int(Inner_Guiding_data[1][1])) + "-" + str(
                b) + ".CATPart")
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
    # ========================================================================================================================================↑內導套改變長度參數
    # ========================================================================================================================================↓等高套筒改變長度參數
    for i in range(1, 1 + 1):
        a = SBT_Material + "_" + str(SBT_data[1][1])
        b = SBT_Length  # -------------------Pin長
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.standard_path + "\\Shoulder_Screw\\" + a + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        strParam1 = parameters1.Item("D" + str(SBT_data[1][1]) + "_L")
        strParam1.Value = str(SBT_data[1][1]) + "-" + str(SBT_Length)
        SBT_data[7][1] = "Shoulder_Screw_" + a + "-" + str(SBT_Length)
        part1.Update()
        length6 = parameters1.Item("LB_hight")
        length6.Value = SBT_data[2][1] - SBT_CB_data[6][1] + int(
            gvar.strip_parameter_list[20])  # stripper_plate_height
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        product1 = partDocument1.getItem(a)
        product1.PartNumber = SBT_data[7][1]  # 樹狀圖名稱
        # partDocument1.SaveAs "C:\\Users\\PDAL\\Desktop\\ting_wei\\catia_input-CL224\\CB_10_60.CATPart"
        partDocument1.SaveAs(gvar.save_path + SBT_data[7][1] + ".CATPart")
        time.sleep(1)
        ##specsAndGeomWindow1 As SpecsAndGeomWindow
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
    # ========================================================================================================================================↑等高套筒改變長度參數
    # ========================================================================================================================================↓外導柱改變參數
    for i in range(1, 2):
        outer_Guiding_data[1][1] = out_Guide_Diameter  # --直徑
        outer_Guiding_data[2][1] = out_Guide_Length  # --長
        b = str(outer_Guiding_data[2][1])
        c = outer_Guiding_data[3][1] = out_Guide_Material  # --型號
        a = c + "_" + str(outer_Guiding_data[1][1])
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.standard_path + "\\Outer_Guiding_Post\\" + c + "\\" + a + "_down.CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        strParam1 = parameters1.Item(c + "_D_L")
        strParam1.Value = a + "-" + str(b)
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        partDocument1.SaveAs(gvar.save_path + a + "-" + str(b) + "_down.CATPart")
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.standard_path + "\\Outer_Guiding_Post\\" + c + "\\" + a + "_up.CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        strParam1 = parameters1.Item(c + "_D_L")
        strParam1.Value = a + "-" + str(b)
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        partDocument1.SaveAs(gvar.save_path + a + "-" + str(b) + "_up.CATPart")
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
    # ---------------------------------------------------------------------------↓外導柱螺栓改變長度參數
    for i in range(1, 3):
        if i == 1:
            t = round(1 / 3 * int(gvar.strip_parameter_list[32]))
        else:
            t = round(1 / 3 * int(gvar.strip_parameter_list[5]))
        if outer_Guiding_data[3][1] == "MYJP":
            if outer_Guiding_data[1][1] == 20:
                a = 5
                b = 15 + t
            if outer_Guiding_data[1][1] == 25:
                a = 8
                b = 20 + t
            if outer_Guiding_data[1][1] == 32:
                a = 10
                b = 20 + t
            if outer_Guiding_data[1][1] == 38:
                a = 10
                b = 25 + t
            if outer_Guiding_data[1][1] == 50:
                a = 12
                b = 25 + t
        if outer_Guiding_data[3][1] == "MYKP":
            if outer_Guiding_data[1][1] == 25:
                a = 6
                b = 20 + t
            if outer_Guiding_data[1][1] == 32:
                a = 8
                b = 20 + t
            if outer_Guiding_data[1][1] == 38:
                a = 8
                b = 25 + t
            if outer_Guiding_data[1][1] == 50:
                a = 12
                b = 25 + t
            if outer_Guiding_data[1][1] == 60:
                a = 14
                b = 25 + t
        # ---------------------------------------------------------------------------↑外導柱螺栓改變長度參數
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.standard_path + "\\Bolt\\CB_" + str(a) + ".CATPart")
        # ------------------------------------------------------------------↓螺栓判斷
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        strParam1 = parameters1.Item("CB_M_L")
        iSize = strParam1.GetEnumerateValuesSize()  # STRING 裡面選擇數量
        myArray = [''] * iSize
        # strParam1.GetEnumerateValues(myArray)  # 抓取 STRING 的數值放入矩陣之中
        for ii in range(0, iSize):
            excelname = 'StandardPart'
            (BoltValue) = defs.ExcelSearch(excelname, "CB", (ii + 1), "CB_" + str(a))
            myArray[ii] = BoltValue
        while JUMP == False:
            for j in range(0, iSize):
                outer_Guiding_bolt = str(a) + "-" + str(b)
                if myArray[j] == outer_Guiding_bolt:
                    JUMP = True
                    break
            b -= 1
            if b < 0:
                print('Error-429')
                sys.exit()
        b += 1
        JUMP = False
        strParam1.Value = outer_Guiding_bolt
        part1.Update()
        outer_Guiding_data[4][i] = "Bolt_CB_" + outer_Guiding_bolt
        # ------------------------------------------------------------------↑螺栓判斷
        # -------------------------------------------搜尋現有螺栓尺寸
        file_name = os.listdir(gvar.save_path)
        if outer_Guiding_data[4][i] + ".CATPart" not in file_name:
            partDocument1 = catapp.ActiveDocument
            product1 = partDocument1.getItem("CB_" + str(a))
            product1.PartNumber = outer_Guiding_data[4][i]  # 樹狀圖名稱
            partDocument1.SaveAs(gvar.save_path + outer_Guiding_data[4][i] + ".CATPart")  # 存檔名稱/存檔
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
        # ---------------------------------------------------------------------------↓外導柱Pin改變長度參數
        if i == 1:
            t = round(1 / 2 * int(gvar.strip_parameter_list[32]))
        else:
            t = round(1 / 2 * int(gvar.strip_parameter_list[5]))
        if outer_Guiding_data[3][1] == "MYJP":
            if outer_Guiding_data[1][1] == 20:
                a = 6
                b = 15 / 2 + t
            if outer_Guiding_data[1][1] == 25:
                a = 8
                b = 20 / 2 + t
            if outer_Guiding_data[1][1] == 32:
                a = 8
                b = 20 / 2 + t
            if outer_Guiding_data[1][1] == 38:
                a = 10
                b = 25 / 2 + t
            if outer_Guiding_data[1][1] == 50:
                a = 10
                b = 25 / 2 + t
        elif outer_Guiding_data[3][1] == "MYKP":
            if outer_Guiding_data[1][1] == 25:
                a = 8
                b = 20 / 2 + t
            if outer_Guiding_data[1][1] == 32:
                a = 8
                b = 20 / 2 + t
            if outer_Guiding_data[1][1] == 38:
                a = 10
                b = 25 / 2 + t
            if outer_Guiding_data[1][1] == 50:
                a = 10
                b = 25 / 2 + t
            if outer_Guiding_data[1][1] == 60:
                a = 13
                b = 25 + t
        b = int(b)
        # ---------------------------------------------------------------------------↑外導柱Pin改變長度參數
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.standard_path + "\\Pin\\MSTM_" + str(a) + ".CATPart")
        # ------------------------------------------------------------------↓Pin判斷
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        strParam1 = parameters1.Item("MSTM_D_L")
        iSize = strParam1.GetEnumerateValuesSize()  # STRING 裡面選擇數量
        myArray = [''] * iSize
        # strParam1.GetEnumerateValues(myArray)  # 抓取 STRING 的數值放入矩陣之中
        for ii in range(0, iSize):
            excelname = 'StandardPart'
            (BoltValue) = defs.ExcelSearch(excelname, "MSTM", (ii + 1), "MSTM" + str(a))
            myArray[ii] = BoltValue
        while JUMP == False:
            for j in range(0, iSize):
                outer_Guiding_pin = str(a) + "-" + str(b)
                if myArray[j] == outer_Guiding_pin:
                    JUMP = True
                    break
            b -= 1
            if b < 0:
                print('Error-509')
                sys.exit()
        b += 1
        JUMP = False
        strParam1.Value = outer_Guiding_pin
        part1.Update()
        outer_Guiding_data[5][i] = "Pin_MSTM_" + outer_Guiding_pin
        # ------------------------------------------------------------------↑Pin判斷
        # -------------------------------------------搜尋現有螺栓尺寸
        file_name = os.listdir(gvar.save_path)
        if outer_Guiding_data[5][i] + ".CATPart" not in file_name:
            partDocument1 = catapp.ActiveDocument
            product1 = partDocument1.getItem("MSTM_" + str(a))
            product1.PartNumber = outer_Guiding_data[5][i]  # 樹狀圖名稱
            partDocument1.SaveAs(gvar.save_path + outer_Guiding_data[5][i] + ".CATPart")
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
    # ========================================================================================================================================↑外導柱改變參數
    # ========================================================================================================================================↓脫料釘改變長度參數
    for i in range(1, 2):
        stripper_pin_data[1][1] = Diameter = Pilot_Punch_Stripper_punch_Diameter  # 脫料釘直徑
        stripper_pin_data[2][1] = length = Pilot_Punch_Stripper_punch_Length  # 脫料釘長度
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(
            gvar.standard_path + "\\Stripper_pin\\" + Pilot_Punch_Stripper_punch_Material + "_" + str(
                stripper_pin_data[1][1]) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        strParam1 = parameters1.Item("LP_D_L")
        strParam1.Value = str(stripper_pin_data[1][1]) + "-" + str(stripper_pin_data[2][1])
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        product1 = partDocument1.getItem("LP_" + str(stripper_pin_data[1][1]))
        product1.PartNumber = "Stripper_pin_" + Pilot_Punch_Stripper_punch_Material + "_" + str(
            stripper_pin_data[1][1]) + "D-" + str(
            stripper_pin_data[2][1]) + "L"  # 樹枝圖名稱 "LP" =Pilot_Punch_Stripper_punch_Material
        partDocument1 = catapp.ActiveDocument
        partDocument1.SaveAs(gvar.save_path + "Stripper_pin_" + Pilot_Punch_Stripper_punch_Material + "_" + str(
            stripper_pin_data[1][1]) + "D-" + str(stripper_pin_data[2][1]) + "L.CATPart")
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
        # ----------------------------------------------------------------------------------------脫料釘止付螺栓參數
        documents1 = catapp.Documents
        if stripper_pin_data[1][1] == 2:
            a = 4
        if stripper_pin_data[1][1] == 3:
            a = 5
        if stripper_pin_data[1][1] == 4:
            a = 8
        if stripper_pin_data[1][1] == 6:
            a = 10
        if stripper_pin_data[1][1] == 8:
            a = 12
        if stripper_pin_data[1][1] == 10:
            a = 16
        if stripper_pin_data[1][1] == 13:
            a = 20
        if stripper_pin_data[1][1] == 16:
            a = 22
        if stripper_pin_data[1][1] == 20:
            a = 27
        partDocument1 = documents1.Open(gvar.standard_path + "\\Stop_payment_bolt\\" + "MSW_M.CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        strParam1 = parameters1.Item("MSW_M")
        strParam1.Value = a
        stripper_pin_stop_bolt[1] = a  # 存止付螺栓直徑
        length1 = parameters1.Item("L")
        stripper_pin_stop_bolt[2] = length1.Value  # 存止付螺栓長度
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        product1 = partDocument1.getItem("MSW_M")
        product1.PartNumber = "MSW_M" + str(a)  # 樹枝圖名稱
        partDocument1 = catapp.ActiveDocument
        partDocument1.SaveAs(gvar.save_path + "MSW_M" + str(a) + ".CATPart")
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
        # ----------------------------------------------------------------------------------------脫料釘彈簧參數
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(gvar.standard_path + "\\Spring\\" + "Spring.CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        length1 = parameters1.Item("spring_D")
        length1.Value = round(Pilot_Punch_Stripper_punch_Diameter) * 2 - 2
        stripper_pin_spring[1] = length1.Value  # 脫料釘彈簧直徑
        length2 = parameters1.Item("spring_L")
        length2.Value = float(gvar.strip_parameter_list[17]) + 6 - stripper_pin_stop_bolt[2]
        stripper_pin_spring[2] = length2.Value  # 脫料釘彈簧長度
        part1.Update()
        partDocument1 = catapp.ActiveDocument
        product1 = partDocument1.getItem("Spring")
        product1.PartNumber = "Spring_" + str(stripper_pin_spring[1]) + "Dx" + str(
            stripper_pin_spring[2]) + "L"  # 樹枝圖名稱
        partDocument1 = catapp.ActiveDocument
        partDocument1.SaveAs(
            gvar.save_path + "Spring_" + str(stripper_pin_spring[1]) + "Dx" + str(stripper_pin_spring[2]) + "L.CATPart")
        time.sleep(1)
        specsAndGeomWindow1 = catapp.ActiveWindow
        specsAndGeomWindow1.Close()
        partDocument1 = catapp.ActiveDocument
        partDocument1.Close()
    # ========================================================================================================================================↑脫料釘改變長度參數
    return Pin_data, SBT_data, outer_Guiding_data, stripper_pin_data

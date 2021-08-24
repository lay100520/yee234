import win32com.client as win32
import global_var as gvar
import time


def drafting():
    # ---------------------------------↓讀取output記事本內容↓-----------------------------------------
    vntLines = [""] * (gvar.all_part_number + 1)
    drafting_page = int()
    drafting_total_page = int()
    for i in range(int(gvar.all_part_number + 1)):
        vntLines[i] = gvar.all_part_name[i]
    print(vntLines)
    for i in range(1, int(gvar.all_part_number) + 1):  # --------------------------出圖次數迴圈
        partname = vntLines[i]  # --------------------------開啟的零件名稱(表第一行,表第二行...)
        print(partname)
        drafting_page = drafting_page + 1  # -------------------頁數加疊
        drafting_total_page = int(gvar.all_part_number)  # ----------總頁數
        # ---------------------------------↑讀取output記事本內容↑-----------------------------------------
        catapp = win32.Dispatch('CATIA.Application')
        documents1 = catapp.Documents
        partDocument1 = documents1.Open(
            gvar.save_path + partname + ".CATPart")  # -----------開啟存檔路徑零件/零件名稱/(Ctia零件output)
        time.sleep(1)
        # -------------↓關閉不出圖零件↓-----------
        # ==========================================================↓建立極值點↓=========================================================
        # ------------↓選擇工作物件↓------------
        partbody_select = "Body.2"
        pad1 = "Pad.1"
        if "lower_die_" in partname:
            pad1 = "Pad.3"
        if partname == "lower_die_set":
            partbody_select = "PartBody"
        if partname == "lower_die_set":
            pad1 = "Pad.1"
        if partname == "upper_die_set":
            partbody_select = "PartBody"
        if "Stripper_" in partname:
            pad1 = "Pad.3"
        if "Splint_" in partname:
            pad1 = "Pad.3"
        if "Stop_plate_" in partname:
            pad1 = "Pad.3"
        if "up_plate_" in partname:
            pad1 = "Pad.3"
        if "_allotype_cut_up_punch_" in partname:
            partbody_select = "shoulder_punch"
            pad1 = "shoulder_Pad"
        if "_half_cut_punch_" in partname:
            partbody_select = "Body.1"
            pad1 = "Pad.8"
        if "_bending_punch_" in partname:
            partbody_select = "bending_punch"
        if "_emboss_forming_punch_" in partname:
            pad1 = "Pad.2"
        if "_A_punch_insert_" in partname:
            pad1 = "Pad.3"
        if "_cut_cavity_insert_" in partname:
            pad1 = "Pad.3"
        if "_A_punch_QR_Splint_insert_" in partname:
            pad1 = "Pad.3"
        if "_A_punch_QR_Stripper_insert_" in partname:
            pad1 = "Pad.3"
        # --------A沖↓↓↓↓↓
        if "_SJAS_" in partname:
            partDocument1.Close()
            continue
        if "_SJAL_" in partname:
            partDocument1.Close()
            continue
        if "_A_SJAS_" in partname:
            partDocument1.Close()
            continue
        if "_A_SJAL_" in partname:
            partDocument1.Close()
            continue
        # --------A沖↑↑↑↑↑
        if "_allotype_cut_insert_" in partname:
            partbody_select = "cut_cavity_insert"
        if "_cutting_cavity_" in partname and "_insert_" in partname:
            partbody_select = "cut_cavity_insert_shear"
            pad1 = "Pad.3"
        if "_cutting_cavity_d_insert_" in partname:
            partbody_select = "cut_cavity_insert_shear"
            pad1 = "Pad.1"
        if "_cutting_cavity_u_insert_" in partname:
            partbody_select = "cut_cavity_insert_shear"
            pad1 = "Pad.2"
        if "_bending_down_punch_surface_" in partname:
            partbody_select = "Body.1"
            pad1 = "Pad.1"
        if "_bending_down_punch_surface_02" in partname:
            partbody_select = "Body.2"
            pad1 = "Pad.4"
        if "_QR_bending_down_punch_" in partname:
            partbody_select = "Body.1"
            pad1 = "Pad.1"
        if "_bending_up_punch_" in partname:
            partbody_select = "Body_1"
            pad1 = "Pad.1"
        if "_bending_up_floating_blocks_hold_" in partname:
            partbody_select = "Body_1"
            pad1 = "Pad.1"
        if "_bending_up_floating_blocks_01" in partname:
            partbody_select = "Body_1"
            pad1 = "Pad.1"
        if "_bending_up_floating_blocks_02" in partname:
            partbody_select = "Body_2"
            pad1 = "Pad.3"
        if "_leveling_block_up_inbolt_" in partname:
            partbody_select = "Body.2"
            pad1 = "Pad.2"
        if "_leveling_block_down_inbolt_" in partname:
            partbody_select = "Body.2"
            pad1 = "Pad.1"
        if "Binder_Plate_" in partname:
            partbody_select = "PartBody"
            pad1 = "Pad.1"
        if "sensor" in partname:
            partDocument1.Close()
            continue
        # ------------↑選擇工作物件↑------------
        partDocument1.Selection.Clear()
        part1 = partDocument1.Part
        bodies1 = part1.Bodies
        time.sleep(0.5)
        print(partbody_select)
        body1 = bodies1.Item(partbody_select)  # ----------定義工作物件
        part1.InWorkObject = body1
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        shapes1 = body1.Shapes
        print(pad1)
        pad1 = shapes1.Item(pad1)
        reference1 = part1.CreateReferenceFromObject(pad1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        hybridShapeDirection3 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)
        hybridShapeExtremum1.Direction3 = hybridShapeDirection3
        hybridShapeExtremum1.ExtremumType3 = 1
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        part1.InWorkObject.Name = "X_max"  # ---------------修改極值點名稱
        partDocument1.Selection.Add(part1.InWorkObject)
        part1.Update()
        hybridShapeDirection4 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        reference2 = part1.CreateReferenceFromObject(pad1)
        hybridShapeExtremum2 = hybridShapeFactory1.AddNewExtremum(reference2, hybridShapeDirection4, 0)
        hybridShapeDirection5 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum2.Direction2 = hybridShapeDirection5
        hybridShapeExtremum2.ExtremumType2 = 1
        hybridShapeDirection6 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)
        hybridShapeExtremum2.Direction3 = hybridShapeDirection6
        hybridShapeExtremum2.ExtremumType3 = 1
        body1.InsertHybridShape(hybridShapeExtremum2)
        part1.InWorkObject = hybridShapeExtremum2
        part1.InWorkObject.Name = "X_min"  # ---------------修改極值點名稱
        partDocument1.Selection.Add(part1.InWorkObject)
        part1.Update()
        hybridShapeDirection7 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        reference3 = part1.CreateReferenceFromObject(pad1)
        hybridShapeExtremum3 = hybridShapeFactory1.AddNewExtremum(reference3, hybridShapeDirection7, 1)
        hybridShapeDirection8 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapeExtremum3.Direction2 = hybridShapeDirection8
        hybridShapeExtremum3.ExtremumType2 = 1
        hybridShapeDirection9 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)
        hybridShapeExtremum3.Direction3 = hybridShapeDirection9
        hybridShapeExtremum3.ExtremumType3 = 1
        body1.InsertHybridShape(hybridShapeExtremum3)
        part1.InWorkObject = hybridShapeExtremum3
        part1.InWorkObject.Name = "Y_max"  # ---------------修改極值點名稱
        partDocument1.Selection.Add(part1.InWorkObject)
        part1.Update()
        hybridShapeDirection10 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        reference4 = part1.CreateReferenceFromObject(pad1)
        hybridShapeExtremum4 = hybridShapeFactory1.AddNewExtremum(reference4, hybridShapeDirection10, 0)
        hybridShapeDirection11 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapeExtremum4.Direction2 = hybridShapeDirection11
        hybridShapeExtremum4.ExtremumType2 = 1
        hybridShapeDirection12 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)
        hybridShapeExtremum4.Direction3 = hybridShapeDirection12
        hybridShapeExtremum4.ExtremumType3 = 1
        body1.InsertHybridShape(hybridShapeExtremum4)
        part1.InWorkObject = hybridShapeExtremum4
        part1.InWorkObject.Name = "Y_min"  # ---------------修改極值點名稱
        partDocument1.Selection.Add(part1.InWorkObject)
        part1.Update()
        # ===================================================================================
        hybridShapeDirection13 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)
        reference5 = part1.CreateReferenceFromObject(pad1)
        hybridShapeExtremum5 = hybridShapeFactory1.AddNewExtremum(reference5, hybridShapeDirection13, 1)
        hybridShapeDirection14 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapeExtremum5.Direction2 = hybridShapeDirection14
        hybridShapeExtremum5.ExtremumType2 = 1
        hybridShapeDirection15 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum5.Direction3 = hybridShapeDirection15
        hybridShapeExtremum5.ExtremumType3 = 1
        body1.InsertHybridShape(hybridShapeExtremum5)
        part1.InWorkObject = hybridShapeExtremum5
        part1.InWorkObject.Name = "Z_max"  # ---------------修改極值點名稱
        partDocument1.Selection.Add(part1.InWorkObject)
        part1.Update()
        hybridShapeDirection16 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)
        reference6 = part1.CreateReferenceFromObject(pad1)
        hybridShapeExtremum6 = hybridShapeFactory1.AddNewExtremum(reference6, hybridShapeDirection16, 0)
        hybridShapeDirection17 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapeExtremum6.Direction2 = hybridShapeDirection17
        hybridShapeExtremum6.ExtremumType2 = 1
        hybridShapeDirection18 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum6.Direction3 = hybridShapeDirection18
        hybridShapeExtremum6.ExtremumType3 = 1
        body1.InsertHybridShape(hybridShapeExtremum6)
        part1.InWorkObject = hybridShapeExtremum6
        part1.InWorkObject.Name = "Z_min"
        partDocument1.Selection.Add(part1.InWorkObject)
        part1.Update()
        visPropertySet1 = partDocument1.Selection.VisProperties
        visPropertySet1 = visPropertySet1.Parent
        visPropertySet1.SetShow(1)
        partDocument1.Selection.Clear()
        # ===================================================================================
        MeasureDistance_number = 3
        item_belong = partbody_select
        Measure_distance_item = [""] * 7
        parameter_name = [""] * 4
        Measure_distance_item[1] = "X_max"
        Measure_distance_item[2] = "X_min"
        Measure_distance_item[3] = "Y_max"
        Measure_distance_item[4] = "Y_min"
        Measure_distance_item[5] = "Z_max"
        Measure_distance_item[6] = "Z_min"
        parameter_name[1] = "Length_max"
        parameter_name[2] = "Width_max"
        parameter_name[3] = "Height_max"
        # ============================================Measure_Distance
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        Measure_distance_item_number = 0
        for j in range(1, MeasureDistance_number + 1):
            parameters1 = part1.Parameters
            length1 = parameters1.CreateDimension("", "LENGTH", 0)
            length1.rename(parameter_name[j])
            relations1 = part1.Relations
            formula1 = relations1.CreateFormula("Formula_1", "", length1,
                                                "distance(" + item_belong + "\\" + Measure_distance_item[
                                                    1 + Measure_distance_item_number] + "," + item_belong + "\\" +
                                                Measure_distance_item[2 + Measure_distance_item_number] + " ) ")
            Measure_distance_item_number = Measure_distance_item_number + 2
            part1.Update()
        # ============================================Measure_Distance
        # ======================↓讀取長/寬/高參數↓======================
        parameters16 = part1.Parameters
        length = [None] * 10
        parameters17 = part1.Parameters
        length[1] = parameters17.Item("Length_max")
        Length_max = length[1].Value
        parameters18 = part1.Parameters
        length[2] = parameters18.Item("Width_max")
        Width_max = length[2].Value
        parameters19 = part1.Parameters
        length[3] = parameters19.Item("Height_max")
        Height_max = length[3].Value
        Frame_Thickness_1 = str(length[3].Value)
        part1.Update()
        L_range = Length_max + Height_max
        W_range = Width_max + Height_max
        size = str(Height_max) + "x" + str(Length_max) + "x" + str(Width_max)
        # ====↓設定性質↓=====================================
        parameters1 = part1.Parameters
        strParam1 = parameters1.Item("Properties\\Size")
        strParam1.Value = size
        # ================================================↓A4↓==========================================================
        if L_range <= 258 / 1.4 and W_range <= 168 / 2.2:
            # A4(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page)
            partDocument1.Close()
            pass
        # ================================================↑A4↑==========================================================
        # ================================================↓A3↓==========================================================
        elif L_range <= 379 / 1.6 and W_range <= 244 / 1.5:
            # A3(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page)
            partDocument1.Close()
            pass
        # ================================================↑A3↑==========================================================
        # ================================================↓A2↓==========================================================
        elif L_range <= 546 / 1.7 and W_range <= 342 / 1.5:
            A2(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page)
        # ================================================↑A2↑==========================================================
        # ================================================↓A1↓==========================================================
        elif L_range <= 790 / 2.1 and W_range <= 495 / 1.9:
            A1(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page)
        # ================================================↑A1↑==========================================================
        # ================================================↓A0↓==========================================================
        elif L_range <= 1138 / 1.5 and W_range <= 714 / 1.3:
            A0(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page)
        # ================================================↑A0↑==========================================================
        # ================================================↓A00↓=========================================================
        elif L_range > 1138 / 1.5 or W_range > 714 / 1.3:
            A00(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page)
        # ================================================↑A00↑==========================================================


def A4(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    total_op_number = int(gvar.strip_parameter_list[2])
    machining_explanation = [None] * 21
    label_determine_1 = int()
    label_determine_2 = int()
    label_determine_3 = int()
    label_determine_5 = int()
    label_determine_6 = int()
    label_determine_10 = int()
    label_determine_11 = int()
    label_determine_12 = int()
    label_determine_13 = int()
    label_determine_14 = int()
    # ==========================↓讀取圖名+材質+熱處理↓==========================
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    product1 = partDocument1.getItem("Part1")
    parameters4 = part1.Parameters
    strParam1 = parameters4.Item("`Properties\\Part Name`")
    parameters5 = part1.Parameters
    parameters6 = part1.Parameters
    parameters7 = part1.Parameters
    strParam4 = parameters7.Item("Properties\\L1")  # 形狀孔
    machining_explanation[1] = strParam4.Value
    parameters8 = part1.Parameters
    strParam5 = parameters8.Item("Properties\\A")  # 螺栓孔
    machining_explanation[2] = strParam5.Value
    parameters9 = part1.Parameters
    strParam6 = parameters9.Item("Properties\\HP")  # 合銷孔
    machining_explanation[3] = strParam6.Value
    parameters10 = part1.Parameters
    strParam7 = parameters10.Item("Properties\\B")  # B型引導沖孔
    machining_explanation[4] = strParam7.Value
    parameters11 = part1.Parameters
    strParam8 = parameters11.Item("Properties\\BP")  # B沖沖孔
    machining_explanation[5] = strParam8.Value
    parameters12 = part1.Parameters
    strParam9 = parameters12.Item("Properties\\TS")  # 浮升引導
    machining_explanation[6] = strParam9.Value
    parameters13 = part1.Parameters
    strParam10 = parameters13.Item("Properties\\IG")  # 內導柱
    machining_explanation[7] = strParam10.Value
    parameters14 = part1.Parameters
    strParam11 = parameters14.Item("Properties\\F")  # 外導柱
    machining_explanation[8] = strParam11.Value
    parameters15 = part1.Parameters
    strParam12 = parameters15.Item("Properties\\CS")  # 等高套筒
    machining_explanation[9] = strParam12.Value
    parameters20 = part1.Parameters
    strParam13 = parameters20.Item("Properties\\AP")  # A沖沖孔
    machining_explanation[10] = strParam13.Value
    (Frame_Material, Frame_Heat_treatment) = Momo_machining_explanation(partname)
    machining_explanation[
        17] = "(" + Frame_Material + "+ " + Frame_Heat_treatment + ")"  # -----------------------材質+熱處理
    machining_explanation[18] = "(" + str(Height_max) + "Tx" + str(Length_max) + "Lx" + str(
        Width_max) + "W" + ")"  # --------高x長x寬
    # -----------------------------------------------------------
    strParam2 = parameters5.CreateString("Size", "")
    Size_Name = "(" + str(Length_max) + "x" + str(Width_max) + "x" + str(Height_max) + ")"
    strParam2.ValuateFromString(Size_Name)
    # -----------------------------------------------------------
    # ===============================↓孔標籤類型判斷↓===============================
    selection1 = partDocument1.Selection
    point_break = False
    label_determine_4 = int()
    for total_op in range(1, 1 + total_op_number):  # --------------總工站數
        if point_break == True:
            point_break = False
            break
        total_op = str(total_op)
        for total_plate in range(1, 1 + 2):  # form7.Text8  #--------------總模板數  *********(之後要分模板)***********
            total_plate = str(total_plate)
            selection1.Clear()
            selection1.Search("Name = lower_die_set,all")  # ---下模座
            label_determine_1 = selection1.Count
            if label_determine_1 > 0:
                # ----------------------↓外導柱座標↓----------------------
                parameters25 = part1.Parameters
                length5 = parameters25.Item("lower_die_set\\Pillar_center_X")
                out_Guide_posts_X = length5.Value
                parameters26 = part1.Parameters
                length6 = parameters26.Item("lower_die_set\\Pillar_center_Y")
                out_Guide_posts_Y = length6.Value
                parameters27 = part1.Parameters
                length7 = parameters27.Item("lower_die_set\\Avoid_Error_X")
                out_Guide_posts_Avoid_Error_X = length7.Value
                parameters28 = part1.Parameters
                length8 = parameters28.Item("lower_die_set\\Pillar_a")
                out_Guide_posts_bolt_hole_X = length8.Value
                parameters29 = part1.Parameters
                length9 = parameters29.Item("lower_die_set\\Pillar_b")
                out_Guide_posts_bolt_hole_Y = length9.Value
                # ----------------------↑外導柱座標↑----------------------
                point_break = True
                break
            selection1.Search("Name = upper_die_set,all")  # ---上模座
            label_determine_2 = selection1.Count
            if label_determine_2 > 0:
                # ----------------------↓外導柱座標↓----------------------
                parameters30 = part1.Parameters
                length10 = parameters30.Item("Pillar_center_X")
                out_Guide_posts_X = length10.Value
                parameters31 = part1.Parameters
                length11 = parameters31.Item("Pillar_center_Y")
                out_Guide_posts_Y = length11.Value
                parameters32 = part1.Parameters
                length12 = parameters32.Item("Avoid_Error_X")
                out_Guide_posts_Avoid_Error_X = length12.Value
                parameters33 = part1.Parameters
                length13 = parameters33.Item("Pillar_a")
                out_Guide_posts_bolt_hole_X = length13.Value
                parameters34 = part1.Parameters
                length14 = parameters34.Item("Pillar_b")
                out_Guide_posts_bolt_hole_Y = length14.Value
                # ----------------------↑外導柱座標↑----------------------
                point_break = True
                break
            selection1.Search("Name = lower_die_" + total_plate + ",all")  # ---下模板
            label_determine_3 = selection1.Count
            if label_determine_3 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_pilot_punch_insert_0" + total_plate + " ,all")  # ---導引沖頭入子
            label_determine_4 = selection1.Count
            if label_determine_4 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cutting_cavity_d_insert_0" + total_plate + ",all")  # ---下靠刀沖頭入子
            label_determine_5 = selection1.Count
            if label_determine_5 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cutting_cavity_u_insert_0" + total_plate + ",all")  # ---上靠刀沖頭入子
            label_determine_6 = selection1.Count
            if label_determine_6 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cut_punch_d_cutting_add_0" + total_plate + ",all")  # ---下靠刀沖頭入塊
            label_determine_7 = selection1.Count
            if label_determine_7 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cut_punch_u_cutting_add_0" + total_plate + ",all")  # ---上靠刀沖頭入塊
            label_determine_8 = selection1.Count
            if label_determine_8 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_Stripper_insert_l_0" + total_plate + ",all")  # ---脫料入子
            label_determine_9 = selection1.Count
            if label_determine_9 > 0:
                point_break = True
                break
            selection1.Search("Name = lower_pad_" + total_plate + ",all")  # ---下墊板
            label_determine_10 = selection1.Count
            if label_determine_10 > 0:
                point_break = True
                break
            selection1.Search("Name = Stripper_" + total_plate + ",all")  # ---脫料板
            label_determine_11 = selection1.Count
            if label_determine_11 > 0:
                point_break = True
                break
            selection1.Search("Name = Stop_plate_" + total_plate + ",all")  # ---止擋板
            label_determine_12 = selection1.Count
            if label_determine_12 > 0:
                point_break = True
                break
            selection1.Search("Name = Splint_" + total_plate + ",all")  # ---上夾板
            label_determine_13 = selection1.Count
            if label_determine_13 > 0:
                point_break = True
                break
            selection1.Search("Name = up_plate_" + total_plate + ",all")  # ---上墊板
            label_determine_14 = selection1.Count
            if label_determine_14 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_cut_punch_d_cutting_" + total_plate + ",all")  # ---下靠刀沖頭
            label_determine_15 = selection1.Count
            if label_determine_15 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_cut_punch_u_cutting_" + total_plate + ",all")  # ---上靠刀沖頭
            label_determine_16 = selection1.Count
            if label_determine_16 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_forming_cavity_" + total_plate + ",all")  # ---下成形沖頭
            label_determine_17 = selection1.Count
            if label_determine_17 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_forming_punch_" + total_plate + ",all")  # --上成形沖頭
            label_determine_18 = selection1.Count
            if label_determine_18 > 0:
                point_break = True
                break
            selection1.Search("Name = pad_lower,all")  # ---墊腳
            label_determine_19 = selection1.Count
            if label_determine_19 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_QR_l_punch_0" + total_plate + ",all")  # ---靠肩沖頭
            label_determine_20 = selection1.Count
            if label_determine_20 > 0:
                point_break = True
                break
    selection1.Clear()
    part1.Update()
    # ================================↑孔標籤類型判斷↑===============================
    # =======================↓入子導引沖頭孔+螺栓孔座標(單一導引衝))↓=======================
    if label_determine_4 > 0:
        # -----------------↓導引沖頭孔↓-----------------
        body1 = bodies1.Item("PartBody")
        sketches1 = body1.Sketches
        originElements1 = part1.OriginElements
        reference1 = originElements1.PlaneXY
        sketch1 = sketches1.Add(reference1)
        arrayOfVariantOfDouble1 = [0, 0, 0, 1, 0, 0, 0, 1, 0]
        sketch1Variant = sketch1
        sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)
        part1.InWorkObject = sketch1
        factory2D1 = sketch1.OpenEdition()
        geometricElements1 = sketch1.GeometricElements
        axis2D1 = geometricElements1.Item("AbsoluteAxis")
        line2D1 = axis2D1.getItem("HDirection")
        line2D1.ReportName = 1
        line2D2 = axis2D1.getItem("VDirection")
        line2D2.ReportName = 2
        hybridShapes1 = body1.HybridShapes
        reference2 = hybridShapes1.Item("Y_min")
        geometricElements2 = factory2D1.CreateProjections(reference2)
        geometry2D1 = geometricElements2.Item("Mark.1")
        geometry2D1.Construction = True
        body2 = bodies1.Item("Body.2")
        hybridShapes2 = body2.HybridShapes
        reference3 = hybridShapes2.Item("Extremum.1(X_max)")
        geometricElements3 = factory2D1.CreateProjections(reference3)
        geometry2D2 = geometricElements3.Item("Mark.1")
        geometry2D2.Construction = True
        constraints1 = sketch1.Constraints
        reference4 = part1.CreateReferenceFromObject(geometry2D1)
        reference5 = part1.CreateReferenceFromObject(geometry2D2)
        reference6 = part1.CreateReferenceFromObject(line2D2)
        constraint1 = constraints1.AddTriEltCst(1, reference4, reference5, reference6)
        constraint1.mode = 1
        constraint1.Name = "pilot_punch_hole_Y_axis"
        reference7 = hybridShapes1.Item("X_min")
        geometricElements4 = factory2D1.CreateProjections(reference7)
        geometry2D3 = geometricElements4.Item("Mark.1")
        geometry2D3.Construction = True
        geometricElements5 = factory2D1.CreateProjections(reference3)
        geometry2D4 = geometricElements5.Item("Mark.1")
        geometry2D4.Construction = True
        reference8 = part1.CreateReferenceFromObject(geometry2D3)
        reference9 = part1.CreateReferenceFromObject(geometry2D4)
        reference10 = part1.CreateReferenceFromObject(line2D1)
        constraint2 = constraints1.AddTriEltCst(1, reference8, reference9, reference10)
        constraint2.mode = 1
        constraint2.Name = "pilot_punch_hole_X_axis"
        sketch1.CloseEdition()
        part1.InWorkObject = sketch1
        part1.InWorkObject.Name = "pilot_punch_label_Sketch"
        # -----------------↑導引沖頭孔↑-----------------
        # -------------------↓螺栓孔↓-------------------
        pilot_punch_insert_bolt_axis()
        # -------------------↑螺栓孔↑-------------------
        parameters21 = part1.Parameters
        length1 = parameters21.CreateDimension("", "LENGTH", 0)
        length1.rename("pilot_punch_hole_X_axis")
        parameters22 = part1.Parameters
        length2 = parameters22.CreateDimension("", "LENGTH", 0)
        length2.rename(
            "pilot_punch_hole_Y_axis")
        parameters23 = part1.Parameters
        length3 = parameters23.CreateDimension("", "LENGTH", 0)
        length3.rename(
            "bolt_hole_X_axis")
        parameters24 = part1.Parameters
        length4 = parameters24.CreateDimension("", "LENGTH", 0)
        length4.rename(
            "bolt_hole_Y_axis")
        part1.Update()
        sketches3 = body1.Sketches
        sketch3 = sketches3.Item("pilot_punch_label_Sketch")
        factory2D2 = sketch1.OpenEdition()
        relations1 = part1.Relations
        formula1 = relations1.Createformula("formula:pilot_punch_hole_X_axis", "", length1,
                                            "PartBody\\pilot_punch_label_Sketch\\pilot_punch_hole_X_axis\\Offset ")
        formula1.rename("formula_pilot_punch_hole_X_axis")
        relations2 = part1.Relations
        formula2 = relations2.Createformula("formula:pilot_punch_hole_Y_axis", "", length2,
                                            "PartBody\\pilot_punch_label_Sketch\\pilot_punch_hole_Y_axis\\Offset ")
        formula2.rename("formula_pilot_punch_hole_Y_axis")
        relations3 = part1.Relations
        formula3 = relations3.Createformula("formula:bolt_hole_X_axis", "", length3,
                                            "PartBody\\bolt_label_Sketch\\bolt_hole_X_axis\\Offset ")
        formula3.rename("formula_bolt_hole_X_axis")
        relations4 = part1.Relations
        formula4 = relations4.Createformula("formula:bolt_hole_X_axis", "", length4,
                                            "PartBody\\bolt_label_Sketch\\bolt_hole_Y_axis\\Offset ")
        formula4.rename("formula_bolt_hole_Y_axis")
        length = [None] * 5
        length[1] = parameters21.Item("pilot_punch_hole_X_axis")
        pilot_punch_hole_point_X = length[1].Value
        length[2] = parameters22.Item("pilot_punch_hole_Y_axis")
        pilot_punch_hole_point_Y = length[2].Value
        length[3] = parameters23.Item("bolt_hole_X_axis")
        pilot_punch_bolt_hole_point_X = length[3].Value
        length[4] = parameters24.Item("bolt_hole_Y_axis")
        pilot_punch_bolt_hole_point_Y = length[4].Value
        sketch1.CloseEdition()
    part1.Update()
    # =======================↑入子導引沖頭孔+螺栓孔座標(單一導引衝))↑=======================
    # =================================↓異形孔座標↓=================================
    # ------------------------------------↓建點↓------------------------------------
    part1 = partDocument1.Part
    selection_point_1 = partDocument1.Selection
    selection_point_2 = partDocument1.Selection
    selection_point_3 = partDocument1.Selection
    selection_point_4 = partDocument1.Selection
    selection_point_41 = partDocument1.Selection
    selection_point_5 = partDocument1.Selection
    selection_point_6 = partDocument1.Selection
    selection_point_7 = partDocument1.Selection
    selection_point_8 = partDocument1.Selection
    selection_point_9 = partDocument1.Selection
    selection_point_10 = partDocument1.Selection
    selection_point_11 = partDocument1.Selection
    selection_point_12 = partDocument1.Selection
    selection_point_13 = partDocument1.Selection
    for total_plate in range(1, 1 + 1):  # form7.Text8 #--------總模板數 *********(之後要分模板)***********
        total_plate = str(total_plate)
        for total_op in range(1, 1 + total_op_number):  # --------總工站數
            total_op = str(total_op)
            selection_point_1.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_punch_d_cutting_*_project_line,all")  # -----下靠刀沖頭孔
            plate_line_cut_punch_d_cutting_machining_shape = selection_point_1.Count
            if plate_line_cut_punch_d_cutting_machining_shape > 0:
                cut_punch_d_cutting_machining_shape_point(plate_line_cut_punch_d_cutting_machining_shape, total_op)
            # plate_line_cut_punch_d_cutting_machining_shape_number = 0 + plate_line_cut_punch_d_cutting_machining_shape #----------數量總和
            selection_point_2.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_punch_u_cutting_*_project_line,all")  # -----上靠刀沖頭孔
            plate_line_cut_punch_u_cutting_machining_shape = selection_point_2.Count
            if plate_line_cut_punch_u_cutting_machining_shape > 0:
                cut_punch_u_cutting_machining_shape_point(plate_line_cut_punch_u_cutting_machining_shape, total_op)
            selection_point_3.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_QR_l_punch_*_project_line,all")  # --------------右靠刀肩沖頭孔
            plate_line_QR_l_punch_machining_shape = selection_point_3.Count
            if plate_line_QR_l_punch_machining_shape > 0:
                QR_l_punch_machining_shape_point(plate_line_QR_l_punch_machining_shape, total_op)
            selection_point_4.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_line_*,all")  # ------------------------------剪切沖頭孔
            selection_point_next = selection_point_4.Count
            selection_point_41.Search("Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_line_*_line,all")
            plate_line_cut_line_machining_shape = (selection_point_next - selection_point_41.Count)
            if plate_line_cut_line_machining_shape > 0:
                plate_line_cut_line_machining_shape_point(plate_line_cut_line_machining_shape, total_op)
            selection_point_5.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_pilot_punch_insert_*_project_line,all")  # -------脫料入子孔
            plate_line_pilot_punch_insert_machining_shape = selection_point_5.Count
            if plate_line_pilot_punch_insert_machining_shape > 0:
                plate_line_pilot_punch_insert_machining_shape_point(plate_line_pilot_punch_insert_machining_shape,
                                                                    total_op)
            selection_point_6.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_Stripper_insert_l_*_project_line,all")  # --------導引沖頭入子孔
            plate_line_Stripper_insert_left_machining_shape = selection_point_6.Count
            if plate_line_Stripper_insert_left_machining_shape > 0:
                plate_line_Stripper_insert_left_machining_shape_point(plate_line_Stripper_insert_left_machining_shape,
                                                                      total_op)
            selection_point_7.Search(
                "Name = plate_line" + total_plate + "_op" + total_op + "0_cutting_cavity_d_insert_*_project_line,all")  # ---下沖頭入子孔
            plate_line_cutting_cavity_d_insert_machining_shape = selection_point_7.Count
            if plate_line_cutting_cavity_d_insert_machining_shape > 0:
                plate_line_cutting_cavity_d_insert_machining_shape_point(
                    plate_line_cutting_cavity_d_insert_machining_shape, total_op)
            selection_point_8.Search(
                "Name = plate_line" + total_plate + "_op" + total_op + "0_cutting_cavity_u_insert_*_project_line,all")  # ---上沖頭入子孔
            plate_line_cutting_cavity_u_insert_machining_shape = selection_point_8.Count
            if plate_line_cutting_cavity_u_insert_machining_shape > 0:
                plate_line_cutting_cavity_u_insert_machining_shape_point(
                    plate_line_cutting_cavity_u_insert_machining_shape, total_op)
            selection_point_9.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_forming_punch_Project_*,all")  # -----------------成形沖頭孔
            plate_line_forming_punch_machining_shape = selection_point_9.Count
            if plate_line_forming_punch_machining_shape > 0:
                plate_line_forming_punch_machining_shape_point(plate_line_forming_punch_machining_shape, total_op)
            selection_point_10.Search(
                "Name = op" + total_op + "0_cutting_cavity_d_insert_*,all")  # -----------------下靠刀入子沖頭孔(搜尋零件名稱)
            plate_line_cutting_cavity_d_insert_shape = selection_point_10.Count
            if plate_line_cutting_cavity_d_insert_shape > 0:
                plate_line_cutting_cavity_d_insert_shape_point(plate_line_cutting_cavity_d_insert_shape, total_op,
                                                               total_plate)
            selection_point_11.Search(
                "Name = op" + total_op + "0_cutting_cavity_u_insert_*,all")  # -----------------上靠刀入子沖頭孔(搜尋零件名稱)
            plate_line_cutting_cavity_u_insert_shape = selection_point_11.Count
            if plate_line_cutting_cavity_u_insert_shape > 0:
                plate_line_cutting_cavity_u_insert_shape_point(plate_line_cutting_cavity_u_insert_shape, total_op,
                                                               total_plate)
            selection_point_12.Search(
                "Name = op" + total_op + "0_Stripper_insert_l_*,all")  # -----------------下料沖頭入子沖孔(搜尋零件名稱)
            plate_line_Stripper_QR_l_punch_shape = selection_point_12.Count
            if plate_line_Stripper_QR_l_punch_shape > 0:
                plate_line_Stripper_QR_l_punch_shape_point(plate_line_Stripper_QR_l_punch_shape, total_op, total_plate)
            # -----------------↓例外(stop_plate),op_70_QR_l_punch_1_project更改名稱?
            selection_point_13.Search("Name = op_" + total_op + "0_QR_l_punch_*_project,all")  # 右靠肩沖頭孔(例外)
            QR_l_punch_machining_shape = selection_point_13.Count
            if QR_l_punch_machining_shape > 0:
                stop_plate_QR_l_punch_machining_shape_point(QR_l_punch_machining_shape, total_op)
            # -----------------↑例外(stop_plate),op_70_QR_l_punch_1_project更改名稱?
    part1.Update()
    # ------------------------------------↓座標↓------------------------------------
    selection_axis_1 = partDocument1.Selection
    selection_axis_1.Search("Name = machining_shape_point_*,all")
    machining_shape_point_number = selection1.Count
    if machining_shape_point_number > 0:
        (machining_shape_point_X, machining_shape_point_Y) = machining_shape_point_axis(partname,
                                                                                        machining_shape_point_number)
    # =================================↑異形孔座標↑================================
    # ================================================↓出圖↓=================================================
    drawingDocument1 = documents1.Open(gvar.open_path + "A0.CATDrawing")  # -------------開啟圖紙(input).
    drawingSheets1 = drawingDocument1.Sheets
    Scaleradio = 1
    viewdistance = 230  # -------------------------視圖間距離
    mainviewx = 96  # -------------------------主視圖X座標
    mainviewy = 150  # -------------------------主視圖Y座標
    topviewy = mainviewy + viewdistance / 3.4
    rightviewx = mainviewx + viewdistance / 2.7
    downviewy = mainviewy - viewdistance / 3.4
    leftview = mainviewx - viewdistance / 2.7
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    MyText1 = MyView.Texts.Add("＊如果對於圖面有更好的建議歡迎提出,我們會虛心接受", 28.56, 31.81)
    MyText1.SetFontSize(0, 0, 3.5)
    MyText2 = MyView.Texts.Add("未 注 公 差", 35.79, 26.22)
    MyText2.SetFontSize(0, 0, 3.5)
    MyText3 = MyView.Texts.Add("角度:", 53.24, 18.78)
    MyText3.SetFontSize(0, 0, 3.5)
    MyText4 = MyView.Texts.Add("圖名", 66.85, 26.23)
    MyText4.SetFontSize(0, 0, 3.5)
    MyText5 = MyView.Texts.Add("圖號", 66.85, 21.19)
    MyText5.SetFontSize(0, 0, 3.5)
    MyText6 = MyView.Texts.Add("路徑", 66.85, 16.32)
    MyText6.SetFontSize(0, 0, 3.5)
    MyText7 = MyView.Texts.Add("材  質", 103.15, 26.23)
    MyText7.SetFontSize(0, 0, 3.5)
    MyText8 = MyView.Texts.Add("熱處理", 103.15, 21.19)
    MyText8.SetFontSize(0, 0, 3.5)
    MyText9 = MyView.Texts.Add("板  厚", 103.15, 16.32)
    MyText9.SetFontSize(0, 0, 3.5)
    MyText10 = MyView.Texts.Add("圖檔比例", 131.23, 26.23)
    MyText10.SetFontSize(0, 0, 3.5)
    MyText12 = MyView.Texts.Add("投影方向", 131.23, 16.32)
    MyText12.SetFontSize(0, 0, 3.5)
    MyText13 = MyView.Texts.Add("設計", 157.19, 26.23)
    MyText13.SetFontSize(0, 0, 3.5)
    MyText13 = MyView.Texts.Add("檢查", 157.19, 21.19)
    MyText13.SetFontSize(0, 0, 3.5)
    MyText14 = MyView.Texts.Add("認可", 157.19, 16.32)
    MyText14.SetFontSize(0, 0, 3.5)
    MyText15 = MyView.Texts.Add("圖發部門", 197.17, 26.23)
    MyText15.SetFontSize(0, 0, 3.5)
    MyText16 = MyView.Texts.Add("客產編號", 197.17, 21.19)
    MyText16.SetFontSize(0, 0, 3.5)
    MyText17 = MyView.Texts.Add("圖印時間", 197.17, 16.32)
    MyText17.SetFontSize(0, 0, 3.5)
    MyText18 = MyView.Texts.Add("頁碼", 242.62, 16.32)
    MyText18.SetFontSize(0, 0, 3.5)
    MyText19 = MyView.Texts.Add("第    頁 ,", 250.48, 16.32)
    MyText19.SetFontSize(0, 0, 3.5)
    MyText20 = MyView.Texts.Add("共    頁", 270, 16.32)
    MyText20.SetFontSize(0, 0, 3.5)
    MyText21 = MyView.Texts.Add("金屬產品開發研究發展中心", 253, 21.35)
    MyText21.SetFontSize(0, 0, 2.8)
    MyText22 = MyView.Texts.Add(" M P R D C", 256.6, 25.5)
    MyText22.SetFontSize(0, 0, 3.5)
    MyText22 = MyView.Texts.Add("圖面版本記錄", 248, 194.67)
    MyText22.SetFontSize(0, 0, 3.5)
    MyText23 = MyView.Texts.Add("版本", 228.8, 189.63)
    MyText23.SetFontSize(0, 0, 3.5)
    MyText24 = MyView.Texts.Add("版本說明", 237.74, 189.63)
    MyText24.SetFontSize(0, 0, 3.5)
    MyText25 = MyView.Texts.Add("日  期", 255.89, 189.63)
    MyText25.SetFontSize(0, 0, 3.5)
    MyText26 = MyView.Texts.Add("設  計", 270.9, 189.63)
    # -----------------------------------------↓加工說明↓---------------------------------------
    MyText62 = MyView.Texts.Add("加工說明: (" + partname + ")", 215, 166)  # ---零件名稱
    MyText62.SetFontSize(0, 0, 3.5)
    machining_explanation_X = 230  # ------整個加工說明X座標
    machining_explanation_Y = 166  # ------整個加工說明Y座標
    machining_explanation_P = 0  # ------加工說明每行間距
    for M in range(1, 1 + 18):
        if machining_explanation[M] == None:
            break
        # ------------------------------------分割字串，一行28個字----------------------------------
        machining_explanation[M] = str(machining_explanation[M])
        MyLen = len(machining_explanation[M])  # 傳回字串中字元個數。
        if MyLen > 28:
            words = -int(-MyLen / 28)
            w = [""] * words + 1
            machining_explanation_temporary = machining_explanation[M][0:28]
            for i in range(1, words + 1):
                w[i] = machining_explanation[M][(i * 28) + 1: (i * 28) + 29]  # ----(i*28)+1:下一行從第29個字開始, 28:一行28個字
                machining_explanation_temporary = machining_explanation_temporary + "\n" + w[i]
            machining_explanation[M] = machining_explanation_temporary
        # ------------------------------------分割字串，一行28個字----------------------------------
        machining_explanation_P = machining_explanation_P + 5
        MyText63 = MyView.Texts.Add(machining_explanation[M], str(machining_explanation_X),
                                    str(machining_explanation_Y - machining_explanation_P))  # ---內容
        MyText63.SetFontSize(0, 0, 3.5)
        if MyLen > 28:
            machining_explanation_Y = machining_explanation_Y - 3.5  # -------字高
    # ---------------------------------A0常修改參數表格內容----------------------------------------------
    Frame_1 = "X.:"
    Frame_2 = "X.X:"
    Frame_3 = "+"
    Frame_4 = "-"
    Frame_5 = "+"
    Frame_6 = "-"
    Frame_7 = " "
    Frame_8 = " "
    Frame_9 = " "
    Frame_10 = " "
    Frame_11 = "X.XX:"
    Frame_12 = "X.XXX:"
    Frame_13 = "±"
    Frame_14 = "±"
    Frame_15 = "0.05"
    Frame_16 = "0.005"
    Frame_17 = "±"
    Frame_18 = "°"
    MyText26.SetFontSize(0, 0, 3.5)
    MyText27 = MyView.Texts.Add(Frame_1, 27.08, 20.5)
    MyText27.SetFontSize(0, 0, 2.5)
    MyText28 = MyView.Texts.Add(Frame_2, 27.08, 15.5)
    MyText28.SetFontSize(0, 0, 2.5)
    MyText29 = MyView.Texts.Add(Frame_11, 38.35, 20.5)
    MyText29.SetFontSize(0, 0, 2.5)
    MyText30 = MyView.Texts.Add(Frame_12, 38.35, 15.5)
    MyText30.SetFontSize(0, 0, 2.5)
    MyText31 = MyView.Texts.Add(Frame_3 + Frame_7, 32.57, 21.47)
    MyText31.SetFontSize(0, 0, 2)
    MyText32 = MyView.Texts.Add(Frame_4 + Frame_8, 32.5, 19.13)
    MyText32.SetFontSize(0, 0, 2)
    MyText33 = MyView.Texts.Add(Frame_5 + Frame_9, 32.5, 16.5)
    MyText33.SetFontSize(0, 0, 2)
    MyText34 = MyView.Texts.Add(Frame_6 + Frame_10, 32.57, 14.16)
    MyText34.SetFontSize(0, 0, 2)
    MyText35 = MyView.Texts.Add(Frame_13 + Frame_15, 45.5, 20.5)
    MyText35.SetFontSize(0, 0, 2.3)
    MyText36 = MyView.Texts.Add(Frame_14 + Frame_16, 45.5, 15.5)
    MyText36.SetFontSize(0, 0, 2.3)
    MyText37 = MyView.Texts.Add(Frame_17 + Frame_18, 61, 18.77)
    MyText37.SetFontSize(0, 0, 3)
    MyText38 = MyView.Texts.Add(partname, 76, 26.23)
    MyText38.SetFontSize(0, 0, 2.5)
    MyText39 = MyView.Texts.Add("USA035", 76, 21.5)  # 圖號
    MyText39.SetFontSize(0, 0, 3.5)
    MyText40 = MyView.Texts.Add("D:\小慈\實驗室", 76, 16.2)  # 路徑
    MyText40.SetFontSize(0, 0, 2.8)
    MyText41 = MyView.Texts.Add(Frame_Material, 115.26, 16.32)  # 自動材質
    MyText41.SetFontSize(0, 0, 3)
    MyText42 = MyView.Texts.Add(Frame_Heat_treatment, 145.08, 26.23)  # 自動熱處理
    MyText42.SetFontSize(0, 0, 3.5)
    MyText43 = MyView.Texts.Add(Frame_Thickness_1, 115.26, 16.32)  # 自動板厚
    MyText43.SetFontSize(0, 0, 3.5)
    MyText44 = MyView.Texts.Add("1：1", 145.08, 26.23)  # 圖檔比例
    MyText44.SetFontSize(0, 0, 3.5)
    MyText46 = MyView.Texts.Add("第三角", 145.08, 16.32)  # 投影方向
    MyText46.SetFontSize(0, 0, 3.5)
    MyText47 = MyView.Texts.Add(" ", 166.12, 26.23)  # 設計專員
    MyText47.SetFontSize(0, 0, 3.5)
    MyText48 = MyView.Texts.Add(" ", 181.13, 26.23)  # 設計專員
    MyText48.SetFontSize(0, 0, 3.5)
    MyText49 = MyView.Texts.Add(" ", 166.12, 21.19)  # 檢查專員
    MyText49.SetFontSize(0, 0, 3.5)
    MyText50 = MyView.Texts.Add(" ", 181.13, 21.19)  # 檢查專員
    MyText50.SetFontSize(0, 0, 3.5)
    MyText51 = MyView.Texts.Add(" ", 166.12, 16.32)  # 認可專員
    MyText51.SetFontSize(0, 0, 3.5)
    MyText52 = MyView.Texts.Add(" ", 181.13, 16.32)  # 認可專員
    MyText52.SetFontSize(0, 0, 3.5)
    MyText53 = MyView.Texts.Add("研發部", 215.16, 26.23)  # 圖發部門
    MyText53.SetFontSize(0, 0, 3.5)
    MyText54 = MyView.Texts.Add("654452", 215.16, 21.19)  # 客產編號
    MyText54.SetFontSize(0, 0, 3.5)
    now_time = time.strftime('%Y/%m/%d', time.localtime())
    MyText55 = MyView.Texts.Add(now_time, 215.16, 16.32)  # 圖印日期
    MyText55.SetFontSize(0, 0, 3.5)
    MyText56 = MyView.Texts.Add(drafting_page, 255.8, 16.32)
    MyText56.SetFontSize(0, 0, 3.5)
    MyText57 = MyView.Texts.Add(drafting_total_page, 276, 16.32)
    MyText57.SetFontSize(0, 0, 3.5)
    MyText58 = MyView.Texts.Add(" ", 228.8, 184.63)  # 版本
    MyText58.SetFontSize(0, 0, 3.5)
    MyText59 = MyView.Texts.Add(" ", 237.74, 184.63)  # 版本說明
    MyText59.SetFontSize(0, 0, 3.5)
    MyText60 = MyView.Texts.Add(now_time, 255.89, 184.63)  # 日期
    MyText60.SetFontSize(0, 0, 3.5)
    MyText61 = MyView.Texts.Add("專案人員", 270.9, 184.63)  # 設計人員
    MyText61.SetFontSize(0, 0, 3.5)
    # ---------------------------------------------------------↓投影↓---------------------------------------------------
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    product1 = partDocument1.getItem("Part1")  # 零件名稱
    drawingViewGenerativeBehavior1.Document = product1
    drawingViewGenerativeBehavior1.DefineFrontView(1, 0, 0, 0, 1, 0)
    drawingView1.X = mainviewx
    drawingView1.Y = mainviewy
    drawingView1.Scale = Scaleradio
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    drawingViewGenerativeBehavior1.Update()
    drawingView1.Activate()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingSheet1.ProjectionMethod = 1
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    # ---------------------------------↓存svg檔↓---------------------------------------------
    if "lower_die_" in partname or "lower_pad_" in partname or "upper_die_set" in partname or "Stripper_" in partname or "Splint_" in partname or "Stop_plate_" in partname or "up_plate_" in partname :
        if "insert" not in partname:
            svg(partname)
    # ---------------------------------↑存svg檔↑---------------------------------------------
    # ---------------------------------右視圖-------------------------------------------------------------
    drawingView2 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingViewGenerativeBehavior2.DefineProjectionView(drawingViewGenerativeBehavior1, 0)
    drawingViewGenerativeLinks1 = drawingView2.GenerativeLinks
    drawingViewGenerativeLinks2 = drawingView1.GenerativeLinks
    drawingViewGenerativeLinks2.CopyLinksTo(drawingViewGenerativeLinks1)
    drawingView2.X = rightviewx
    drawingView2.Y = mainviewy
    double1 = drawingView1.Scale
    drawingView2.Scale = Scaleradio
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingViewGenerativeBehavior2.Update()
    drawingView2.ReferenceView = drawingView1
    drawingView2.AlignedWithReferenceView()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    # ---------------------------------下視圖-------------------------------------------------------------
    drawingView4 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeBehavior4 = drawingView4.GenerativeBehavior
    drawingViewGenerativeBehavior4.DefineProjectionView(drawingViewGenerativeBehavior1, 3)
    drawingViewGenerativeLinks4 = drawingView4.GenerativeLinks
    drawingViewGenerativeLinks2 = drawingView1.GenerativeLinks
    drawingViewGenerativeLinks2.CopyLinksTo(drawingViewGenerativeLinks4)
    drawingView4.X = mainviewx
    drawingView4.Y = downviewy
    double3 = drawingView1.Scale
    drawingView4.Scale = Scaleradio
    drawingViewGenerativeBehavior4 = drawingView4.GenerativeBehavior
    drawingViewGenerativeBehavior4.Update()
    drawingView4.ReferenceView = drawingView1
    drawingView4.AlignedWithReferenceView()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    # ================↓切換孔標籤模組↓================
    if label_determine_1 > 0:  # 下模座
        lower_die_set_label(machining_shape_point_number, mainviewx, Length_max, mainviewy, Width_max,
                            machining_shape_point_X, machining_shape_point_Y, out_Guide_posts_X, out_Guide_posts_Y,
                            out_Guide_posts_bolt_hole_X, out_Guide_posts_bolt_hole_Y, out_Guide_posts_Avoid_Error_X)
    if label_determine_2 > 0:  # 上模座
        upper_die_set_label()
    if label_determine_3 > 0:  # 下模板
        lower_die_label(machining_shape_point_number, mainviewx, Length_max, mainviewy, Width_max,
                        machining_shape_point_X,
                        machining_shape_point_Y)
    if label_determine_4 > 0:  # 導引沖入子
        pilot_punch_insert_label()
    if label_determine_5 > 0:  # 下靠刀沖頭入子
        cutting_cavity_d_insert_label(machining_explanation, machining_shape_point_number, mainviewx, mainviewy,
                                      Length_max,
                                      machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_6 > 0:  # 上靠刀沖頭入子
        cutting_cavity_u_insert_label(machining_explanation, machining_shape_point_number, mainviewx, mainviewy,
                                      Length_max,
                                      machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_10 > 0:  # 下背板
        lower_pad_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                        machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_11 > 0:  # 脫料板
        Stripper_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                       machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_12 > 0:  # 止擋板
        Stop_plate_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                         machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_13 > 0:  # 上夾板
        Splint_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                     machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_14 > 0:  # 上墊板
        up_plate_label()
    # ================↑切換孔標籤模組↑================
    # ============================================↓存檔↓============================================
    partDocument1.save()
    partDocument1.Close()
    drawingDocument1.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".CATDrawing",
                                "CATDrawing")  # 更新儲存路徑(2D output)
    drawingDocument1.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".dwg", "dwg")  # 使用dwg存檔
    drawingDocument1.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".jpg", "jpg")  # 使用JPG存檔
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()
    drawingDocument1 = catapp.ActiveDocument
    drawingDocument1.Close()


def A3(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    total_op_number = int(gvar.strip_parameter_list[2])
    machining_explanation = [None] * 21
    label_determine_1 = int()
    label_determine_2 = int()
    label_determine_3 = int()
    label_determine_5 = int()
    label_determine_6 = int()
    label_determine_10 = int()
    label_determine_11 = int()
    label_determine_12 = int()
    label_determine_13 = int()
    label_determine_14 = int()
    # ==========================↓讀取圖名+材質+熱處理↓==========================
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    product1 = partDocument1.getItem("Part1")
    parameters4 = part1.Parameters
    strParam1 = parameters4.Item("`Properties\\Part Name`")
    parameters5 = part1.Parameters
    parameters6 = part1.Parameters
    parameters7 = part1.Parameters
    strParam4 = parameters7.Item("Properties\\L1")  # 形狀孔
    machining_explanation[1] = strParam4.Value
    parameters8 = part1.Parameters
    strParam5 = parameters8.Item("Properties\\A")  # 螺栓孔
    machining_explanation[2] = strParam5.Value
    parameters9 = part1.Parameters
    strParam6 = parameters9.Item("Properties\\HP")  # 合銷孔
    machining_explanation[3] = strParam6.Value
    parameters10 = part1.Parameters
    strParam7 = parameters10.Item("Properties\\B")  # B型引導沖孔
    machining_explanation[4] = strParam7.Value
    parameters11 = part1.Parameters
    strParam8 = parameters11.Item("Properties\\BP")  # B沖沖孔
    machining_explanation[5] = strParam8.Value
    parameters12 = part1.Parameters
    strParam9 = parameters12.Item("Properties\\TS")  # 浮升引導
    machining_explanation[6] = strParam9.Value
    parameters13 = part1.Parameters
    strParam10 = parameters13.Item("Properties\\IG")  # 內導柱
    machining_explanation[7] = strParam10.Value
    parameters14 = part1.Parameters
    strParam11 = parameters14.Item("Properties\\F")  # 外導柱
    machining_explanation[8] = strParam11.Value
    parameters15 = part1.Parameters
    strParam12 = parameters15.Item("Properties\\CS")  # 等高套筒
    machining_explanation[9] = strParam12.Value
    parameters20 = part1.Parameters
    strParam13 = parameters20.Item("Properties\\AP")  # A沖沖孔
    machining_explanation[10] = strParam13.Value
    (Frame_Material, Frame_Heat_treatment) = Momo_machining_explanation(partname)
    machining_explanation[
        17] = "(" + Frame_Material + "+ " + Frame_Heat_treatment + ")"  # -----------------------材質+熱處理
    machining_explanation[18] = "(" + str(Height_max) + "Tx" + str(Length_max) + "Lx" + str(
        Width_max) + "W" + ")"  # --------高x長x寬
    # -----------------------------------------------------------
    strParam2 = parameters5.CreateString("Size", "")
    Size_Name = "(" + str(Length_max) + "x" + str(Width_max) + "x" + str(Height_max) + ")"
    strParam2.ValuateFromString(Size_Name)
    # -----------------------------------------------------------
    # ===============================↓孔標籤類型判斷↓===============================
    selection1 = partDocument1.Selection
    point_break = False
    label_determine_4 = int()
    for total_op in range(1, 1 + total_op_number):  # --------------總工站數
        if point_break == True:
            point_break = False
            break
        total_op = str(total_op)
        for total_plate in range(1, 1 + 2):  # form7.Text8  #--------------總模板數  *********(之後要分模板)***********
            total_plate = str(total_plate)
            selection1.Clear()
            selection1.Search("Name = lower_die_set,all")  # ---下模座
            label_determine_1 = selection1.Count
            if label_determine_1 > 0:
                # ----------------------↓外導柱座標↓----------------------
                parameters25 = part1.Parameters
                length5 = parameters25.Item("lower_die_set\\Pillar_center_X")
                out_Guide_posts_X = length5.Value
                parameters26 = part1.Parameters
                length6 = parameters26.Item("lower_die_set\\Pillar_center_Y")
                out_Guide_posts_Y = length6.Value
                parameters27 = part1.Parameters
                length7 = parameters27.Item("lower_die_set\\Avoid_Error_X")
                out_Guide_posts_Avoid_Error_X = length7.Value
                parameters28 = part1.Parameters
                length8 = parameters28.Item("lower_die_set\\Pillar_a")
                out_Guide_posts_bolt_hole_X = length8.Value
                parameters29 = part1.Parameters
                length9 = parameters29.Item("lower_die_set\\Pillar_b")
                out_Guide_posts_bolt_hole_Y = length9.Value
                # ----------------------↑外導柱座標↑----------------------
                point_break = True
                break
            selection1.Search("Name = upper_die_set,all")  # ---上模座
            label_determine_2 = selection1.Count
            if label_determine_2 > 0:
                # ----------------------↓外導柱座標↓----------------------
                parameters30 = part1.Parameters
                length10 = parameters30.Item("Pillar_center_X")
                out_Guide_posts_X = length10.Value
                parameters31 = part1.Parameters
                length11 = parameters31.Item("Pillar_center_Y")
                out_Guide_posts_Y = length11.Value
                parameters32 = part1.Parameters
                length12 = parameters32.Item("Avoid_Error_X")
                out_Guide_posts_Avoid_Error_X = length12.Value
                parameters33 = part1.Parameters
                length13 = parameters33.Item("Pillar_a")
                out_Guide_posts_bolt_hole_X = length13.Value
                parameters34 = part1.Parameters
                length14 = parameters34.Item("Pillar_b")
                out_Guide_posts_bolt_hole_Y = length14.Value
                # ----------------------↑外導柱座標↑----------------------
                point_break = True
                break
            selection1.Search("Name = lower_die_" + total_plate + ",all")  # ---下模板
            label_determine_3 = selection1.Count
            if label_determine_3 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_pilot_punch_insert_0" + total_plate + " ,all")  # ---導引沖頭入子
            label_determine_4 = selection1.Count
            if label_determine_4 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cutting_cavity_d_insert_0" + total_plate + ",all")  # ---下靠刀沖頭入子
            label_determine_5 = selection1.Count
            if label_determine_5 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cutting_cavity_u_insert_0" + total_plate + ",all")  # ---上靠刀沖頭入子
            label_determine_6 = selection1.Count
            if label_determine_6 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cut_punch_d_cutting_add_0" + total_plate + ",all")  # ---下靠刀沖頭入塊
            label_determine_7 = selection1.Count
            if label_determine_7 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cut_punch_u_cutting_add_0" + total_plate + ",all")  # ---上靠刀沖頭入塊
            label_determine_8 = selection1.Count
            if label_determine_8 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_Stripper_insert_l_0" + total_plate + ",all")  # ---脫料入子
            label_determine_9 = selection1.Count
            if label_determine_9 > 0:
                point_break = True
                break
            selection1.Search("Name = lower_pad_" + total_plate + ",all")  # ---下墊板
            label_determine_10 = selection1.Count
            if label_determine_10 > 0:
                point_break = True
                break
            selection1.Search("Name = Stripper_" + total_plate + ",all")  # ---脫料板
            label_determine_11 = selection1.Count
            if label_determine_11 > 0:
                point_break = True
                break
            selection1.Search("Name = Stop_plate_" + total_plate + ",all")  # ---止擋板
            label_determine_12 = selection1.Count
            if label_determine_12 > 0:
                point_break = True
                break
            selection1.Search("Name = Splint_" + total_plate + ",all")  # ---上夾板
            label_determine_13 = selection1.Count
            if label_determine_13 > 0:
                point_break = True
                break
            selection1.Search("Name = up_plate_" + total_plate + ",all")  # ---上墊板
            label_determine_14 = selection1.Count
            if label_determine_14 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_cut_punch_d_cutting_" + total_plate + ",all")  # ---下靠刀沖頭
            label_determine_15 = selection1.Count
            if label_determine_15 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_cut_punch_u_cutting_" + total_plate + ",all")  # ---上靠刀沖頭
            label_determine_16 = selection1.Count
            if label_determine_16 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_forming_cavity_" + total_plate + ",all")  # ---下成形沖頭
            label_determine_17 = selection1.Count
            if label_determine_17 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_forming_punch_" + total_plate + ",all")  # --上成形沖頭
            label_determine_18 = selection1.Count
            if label_determine_18 > 0:
                point_break = True
                break
            selection1.Search("Name = pad_lower,all")  # ---墊腳
            label_determine_19 = selection1.Count
            if label_determine_19 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_QR_l_punch_0" + total_plate + ",all")  # ---靠肩沖頭
            label_determine_20 = selection1.Count
            if label_determine_20 > 0:
                point_break = True
                break
    selection1.Clear()
    part1.Update()
    # ================================↑孔標籤類型判斷↑===============================
    # =======================↓入子導引沖頭孔+螺栓孔座標(單一導引衝))↓=======================
    if label_determine_4 > 0:
        # -----------------↓導引沖頭孔↓-----------------
        body1 = bodies1.Item("PartBody")
        sketches1 = body1.Sketches
        originElements1 = part1.OriginElements
        reference1 = originElements1.PlaneXY
        sketch1 = sketches1.Add(reference1)
        arrayOfVariantOfDouble1 = [0, 0, 0, 1, 0, 0, 0, 1, 0]
        sketch1Variant = sketch1
        sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)
        part1.InWorkObject = sketch1
        factory2D1 = sketch1.OpenEdition()
        geometricElements1 = sketch1.GeometricElements
        axis2D1 = geometricElements1.Item("AbsoluteAxis")
        line2D1 = axis2D1.getItem("HDirection")
        line2D1.ReportName = 1
        line2D2 = axis2D1.getItem("VDirection")
        line2D2.ReportName = 2
        hybridShapes1 = body1.HybridShapes
        reference2 = hybridShapes1.Item("Y_min")
        geometricElements2 = factory2D1.CreateProjections(reference2)
        geometry2D1 = geometricElements2.Item("Mark.1")
        geometry2D1.Construction = True
        body2 = bodies1.Item("Body.2")
        hybridShapes2 = body2.HybridShapes
        reference3 = hybridShapes2.Item("Extremum.1(X_max)")
        geometricElements3 = factory2D1.CreateProjections(reference3)
        geometry2D2 = geometricElements3.Item("Mark.1")
        geometry2D2.Construction = True
        constraints1 = sketch1.Constraints
        reference4 = part1.CreateReferenceFromObject(geometry2D1)
        reference5 = part1.CreateReferenceFromObject(geometry2D2)
        reference6 = part1.CreateReferenceFromObject(line2D2)
        constraint1 = constraints1.AddTriEltCst(1, reference4, reference5, reference6)
        constraint1.mode = 1
        constraint1.Name = "pilot_punch_hole_Y_axis"
        reference7 = hybridShapes1.Item("X_min")
        geometricElements4 = factory2D1.CreateProjections(reference7)
        geometry2D3 = geometricElements4.Item("Mark.1")
        geometry2D3.Construction = True
        geometricElements5 = factory2D1.CreateProjections(reference3)
        geometry2D4 = geometricElements5.Item("Mark.1")
        geometry2D4.Construction = True
        reference8 = part1.CreateReferenceFromObject(geometry2D3)
        reference9 = part1.CreateReferenceFromObject(geometry2D4)
        reference10 = part1.CreateReferenceFromObject(line2D1)
        constraint2 = constraints1.AddTriEltCst(1, reference8, reference9, reference10)
        constraint2.mode = 1
        constraint2.Name = "pilot_punch_hole_X_axis"
        sketch1.CloseEdition()
        part1.InWorkObject = sketch1
        part1.InWorkObject.Name = "pilot_punch_label_Sketch"
        # -----------------↑導引沖頭孔↑-----------------
        # -------------------↓螺栓孔↓-------------------
        pilot_punch_insert_bolt_axis()
        # -------------------↑螺栓孔↑-------------------
        parameters21 = part1.Parameters
        length1 = parameters21.CreateDimension("", "LENGTH", 0)
        length1.rename("pilot_punch_hole_X_axis")
        parameters22 = part1.Parameters
        length2 = parameters22.CreateDimension("", "LENGTH", 0)
        length2.rename(
            "pilot_punch_hole_Y_axis")
        parameters23 = part1.Parameters
        length3 = parameters23.CreateDimension("", "LENGTH", 0)
        length3.rename(
            "bolt_hole_X_axis")
        parameters24 = part1.Parameters
        length4 = parameters24.CreateDimension("", "LENGTH", 0)
        length4.rename(
            "bolt_hole_Y_axis")
        part1.Update()
        sketches3 = body1.Sketches
        sketch3 = sketches3.Item("pilot_punch_label_Sketch")
        factory2D2 = sketch1.OpenEdition()
        relations1 = part1.Relations
        formula1 = relations1.Createformula("formula:pilot_punch_hole_X_axis", "", length1,
                                            "PartBody\\pilot_punch_label_Sketch\\pilot_punch_hole_X_axis\\Offset ")
        formula1.rename("formula_pilot_punch_hole_X_axis")
        relations2 = part1.Relations
        formula2 = relations2.Createformula("formula:pilot_punch_hole_Y_axis", "", length2,
                                            "PartBody\\pilot_punch_label_Sketch\\pilot_punch_hole_Y_axis\\Offset ")
        formula2.rename("formula_pilot_punch_hole_Y_axis")
        relations3 = part1.Relations
        formula3 = relations3.Createformula("formula:bolt_hole_X_axis", "", length3,
                                            "PartBody\\bolt_label_Sketch\\bolt_hole_X_axis\\Offset ")
        formula3.rename("formula_bolt_hole_X_axis")
        relations4 = part1.Relations
        formula4 = relations4.Createformula("formula:bolt_hole_X_axis", "", length4,
                                            "PartBody\\bolt_label_Sketch\\bolt_hole_Y_axis\\Offset ")
        formula4.rename("formula_bolt_hole_Y_axis")
        length = [None] * 5
        length[1] = parameters21.Item("pilot_punch_hole_X_axis")
        pilot_punch_hole_point_X = length[1].Value
        length[2] = parameters22.Item("pilot_punch_hole_Y_axis")
        pilot_punch_hole_point_Y = length[2].Value
        length[3] = parameters23.Item("bolt_hole_X_axis")
        pilot_punch_bolt_hole_point_X = length[3].Value
        length[4] = parameters24.Item("bolt_hole_Y_axis")
        pilot_punch_bolt_hole_point_Y = length[4].Value
        sketch1.CloseEdition()
    part1.Update()
    # =======================↑入子導引沖頭孔+螺栓孔座標(單一導引衝))↑=======================
    # =================================↓異形孔座標↓=================================
    # ------------------------------------↓建點↓------------------------------------
    part1 = partDocument1.Part
    selection_point_1 = partDocument1.Selection
    selection_point_2 = partDocument1.Selection
    selection_point_3 = partDocument1.Selection
    selection_point_4 = partDocument1.Selection
    selection_point_41 = partDocument1.Selection
    selection_point_5 = partDocument1.Selection
    selection_point_6 = partDocument1.Selection
    selection_point_7 = partDocument1.Selection
    selection_point_8 = partDocument1.Selection
    selection_point_9 = partDocument1.Selection
    selection_point_10 = partDocument1.Selection
    selection_point_11 = partDocument1.Selection
    selection_point_12 = partDocument1.Selection
    selection_point_13 = partDocument1.Selection
    for total_plate in range(1, 1 + 1):  # form7.Text8 #--------總模板數 *********(之後要分模板)***********
        total_plate = str(total_plate)
        for total_op in range(1, 1 + total_op_number):  # --------總工站數
            total_op = str(total_op)
            selection_point_1.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_punch_d_cutting_*_project_line,all")  # -----下靠刀沖頭孔
            plate_line_cut_punch_d_cutting_machining_shape = selection_point_1.Count
            if plate_line_cut_punch_d_cutting_machining_shape > 0:
                cut_punch_d_cutting_machining_shape_point(plate_line_cut_punch_d_cutting_machining_shape, total_op)
            # plate_line_cut_punch_d_cutting_machining_shape_number = 0 + plate_line_cut_punch_d_cutting_machining_shape #----------數量總和
            selection_point_2.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_punch_u_cutting_*_project_line,all")  # -----上靠刀沖頭孔
            plate_line_cut_punch_u_cutting_machining_shape = selection_point_2.Count
            if plate_line_cut_punch_u_cutting_machining_shape > 0:
                cut_punch_u_cutting_machining_shape_point(plate_line_cut_punch_u_cutting_machining_shape, total_op)
            selection_point_3.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_QR_l_punch_*_project_line,all")  # --------------右靠刀肩沖頭孔
            plate_line_QR_l_punch_machining_shape = selection_point_3.Count
            if plate_line_QR_l_punch_machining_shape > 0:
                QR_l_punch_machining_shape_point(plate_line_QR_l_punch_machining_shape, total_op)
            selection_point_4.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_line_*,all")  # ------------------------------剪切沖頭孔
            selection_point_next = selection_point_4.Count
            selection_point_41.Search("Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_line_*_line,all")
            plate_line_cut_line_machining_shape = (selection_point_next - selection_point_41.Count)
            if plate_line_cut_line_machining_shape > 0:
                plate_line_cut_line_machining_shape_point(plate_line_cut_line_machining_shape, total_op)
            selection_point_5.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_pilot_punch_insert_*_project_line,all")  # -------脫料入子孔
            plate_line_pilot_punch_insert_machining_shape = selection_point_5.Count
            if plate_line_pilot_punch_insert_machining_shape > 0:
                plate_line_pilot_punch_insert_machining_shape_point(plate_line_pilot_punch_insert_machining_shape,
                                                                    total_op)
            selection_point_6.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_Stripper_insert_l_*_project_line,all")  # --------導引沖頭入子孔
            plate_line_Stripper_insert_left_machining_shape = selection_point_6.Count
            if plate_line_Stripper_insert_left_machining_shape > 0:
                plate_line_Stripper_insert_left_machining_shape_point(plate_line_Stripper_insert_left_machining_shape,
                                                                      total_op)
            selection_point_7.Search(
                "Name = plate_line" + total_plate + "_op" + total_op + "0_cutting_cavity_d_insert_*_project_line,all")  # ---下沖頭入子孔
            plate_line_cutting_cavity_d_insert_machining_shape = selection_point_7.Count
            if plate_line_cutting_cavity_d_insert_machining_shape > 0:
                plate_line_cutting_cavity_d_insert_machining_shape_point(
                    plate_line_cutting_cavity_d_insert_machining_shape, total_op)
            selection_point_8.Search(
                "Name = plate_line" + total_plate + "_op" + total_op + "0_cutting_cavity_u_insert_*_project_line,all")  # ---上沖頭入子孔
            plate_line_cutting_cavity_u_insert_machining_shape = selection_point_8.Count
            if plate_line_cutting_cavity_u_insert_machining_shape > 0:
                plate_line_cutting_cavity_u_insert_machining_shape_point(
                    plate_line_cutting_cavity_u_insert_machining_shape, total_op)
            selection_point_9.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_forming_punch_Project_*,all")  # -----------------成形沖頭孔
            plate_line_forming_punch_machining_shape = selection_point_9.Count
            if plate_line_forming_punch_machining_shape > 0:
                plate_line_forming_punch_machining_shape_point(plate_line_forming_punch_machining_shape, total_op)
            selection_point_10.Search(
                "Name = op" + total_op + "0_cutting_cavity_d_insert_*,all")  # -----------------下靠刀入子沖頭孔(搜尋零件名稱)
            plate_line_cutting_cavity_d_insert_shape = selection_point_10.Count
            if plate_line_cutting_cavity_d_insert_shape > 0:
                plate_line_cutting_cavity_d_insert_shape_point(plate_line_cutting_cavity_d_insert_shape, total_op,
                                                               total_plate)
            selection_point_11.Search(
                "Name = op" + total_op + "0_cutting_cavity_u_insert_*,all")  # -----------------上靠刀入子沖頭孔(搜尋零件名稱)
            plate_line_cutting_cavity_u_insert_shape = selection_point_11.Count
            if plate_line_cutting_cavity_u_insert_shape > 0:
                plate_line_cutting_cavity_u_insert_shape_point(plate_line_cutting_cavity_u_insert_shape, total_op,
                                                               total_plate)
            selection_point_12.Search(
                "Name = op" + total_op + "0_Stripper_insert_l_*,all")  # -----------------下料沖頭入子沖孔(搜尋零件名稱)
            plate_line_Stripper_QR_l_punch_shape = selection_point_12.Count
            if plate_line_Stripper_QR_l_punch_shape > 0:
                plate_line_Stripper_QR_l_punch_shape_point(plate_line_Stripper_QR_l_punch_shape, total_op, total_plate)
            # -----------------↓例外(stop_plate),op_70_QR_l_punch_1_project更改名稱?
            selection_point_13.Search("Name = op_" + total_op + "0_QR_l_punch_*_project,all")  # 右靠肩沖頭孔(例外)
            QR_l_punch_machining_shape = selection_point_13.Count
            if QR_l_punch_machining_shape > 0:
                stop_plate_QR_l_punch_machining_shape_point(QR_l_punch_machining_shape, total_op)
            # -----------------↑例外(stop_plate),op_70_QR_l_punch_1_project更改名稱?
    part1.Update()
    # ------------------------------------↓座標↓------------------------------------
    selection_axis_1 = partDocument1.Selection
    selection_axis_1.Search("Name = machining_shape_point_*,all")
    machining_shape_point_number = selection1.Count
    if machining_shape_point_number > 0:
        (machining_shape_point_X, machining_shape_point_Y) = machining_shape_point_axis(partname,
                                                                                        machining_shape_point_number)
    # =================================↑異形孔座標↑================================
    # ================================================↓出圖↓=================================================
    drawingDocument1 = documents1.Open(gvar.open_path + "A0.CATDrawing")  # -------------開啟圖紙(input).
    drawingSheets1 = drawingDocument1.Sheets
    Scaleradio = 1
    viewdistance = 230  # -------------------------視圖間距離
    mainviewx = 96  # -------------------------主視圖X座標
    mainviewy = 150  # -------------------------主視圖Y座標
    topviewy = mainviewy + viewdistance / 3.4
    rightviewx = mainviewx + viewdistance / 2.7
    downviewy = mainviewy - viewdistance / 3.4
    leftview = mainviewx - viewdistance / 2.7
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    MyText1 = MyView.Texts.Add("＊如果對於圖面有更好的建議歡迎提出,我們會虛心接受", 30.12, 43.58)
    MyText1.SetFontSize(0, 0, 5)
    MyText2 = MyView.Texts.Add("未 注 公 差", 44.65, 35.04)
    MyText2.SetFontSize(0, 0, 5)
    MyText3 = MyView.Texts.Add("角度:", 70.05, 23.67)
    MyText3.SetFontSize(0, 0, 5)
    MyText4 = MyView.Texts.Add("圖名", 93.81, 35.04)
    MyText4.SetFontSize(0, 0, 5)
    MyText5 = MyView.Texts.Add("圖號", 93.81, 27.05)
    MyText5.SetFontSize(0, 0, 5)
    MyText6 = MyView.Texts.Add("路徑", 93.81, 19.19)
    MyText6.SetFontSize(0, 0, 5)
    MyText7 = MyView.Texts.Add("材  質", 147.17, 35.04)
    MyText7.SetFontSize(0, 0, 5)
    MyText8 = MyView.Texts.Add("熱處理", 147.17, 27.25)
    MyText8.SetFontSize(0, 0, 5)
    MyText9 = MyView.Texts.Add("板  厚", 147.17, 19.19)
    MyText9.SetFontSize(0, 0, 5)
    MyText10 = MyView.Texts.Add("圖檔比例", 181, 35.04)
    MyText10.SetFontSize(0, 0, 5)
    MyText12 = MyView.Texts.Add("投影方向", 181, 19.19)
    MyText12.SetFontSize(0, 0, 5)
    MyText13 = MyView.Texts.Add("設計", 219.51, 35.04)
    MyText13.SetFontSize(0, 0, 5)
    MyText13 = MyView.Texts.Add("檢查", 219.51, 27.25)
    MyText13.SetFontSize(0, 0, 5)
    MyText14 = MyView.Texts.Add("認可", 219.51, 19.19)
    MyText14.SetFontSize(0, 0, 5)
    MyText15 = MyView.Texts.Add("圖發部門", 279.3, 35.04)
    MyText15.SetFontSize(0, 0, 5)
    MyText16 = MyView.Texts.Add("客產編號", 279.3, 27.25)
    MyText16.SetFontSize(0, 0, 5)
    MyText17 = MyView.Texts.Add("圖印時間", 279.3, 19.19)
    MyText17.SetFontSize(0, 0, 5)
    MyText18 = MyView.Texts.Add("頁碼", 346.47, 19.19)
    MyText18.SetFontSize(0, 0, 5)
    MyText19 = MyView.Texts.Add("第    頁 ,", 360.31, 19.19)
    MyText19.SetFontSize(0, 0, 5)
    MyText20 = MyView.Texts.Add("共    頁", 386, 19.19)
    MyText20.SetFontSize(0, 0, 5)
    MyText21 = MyView.Texts.Add("金屬產品開發研究發展中心", 361, 27)
    MyText21.SetFontSize(0, 0, 4)
    MyText22 = MyView.Texts.Add(" M P R D C", 366, 34)
    MyText22.SetFontSize(0, 0, 5)
    MyText22 = MyView.Texts.Add("圖面版本記錄", 352.5, 280)
    MyText22.SetFontSize(0, 0, 5)
    MyText23 = MyView.Texts.Add("版本", 326.59, 272.37)
    MyText23.SetFontSize(0, 0, 5)
    MyText24 = MyView.Texts.Add("版本說明", 340.25, 272.37)
    MyText24.SetFontSize(0, 0, 5)
    MyText25 = MyView.Texts.Add("日  期", 367.05, 272.37)
    MyText25.SetFontSize(0, 0, 5)
    MyText26 = MyView.Texts.Add("設  計", 388.41, 272.37)
    MyText26.SetFontSize(0, 0, 5)
    # -----------------------------------------↓加工說明↓---------------------------------------
    MyText62 = MyView.Texts.Add("加工說明: (" + partname + ")", 300, 245)  # ---零件名稱
    MyText62.SetFontSize(0, 0, 6)
    machining_explanation_X = 320  # ------整個加工說明X座標
    machining_explanation_Y = 245  # ------整個加工說明Y座標
    machining_explanation_P = 0  # ------加工說明每行間距
    for M in range(1, 1 + 18):
        if machining_explanation[M] == None:
            break
        # ------------------------------------分割字串，一行28個字----------------------------------
        machining_explanation[M] = str(machining_explanation[M])
        MyLen = len(machining_explanation[M])  # 傳回字串中字元個數。
        if MyLen > 25:
            words = -int(-MyLen / 25)
            w = [""] * words + 1
            machining_explanation_temporary = machining_explanation[M][0:25]
            for i in range(1, words + 1):
                w[i] = machining_explanation[M][(i * 25) + 1: (i * 25) + 26]  # ----(i*25)+1:下一行從第69個字開始, 25:一行25個字
                machining_explanation_temporary = machining_explanation_temporary + "\n" + w[i]
            machining_explanation[M] = machining_explanation_temporary
        # ------------------------------------分割字串，一行28個字----------------------------------
        machining_explanation_P = machining_explanation_P + 10
        MyText63 = MyView.Texts.Add(machining_explanation[M], str(machining_explanation_X),
                                    str(machining_explanation_Y - machining_explanation_P))  # ---內容
        MyText63.SetFontSize(0, 0, 5)
        if MyLen > 25:
            machining_explanation_Y = machining_explanation_Y - 5  # -------字高
    # ---------------------------------A0常修改參數表格內容----------------------------------------------
    Frame_1 = "X.:"
    Frame_2 = "X.X:"
    Frame_3 = "+"
    Frame_4 = "-"
    Frame_5 = "+"
    Frame_6 = "-"
    Frame_7 = " "
    Frame_8 = " "
    Frame_9 = " "
    Frame_10 = " "
    Frame_11 = "X.XX:"
    Frame_12 = "X.XXX:"
    Frame_13 = "±"
    Frame_14 = "±"
    Frame_15 = "0.05"
    Frame_16 = "0.005"
    Frame_17 = "±"
    Frame_18 = "°"
    MyText26.SetFontSize(0, 0, 5)
    MyText27 = MyView.Texts.Add(Frame_1, 28.17, 27.13)
    MyText27.SetFontSize(0, 0, 4)
    MyText28 = MyView.Texts.Add(Frame_2, 28.17, 19.07)
    MyText28.SetFontSize(0, 0, 4)
    MyText29 = MyView.Texts.Add(Frame_11, 46.63, 27.13)
    MyText29.SetFontSize(0, 0, 4)
    MyText30 = MyView.Texts.Add(Frame_12, 46.63, 19.07)
    MyText30.SetFontSize(0, 0, 4)
    MyText31 = MyView.Texts.Add(Frame_3 + Frame_7, 36.76, 27.88)
    MyText31.SetFontSize(0, 0, 3)
    MyText32 = MyView.Texts.Add(Frame_4 + Frame_8, 36.76, 24.57)
    MyText32.SetFontSize(0, 0, 3)
    MyText33 = MyView.Texts.Add(Frame_5 + Frame_9, 36.76, 19.81)
    MyText33.SetFontSize(0, 0, 3)
    MyText34 = MyView.Texts.Add(Frame_6 + Frame_10, 36.76, 16.5)
    MyText34.SetFontSize(0, 0, 3)
    MyText35 = MyView.Texts.Add(Frame_13 + Frame_15, 57.74, 27.13)
    MyText35.SetFontSize(0, 0, 3.5)
    MyText36 = MyView.Texts.Add(Frame_14 + Frame_16, 57.74, 19.04)
    MyText36.SetFontSize(0, 0, 3.5)
    MyText37 = MyView.Texts.Add(Frame_17 + Frame_18, 81, 23.64)
    MyText37.SetFontSize(0, 0, 5)
    MyText38 = MyView.Texts.Add(partname, 107.8, 35.04)
    MyText38.SetFontSize(0, 0, 4)
    MyText39 = MyView.Texts.Add("USA035", 107.8, 27.25)  # 圖號
    MyText39.SetFontSize(0, 0, 5)
    MyText40 = MyView.Texts.Add("D:\小慈\實驗室", 107.8, 19.19)  # 路徑
    MyText40.SetFontSize(0, 0, 3.5)
    MyText41 = MyView.Texts.Add(Frame_Material, 165.24, 35.04)  # 自動材質
    MyText41.SetFontSize(0, 0, 5)
    MyText42 = MyView.Texts.Add(Frame_Heat_treatment, 165.24, 27.25)  # 自動熱處理
    MyText42.SetFontSize(0, 0, 5)
    MyText43 = MyView.Texts.Add(Frame_Thickness_1, 165.24, 19.19)  # 自動板厚
    MyText43.SetFontSize(0, 0, 5)
    MyText44 = MyView.Texts.Add("1：1", 201.35, 35.04)  # 圖檔比例
    MyText44.SetFontSize(0, 0, 5)
    MyText46 = MyView.Texts.Add("第三角", 201.35, 19.19)  # 投影方向
    MyText46.SetFontSize(0, 0, 5)
    MyText47 = MyView.Texts.Add(" ", 233, 35.04)  # 設計專員
    MyText47.SetFontSize(0, 0, 5)
    MyText48 = MyView.Texts.Add(" ", 255, 35.04)  # 設計專員
    MyText48.SetFontSize(0, 0, 5)
    MyText49 = MyView.Texts.Add(" ", 233, 27.25)  # 檢查專員
    MyText49.SetFontSize(0, 0, 5)
    MyText50 = MyView.Texts.Add(" ", 255, 27.25)  # 檢查專員
    MyText50.SetFontSize(0, 0, 5)
    MyText51 = MyView.Texts.Add(" ", 233, 19.19)  # 認可專員
    MyText51.SetFontSize(0, 0, 5)
    MyText52 = MyView.Texts.Add(" ", 255, 19.19)  # 認可專員
    MyText52.SetFontSize(0, 0, 5)
    MyText53 = MyView.Texts.Add("研發部", 306.87, 35.04)  # 圖發部門
    MyText53.SetFontSize(0, 0, 5)
    MyText54 = MyView.Texts.Add("654452", 306.87, 27.25)  # 客產編號
    MyText54.SetFontSize(0, 0, 5)
    now_time = time.strftime('%Y/%m/%d', time.localtime())
    MyText55 = MyView.Texts.Add(now_time, 306.87, 19.19)  # 圖印日期
    MyText55.SetFontSize(0, 0, 5)
    MyText56 = MyView.Texts.Add(drafting_page, 369, 19.19)
    MyText56.SetFontSize(0, 0, 5)
    MyText57 = MyView.Texts.Add(drafting_total_page, 394, 19.19)
    MyText57.SetFontSize(0, 0, 5)
    MyText58 = MyView.Texts.Add(" ", 326.59, 264.15)  # 版本
    MyText58.SetFontSize(0, 0, 5)
    MyText59 = MyView.Texts.Add(" ", 340.25, 264.15)  # 版本說明
    MyText59.SetFontSize(0, 0, 5)
    MyText60 = MyView.Texts.Add(now_time, 367.05, 264.15)  # 日期
    MyText60.SetFontSize(0, 0, 4)
    MyText61 = MyView.Texts.Add("專案人員", 388.14, 264.15)  # 設計人員
    MyText61.SetFontSize(0, 0, 5)
    # ---------------------------------------------------------↓投影↓---------------------------------------------------
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    product1 = partDocument1.getItem("Part1")  # 零件名稱
    drawingViewGenerativeBehavior1.Document = product1
    drawingViewGenerativeBehavior1.DefineFrontView(1, 0, 0, 0, 1, 0)
    drawingView1.X = mainviewx
    drawingView1.Y = mainviewy
    drawingView1.Scale = Scaleradio
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    drawingViewGenerativeBehavior1.Update()
    drawingView1.Activate()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingSheet1.ProjectionMethod = 1
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    # ---------------------------------↓存svg檔↓---------------------------------------------
    if "lower_die_" in partname or "lower_pad_" in partname or "upper_die_set" in partname or "Stripper_" in partname or "Splint_" in partname or "Stop_plate_" in partname or "up_plate_" in partname:
        if "insert" not in partname:
            svg(partname)
    # ---------------------------------↑存svg檔↑---------------------------------------------
    # ---------------------------------右視圖-------------------------------------------------------------
    drawingView2 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingViewGenerativeBehavior2.DefineProjectionView(drawingViewGenerativeBehavior1, 0)
    drawingViewGenerativeLinks1 = drawingView2.GenerativeLinks
    drawingViewGenerativeLinks2 = drawingView1.GenerativeLinks
    drawingViewGenerativeLinks2.CopyLinksTo(drawingViewGenerativeLinks1)
    drawingView2.X = rightviewx
    drawingView2.Y = mainviewy
    double1 = drawingView1.Scale
    drawingView2.Scale = Scaleradio
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingViewGenerativeBehavior2.Update()
    drawingView2.ReferenceView = drawingView1
    drawingView2.AlignedWithReferenceView()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    # ---------------------------------下視圖-------------------------------------------------------------
    drawingView4 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeBehavior4 = drawingView4.GenerativeBehavior
    drawingViewGenerativeBehavior4.DefineProjectionView(drawingViewGenerativeBehavior1, 3)
    drawingViewGenerativeLinks4 = drawingView4.GenerativeLinks
    drawingViewGenerativeLinks2 = drawingView1.GenerativeLinks
    drawingViewGenerativeLinks2.CopyLinksTo(drawingViewGenerativeLinks4)
    drawingView4.X = mainviewx
    drawingView4.Y = downviewy
    double3 = drawingView1.Scale
    drawingView4.Scale = Scaleradio
    drawingViewGenerativeBehavior4 = drawingView4.GenerativeBehavior
    drawingViewGenerativeBehavior4.Update()
    drawingView4.ReferenceView = drawingView1
    drawingView4.AlignedWithReferenceView()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    # ================↓切換孔標籤模組↓================
    if label_determine_1 > 0:  # 下模座
        lower_die_set_label(machining_shape_point_number, mainviewx, Length_max, mainviewy, Width_max,
                            machining_shape_point_X, machining_shape_point_Y, out_Guide_posts_X, out_Guide_posts_Y,
                            out_Guide_posts_bolt_hole_X, out_Guide_posts_bolt_hole_Y, out_Guide_posts_Avoid_Error_X)
    if label_determine_2 > 0:  # 上模座
        upper_die_set_label()
    if label_determine_3 > 0:  # 下模板
        lower_die_label(machining_shape_point_number, mainviewx, Length_max, mainviewy, Width_max,
                        machining_shape_point_X,
                        machining_shape_point_Y)
    if label_determine_4 > 0:  # 導引沖入子
        pilot_punch_insert_label()
    if label_determine_5 > 0:  # 下靠刀沖頭入子
        cutting_cavity_d_insert_label(machining_explanation, machining_shape_point_number, mainviewx, mainviewy,
                                      Length_max,
                                      machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_6 > 0:  # 上靠刀沖頭入子
        cutting_cavity_u_insert_label(machining_explanation, machining_shape_point_number, mainviewx, mainviewy,
                                      Length_max,
                                      machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_10 > 0:  # 下背板
        lower_pad_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                        machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_11 > 0:  # 脫料板
        Stripper_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                       machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_12 > 0:  # 止擋板
        Stop_plate_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                         machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_13 > 0:  # 上夾板
        Splint_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                     machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_14 > 0:  # 上墊板
        up_plate_label()
    # ================↑切換孔標籤模組↑================
    # ============================================↓存檔↓============================================
    partDocument1.save()
    partDocument1.Close()
    drawingDocument1.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".CATDrawing",
                                "CATDrawing")  # 更新儲存路徑(2D output)
    drawingDocument1.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".dwg", "dwg")  # 使用dwg存檔
    drawingDocument1.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".jpg", "jpg")  # 使用JPG存檔
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()
    drawingDocument1 = catapp.ActiveDocument
    drawingDocument1.Close()


def A2(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    partDocument1.Close()
    pass


def A1(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    partDocument1.Close()
    pass


def A0(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    total_op_number = int(gvar.strip_parameter_list[2])
    machining_explanation = [None] * 21
    label_determine_1 = int()
    label_determine_2 = int()
    label_determine_3 = int()
    label_determine_5 = int()
    label_determine_6 = int()
    label_determine_10 = int()
    label_determine_11 = int()
    label_determine_12 = int()
    label_determine_13 = int()
    label_determine_14 = int()
    # ==========================↓讀取圖名+材質+熱處理↓==========================
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    product1 = partDocument1.getItem("Part1")
    parameters4 = part1.Parameters
    strParam1 = parameters4.Item("`Properties\\Part Name`")
    parameters5 = part1.Parameters
    parameters6 = part1.Parameters
    parameters7 = part1.Parameters
    strParam4 = parameters7.Item("Properties\\L1")  # 形狀孔
    machining_explanation[1] = strParam4.Value
    parameters8 = part1.Parameters
    strParam5 = parameters8.Item("Properties\\A")  # 螺栓孔
    machining_explanation[2] = strParam5.Value
    parameters9 = part1.Parameters
    strParam6 = parameters9.Item("Properties\\HP")  # 合銷孔
    machining_explanation[3] = strParam6.Value
    parameters10 = part1.Parameters
    strParam7 = parameters10.Item("Properties\\B")  # B型引導沖孔
    machining_explanation[4] = strParam7.Value
    parameters11 = part1.Parameters
    strParam8 = parameters11.Item("Properties\\BP")  # B沖沖孔
    machining_explanation[5] = strParam8.Value
    parameters12 = part1.Parameters
    strParam9 = parameters12.Item("Properties\\TS")  # 浮升引導
    machining_explanation[6] = strParam9.Value
    parameters13 = part1.Parameters
    strParam10 = parameters13.Item("Properties\\IG")  # 內導柱
    machining_explanation[7] = strParam10.Value
    parameters14 = part1.Parameters
    strParam11 = parameters14.Item("Properties\\F")  # 外導柱
    machining_explanation[8] = strParam11.Value
    parameters15 = part1.Parameters
    strParam12 = parameters15.Item("Properties\\CS")  # 等高套筒
    machining_explanation[9] = strParam12.Value
    parameters20 = part1.Parameters
    strParam13 = parameters20.Item("Properties\\AP")  # A沖沖孔
    machining_explanation[10] = strParam13.Value
    (Frame_Material, Frame_Heat_treatment) = Momo_machining_explanation(partname)
    machining_explanation[
        17] = "(" + Frame_Material + "+ " + Frame_Heat_treatment + ")"  # -----------------------材質+熱處理
    machining_explanation[18] = "(" + str(Height_max) + "Tx" + str(Length_max) + "Lx" + str(
        Width_max) + "W" + ")"  # --------高x長x寬
    # -----------------------------------------------------------
    strParam2 = parameters5.CreateString("Size", "")
    Size_Name = "(" + str(Length_max) + "x" + str(Width_max) + "x" + str(Height_max) + ")"
    strParam2.ValuateFromString(Size_Name)
    # -----------------------------------------------------------
    # ===============================↓孔標籤類型判斷↓===============================
    selection1 = partDocument1.Selection
    point_break = False
    label_determine_4 = int()
    for total_op in range(1, 1 + total_op_number):  # --------------總工站數
        if point_break == True:
            point_break = False
            break
        total_op = str(total_op)
        for total_plate in range(1, 1 + 2):  # form7.Text8  #--------------總模板數  *********(之後要分模板)***********
            total_plate = str(total_plate)
            selection1.Clear()
            selection1.Search("Name = lower_die_set,all")  # ---下模座
            label_determine_1 = selection1.Count
            if label_determine_1 > 0:
                # ----------------------↓外導柱座標↓----------------------
                parameters25 = part1.Parameters
                length5 = parameters25.Item("lower_die_set\\Pillar_center_X")
                out_Guide_posts_X = length5.Value
                parameters26 = part1.Parameters
                length6 = parameters26.Item("lower_die_set\\Pillar_center_Y")
                out_Guide_posts_Y = length6.Value
                parameters27 = part1.Parameters
                length7 = parameters27.Item("lower_die_set\\Avoid_Error_X")
                out_Guide_posts_Avoid_Error_X = length7.Value
                parameters28 = part1.Parameters
                length8 = parameters28.Item("lower_die_set\\Pillar_a")
                out_Guide_posts_bolt_hole_X = length8.Value
                parameters29 = part1.Parameters
                length9 = parameters29.Item("lower_die_set\\Pillar_b")
                out_Guide_posts_bolt_hole_Y = length9.Value
                # ----------------------↑外導柱座標↑----------------------
                point_break = True
                break
            selection1.Search("Name = upper_die_set,all")  # ---上模座
            label_determine_2 = selection1.Count
            if label_determine_2 > 0:
                # ----------------------↓外導柱座標↓----------------------
                parameters30 = part1.Parameters
                length10 = parameters30.Item("Pillar_center_X")
                out_Guide_posts_X = length10.Value
                parameters31 = part1.Parameters
                length11 = parameters31.Item("Pillar_center_Y")
                out_Guide_posts_Y = length11.Value
                parameters32 = part1.Parameters
                length12 = parameters32.Item("Avoid_Error_X")
                out_Guide_posts_Avoid_Error_X = length12.Value
                parameters33 = part1.Parameters
                length13 = parameters33.Item("Pillar_a")
                out_Guide_posts_bolt_hole_X = length13.Value
                parameters34 = part1.Parameters
                length14 = parameters34.Item("Pillar_b")
                out_Guide_posts_bolt_hole_Y = length14.Value
                # ----------------------↑外導柱座標↑----------------------
                point_break = True
                break
            selection1.Search("Name = lower_die_" + total_plate + ",all")  # ---下模板
            label_determine_3 = selection1.Count
            if label_determine_3 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_pilot_punch_insert_0" + total_plate + " ,all")  # ---導引沖頭入子
            label_determine_4 = selection1.Count
            if label_determine_4 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cutting_cavity_d_insert_0" + total_plate + ",all")  # ---下靠刀沖頭入子
            label_determine_5 = selection1.Count
            if label_determine_5 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cutting_cavity_u_insert_0" + total_plate + ",all")  # ---上靠刀沖頭入子
            label_determine_6 = selection1.Count
            if label_determine_6 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cut_punch_d_cutting_add_0" + total_plate + ",all")  # ---下靠刀沖頭入塊
            label_determine_7 = selection1.Count
            if label_determine_7 > 0:
                point_break = True
                break
            selection1.Search(
                "Name = op" + total_op + "0_cut_punch_u_cutting_add_0" + total_plate + ",all")  # ---上靠刀沖頭入塊
            label_determine_8 = selection1.Count
            if label_determine_8 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_Stripper_insert_l_0" + total_plate + ",all")  # ---脫料入子
            label_determine_9 = selection1.Count
            if label_determine_9 > 0:
                point_break = True
                break
            selection1.Search("Name = lower_pad_" + total_plate + ",all")  # ---下墊板
            label_determine_10 = selection1.Count
            if label_determine_10 > 0:
                point_break = True
                break
            selection1.Search("Name = Stripper_" + total_plate + ",all")  # ---脫料板
            label_determine_11 = selection1.Count
            if label_determine_11 > 0:
                point_break = True
                break
            selection1.Search("Name = Stop_plate_" + total_plate + ",all")  # ---止擋板
            label_determine_12 = selection1.Count
            if label_determine_12 > 0:
                point_break = True
                break
            selection1.Search("Name = Splint_" + total_plate + ",all")  # ---上夾板
            label_determine_13 = selection1.Count
            if label_determine_13 > 0:
                point_break = True
                break
            selection1.Search("Name = up_plate_" + total_plate + ",all")  # ---上墊板
            label_determine_14 = selection1.Count
            if label_determine_14 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_cut_punch_d_cutting_" + total_plate + ",all")  # ---下靠刀沖頭
            label_determine_15 = selection1.Count
            if label_determine_15 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_cut_punch_u_cutting_" + total_plate + ",all")  # ---上靠刀沖頭
            label_determine_16 = selection1.Count
            if label_determine_16 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_forming_cavity_" + total_plate + ",all")  # ---下成形沖頭
            label_determine_17 = selection1.Count
            if label_determine_17 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_forming_punch_" + total_plate + ",all")  # --上成形沖頭
            label_determine_18 = selection1.Count
            if label_determine_18 > 0:
                point_break = True
                break
            selection1.Search("Name = pad_lower,all")  # ---墊腳
            label_determine_19 = selection1.Count
            if label_determine_19 > 0:
                point_break = True
                break
            selection1.Search("Name = op" + total_op + "0_QR_l_punch_0" + total_plate + ",all")  # ---靠肩沖頭
            label_determine_20 = selection1.Count
            if label_determine_20 > 0:
                point_break = True
                break
    selection1.Clear()
    part1.Update()
    # ================================↑孔標籤類型判斷↑===============================
    # =======================↓入子導引沖頭孔+螺栓孔座標(單一導引衝))↓=======================
    if label_determine_4 > 0:
        # -----------------↓導引沖頭孔↓-----------------
        body1 = bodies1.Item("PartBody")
        sketches1 = body1.Sketches
        originElements1 = part1.OriginElements
        reference1 = originElements1.PlaneXY
        sketch1 = sketches1.Add(reference1)
        arrayOfVariantOfDouble1 = [0, 0, 0, 1, 0, 0, 0, 1, 0]
        sketch1Variant = sketch1
        sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)
        part1.InWorkObject = sketch1
        factory2D1 = sketch1.OpenEdition()
        geometricElements1 = sketch1.GeometricElements
        axis2D1 = geometricElements1.Item("AbsoluteAxis")
        line2D1 = axis2D1.getItem("HDirection")
        line2D1.ReportName = 1
        line2D2 = axis2D1.getItem("VDirection")
        line2D2.ReportName = 2
        hybridShapes1 = body1.HybridShapes
        reference2 = hybridShapes1.Item("Y_min")
        geometricElements2 = factory2D1.CreateProjections(reference2)
        geometry2D1 = geometricElements2.Item("Mark.1")
        geometry2D1.Construction = True
        body2 = bodies1.Item("Body.2")
        hybridShapes2 = body2.HybridShapes
        reference3 = hybridShapes2.Item("Extremum.1(X_max)")
        geometricElements3 = factory2D1.CreateProjections(reference3)
        geometry2D2 = geometricElements3.Item("Mark.1")
        geometry2D2.Construction = True
        constraints1 = sketch1.Constraints
        reference4 = part1.CreateReferenceFromObject(geometry2D1)
        reference5 = part1.CreateReferenceFromObject(geometry2D2)
        reference6 = part1.CreateReferenceFromObject(line2D2)
        constraint1 = constraints1.AddTriEltCst(1, reference4, reference5, reference6)
        constraint1.mode = 1
        constraint1.Name = "pilot_punch_hole_Y_axis"
        reference7 = hybridShapes1.Item("X_min")
        geometricElements4 = factory2D1.CreateProjections(reference7)
        geometry2D3 = geometricElements4.Item("Mark.1")
        geometry2D3.Construction = True
        geometricElements5 = factory2D1.CreateProjections(reference3)
        geometry2D4 = geometricElements5.Item("Mark.1")
        geometry2D4.Construction = True
        reference8 = part1.CreateReferenceFromObject(geometry2D3)
        reference9 = part1.CreateReferenceFromObject(geometry2D4)
        reference10 = part1.CreateReferenceFromObject(line2D1)
        constraint2 = constraints1.AddTriEltCst(1, reference8, reference9, reference10)
        constraint2.mode = 1
        constraint2.Name = "pilot_punch_hole_X_axis"
        sketch1.CloseEdition()
        part1.InWorkObject = sketch1
        part1.InWorkObject.Name = "pilot_punch_label_Sketch"
        # -----------------↑導引沖頭孔↑-----------------
        # -------------------↓螺栓孔↓-------------------
        pilot_punch_insert_bolt_axis()
        # -------------------↑螺栓孔↑-------------------
        parameters21 = part1.Parameters
        length1 = parameters21.CreateDimension("", "LENGTH", 0)
        length1.rename("pilot_punch_hole_X_axis")
        parameters22 = part1.Parameters
        length2 = parameters22.CreateDimension("", "LENGTH", 0)
        length2.rename(
            "pilot_punch_hole_Y_axis")
        parameters23 = part1.Parameters
        length3 = parameters23.CreateDimension("", "LENGTH", 0)
        length3.rename(
            "bolt_hole_X_axis")
        parameters24 = part1.Parameters
        length4 = parameters24.CreateDimension("", "LENGTH", 0)
        length4.rename(
            "bolt_hole_Y_axis")
        part1.Update()
        sketches3 = body1.Sketches
        sketch3 = sketches3.Item("pilot_punch_label_Sketch")
        factory2D2 = sketch1.OpenEdition()
        relations1 = part1.Relations
        formula1 = relations1.Createformula("formula:pilot_punch_hole_X_axis", "", length1,
                                            "PartBody\\pilot_punch_label_Sketch\\pilot_punch_hole_X_axis\\Offset ")
        formula1.rename("formula_pilot_punch_hole_X_axis")
        relations2 = part1.Relations
        formula2 = relations2.Createformula("formula:pilot_punch_hole_Y_axis", "", length2,
                                            "PartBody\\pilot_punch_label_Sketch\\pilot_punch_hole_Y_axis\\Offset ")
        formula2.rename("formula_pilot_punch_hole_Y_axis")
        relations3 = part1.Relations
        formula3 = relations3.Createformula("formula:bolt_hole_X_axis", "", length3,
                                            "PartBody\\bolt_label_Sketch\\bolt_hole_X_axis\\Offset ")
        formula3.rename("formula_bolt_hole_X_axis")
        relations4 = part1.Relations
        formula4 = relations4.Createformula("formula:bolt_hole_X_axis", "", length4,
                                            "PartBody\\bolt_label_Sketch\\bolt_hole_Y_axis\\Offset ")
        formula4.rename("formula_bolt_hole_Y_axis")
        length = [None] * 5
        length[1] = parameters21.Item("pilot_punch_hole_X_axis")
        pilot_punch_hole_point_X = length[1].Value
        length[2] = parameters22.Item("pilot_punch_hole_Y_axis")
        pilot_punch_hole_point_Y = length[2].Value
        length[3] = parameters23.Item("bolt_hole_X_axis")
        pilot_punch_bolt_hole_point_X = length[3].Value
        length[4] = parameters24.Item("bolt_hole_Y_axis")
        pilot_punch_bolt_hole_point_Y = length[4].Value
        sketch1.CloseEdition()
    part1.Update()
    # =======================↑入子導引沖頭孔+螺栓孔座標(單一導引衝))↑=======================
    # =================================↓異形孔座標↓=================================
    # ------------------------------------↓建點↓------------------------------------
    part1 = partDocument1.Part
    selection_point_1 = partDocument1.Selection
    selection_point_2 = partDocument1.Selection
    selection_point_3 = partDocument1.Selection
    selection_point_4 = partDocument1.Selection
    selection_point_41 = partDocument1.Selection
    selection_point_5 = partDocument1.Selection
    selection_point_6 = partDocument1.Selection
    selection_point_7 = partDocument1.Selection
    selection_point_8 = partDocument1.Selection
    selection_point_9 = partDocument1.Selection
    selection_point_10 = partDocument1.Selection
    selection_point_11 = partDocument1.Selection
    selection_point_12 = partDocument1.Selection
    selection_point_13 = partDocument1.Selection
    for total_plate in range(1, 1 + 1):  # form7.Text8 #--------總模板數 *********(之後要分模板)***********
        total_plate = str(total_plate)
        for total_op in range(1, 1 + total_op_number):  # --------總工站數
            total_op = str(total_op)
            selection_point_1.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_punch_d_cutting_*_project_line,all")  # -----下靠刀沖頭孔
            plate_line_cut_punch_d_cutting_machining_shape = selection_point_1.Count
            if plate_line_cut_punch_d_cutting_machining_shape > 0:
                cut_punch_d_cutting_machining_shape_point(plate_line_cut_punch_d_cutting_machining_shape, total_op)
            # plate_line_cut_punch_d_cutting_machining_shape_number = 0 + plate_line_cut_punch_d_cutting_machining_shape #----------數量總和
            selection_point_2.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_punch_u_cutting_*_project_line,all")  # -----上靠刀沖頭孔
            plate_line_cut_punch_u_cutting_machining_shape = selection_point_2.Count
            if plate_line_cut_punch_u_cutting_machining_shape > 0:
                cut_punch_u_cutting_machining_shape_point(plate_line_cut_punch_u_cutting_machining_shape, total_op)
            selection_point_3.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_QR_l_punch_*_project_line,all")  # --------------右靠刀肩沖頭孔
            plate_line_QR_l_punch_machining_shape = selection_point_3.Count
            if plate_line_QR_l_punch_machining_shape > 0:
                QR_l_punch_machining_shape_point(plate_line_QR_l_punch_machining_shape, total_op)
            selection_point_4.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_line_*,all")  # ------------------------------剪切沖頭孔
            selection_point_next = selection_point_4.Count
            selection_point_41.Search("Name = plate_line_" + total_plate + "_op" + total_op + "0_cut_line_*_line,all")
            plate_line_cut_line_machining_shape = (selection_point_next - selection_point_41.Count)
            if plate_line_cut_line_machining_shape > 0:
                plate_line_cut_line_machining_shape_point(plate_line_cut_line_machining_shape, total_op)
            selection_point_5.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_pilot_punch_insert_*_project_line,all")  # -------脫料入子孔
            plate_line_pilot_punch_insert_machining_shape = selection_point_5.Count
            if plate_line_pilot_punch_insert_machining_shape > 0:
                plate_line_pilot_punch_insert_machining_shape_point(plate_line_pilot_punch_insert_machining_shape,
                                                                    total_op)
            selection_point_6.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_Stripper_insert_l_*_project_line,all")  # --------導引沖頭入子孔
            plate_line_Stripper_insert_left_machining_shape = selection_point_6.Count
            if plate_line_Stripper_insert_left_machining_shape > 0:
                plate_line_Stripper_insert_left_machining_shape_point(plate_line_Stripper_insert_left_machining_shape,
                                                                      total_op)
            selection_point_7.Search(
                "Name = plate_line" + total_plate + "_op" + total_op + "0_cutting_cavity_d_insert_*_project_line,all")  # ---下沖頭入子孔
            plate_line_cutting_cavity_d_insert_machining_shape = selection_point_7.Count
            if plate_line_cutting_cavity_d_insert_machining_shape > 0:
                plate_line_cutting_cavity_d_insert_machining_shape_point(
                    plate_line_cutting_cavity_d_insert_machining_shape, total_op)
            selection_point_8.Search(
                "Name = plate_line" + total_plate + "_op" + total_op + "0_cutting_cavity_u_insert_*_project_line,all")  # ---上沖頭入子孔
            plate_line_cutting_cavity_u_insert_machining_shape = selection_point_8.Count
            if plate_line_cutting_cavity_u_insert_machining_shape > 0:
                plate_line_cutting_cavity_u_insert_machining_shape_point(
                    plate_line_cutting_cavity_u_insert_machining_shape, total_op)
            selection_point_9.Search(
                "Name = plate_line_" + total_plate + "_op" + total_op + "0_forming_punch_Project_*,all")  # -----------------成形沖頭孔
            plate_line_forming_punch_machining_shape = selection_point_9.Count
            if plate_line_forming_punch_machining_shape > 0:
                plate_line_forming_punch_machining_shape_point(plate_line_forming_punch_machining_shape, total_op)
            selection_point_10.Search(
                "Name = op" + total_op + "0_cutting_cavity_d_insert_*,all")  # -----------------下靠刀入子沖頭孔(搜尋零件名稱)
            plate_line_cutting_cavity_d_insert_shape = selection_point_10.Count
            if plate_line_cutting_cavity_d_insert_shape > 0:
                plate_line_cutting_cavity_d_insert_shape_point(plate_line_cutting_cavity_d_insert_shape, total_op,
                                                               total_plate)
            selection_point_11.Search(
                "Name = op" + total_op + "0_cutting_cavity_u_insert_*,all")  # -----------------上靠刀入子沖頭孔(搜尋零件名稱)
            plate_line_cutting_cavity_u_insert_shape = selection_point_11.Count
            if plate_line_cutting_cavity_u_insert_shape > 0:
                plate_line_cutting_cavity_u_insert_shape_point(plate_line_cutting_cavity_u_insert_shape, total_op,
                                                               total_plate)
            selection_point_12.Search(
                "Name = op" + total_op + "0_Stripper_insert_l_*,all")  # -----------------下料沖頭入子沖孔(搜尋零件名稱)
            plate_line_Stripper_QR_l_punch_shape = selection_point_12.Count
            if plate_line_Stripper_QR_l_punch_shape > 0:
                plate_line_Stripper_QR_l_punch_shape_point(plate_line_Stripper_QR_l_punch_shape, total_op, total_plate)
            # -----------------↓例外(stop_plate),op_70_QR_l_punch_1_project更改名稱?
            selection_point_13.Search("Name = op_" + total_op + "0_QR_l_punch_*_project,all")  # 右靠肩沖頭孔(例外)
            QR_l_punch_machining_shape = selection_point_13.Count
            if QR_l_punch_machining_shape > 0:
                stop_plate_QR_l_punch_machining_shape_point(QR_l_punch_machining_shape, total_op)
            # -----------------↑例外(stop_plate),op_70_QR_l_punch_1_project更改名稱?
    part1.Update()
    # ------------------------------------↓座標↓------------------------------------
    selection_axis_1 = partDocument1.Selection
    selection_axis_1.Search("Name = machining_shape_point_*,all")
    machining_shape_point_number = selection1.Count
    if machining_shape_point_number > 0:
        (machining_shape_point_X, machining_shape_point_Y) = machining_shape_point_axis(partname,
                                                                                        machining_shape_point_number)
    # =================================↑異形孔座標↑================================
    # ================================================↓出圖↓=================================================
    drawingDocument1 = documents1.Open(gvar.open_path + "A0.CATDrawing")  # -------------開啟圖紙(input).
    drawingSheets1 = drawingDocument1.Sheets
    Scaleradio = 1
    viewdistance = 800  # -------------------------視圖間距離
    mainviewx = 430  # -------------------------主視圖X座標
    mainviewy = 526  # -------------------------主視圖Y座標
    topviewy = mainviewy + viewdistance / 3.4
    rightviewx = mainviewx + viewdistance / 2.52
    downviewy = mainviewy - viewdistance / 3.4
    leftview = mainviewx - viewdistance / 2.55
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    MyText1 = MyView.Texts.Add("＊如果對於圖面有更好的建議歡迎提出,我們會虛心接受", 44, 117)
    MyText1.SetFontSize(0, 0, 20)
    MyText2 = MyView.Texts.Add("未 注 公 差", 70, 91)
    MyText2.SetFontSize(0, 0, 20)
    MyText3 = MyView.Texts.Add("角度:", 152, 54)
    MyText3.SetFontSize(0, 0, 12)
    MyText4 = MyView.Texts.Add("圖名", 218, 90)
    MyText4.SetFontSize(0, 0, 17)
    MyText5 = MyView.Texts.Add("圖號", 218, 66)
    MyText5.SetFontSize(0, 0, 17)
    MyText6 = MyView.Texts.Add("路徑", 218, 43)
    MyText6.SetFontSize(0, 0, 17)
    MyText7 = MyView.Texts.Add("材  質", 372, 90)
    MyText7.SetFontSize(0, 0, 17)
    MyText8 = MyView.Texts.Add("熱處理", 372, 66)
    MyText8.SetFontSize(0, 0, 17)
    MyText9 = MyView.Texts.Add("板  厚", 372, 43)
    MyText9.SetFontSize(0, 0, 17)
    MyText10 = MyView.Texts.Add("圖檔比例", 483, 90)
    MyText10.SetFontSize(0, 0, 17)
    MyText12 = MyView.Texts.Add("投影方向", 483, 43)
    MyText12.SetFontSize(0, 0, 17)
    MyText13 = MyView.Texts.Add("設計", 622, 90)
    MyText13.SetFontSize(0, 0, 17)
    MyText13 = MyView.Texts.Add("檢查", 622, 66)
    MyText13.SetFontSize(0, 0, 17)
    MyText14 = MyView.Texts.Add("認可", 622, 43)
    MyText14.SetFontSize(0, 0, 17)
    MyText15 = MyView.Texts.Add("圖發部門", 790, 90)
    MyText15.SetFontSize(0, 0, 17)
    MyText16 = MyView.Texts.Add("客產編號", 790, 66)
    MyText16.SetFontSize(0, 0, 17)
    MyText17 = MyView.Texts.Add("圖印時間", 790, 43)
    MyText17.SetFontSize(0, 0, 17)
    MyText18 = MyView.Texts.Add("頁碼", 995, 43)
    MyText18.SetFontSize(0, 0, 17)
    MyText19 = MyView.Texts.Add("第    頁 ,", 1030, 43)
    MyText19.SetFontSize(0, 0, 15)
    MyText20 = MyView.Texts.Add("共    頁", 1108, 43)
    MyText20.SetFontSize(0, 0, 15)
    MyText21 = MyView.Texts.Add("金屬產品開發研究發展中心", 1034, 70)
    MyText21.SetFontSize(0, 0, 12)
    MyText22 = MyView.Texts.Add(" M P R D C", 1060, 85)
    MyText22.SetFontSize(0, 0, 12)
    MyText22 = MyView.Texts.Add("圖面版本記錄", 1010, 805)
    MyText22.SetFontSize(0, 0, 17)
    MyText23 = MyView.Texts.Add("版本", 933, 781)
    MyText23.SetFontSize(0, 0, 17)
    MyText24 = MyView.Texts.Add("版本說明", 975, 781)
    MyText24.SetFontSize(0, 0, 17)
    MyText25 = MyView.Texts.Add("日  期", 1050, 781)
    MyText25.SetFontSize(0, 0, 17)
    MyText26 = MyView.Texts.Add("設  計", 1114, 781)
    MyText26.SetFontSize(0, 0, 17)
    # -----------------------------------------↓加工說明↓---------------------------------------
    MyText62 = MyView.Texts.Add("加工說明: (" + partname + ")", 916, 710)  # ---零件名稱
    MyText62.SetFontSize(0, 0, 13.5)
    machining_explanation_X = 916  # ------整個加工說明X座標
    machining_explanation_Y = 710  # ------整個加工說明Y座標
    machining_explanation_P = 0  # ------加工說明每行間距
    for M in range(1, 1 + 18):
        if machining_explanation[M] == None:
            break
        # ------------------------------------分割字串，一行26個字----------------------------------
        machining_explanation[M] = str(machining_explanation[M])
        MyLen = len(machining_explanation[M])  # 傳回字串中字元個數。
        if MyLen > 26:
            words = -int(-MyLen / 26)
            w = [""] * (words + 1)
            machining_explanation_temporary = machining_explanation[M][0:26]
            for i in range(1, words + 1):
                w[i] = machining_explanation[M][(i * 26) + 1: (i * 26) + 27]  # ----(i*26)+1:下一行從第27個字開始, 26:一行26個字
                machining_explanation_temporary = machining_explanation_temporary + "\n" + w[i]
            machining_explanation[M] = machining_explanation_temporary
        # ------------------------------------分割字串，一行26個字----------------------------------
        machining_explanation_P = machining_explanation_P + 22
        MyText63 = MyView.Texts.Add(machining_explanation[M], str(machining_explanation_X),
                                    str(machining_explanation_Y - machining_explanation_P))  # ---內容
        MyText63.SetFontSize(0, 0, 12)
        if MyLen > 26:
            machining_explanation_Y = machining_explanation_Y - 12  # -------字高
    # ---------------------------------A0常修改參數表格內容----------------------------------------------
    Frame_1 = "X.:"
    Frame_2 = "X.X:"
    Frame_3 = "+"
    Frame_4 = "-"
    Frame_5 = "+"
    Frame_6 = "-"
    Frame_7 = " "
    Frame_8 = " "
    Frame_9 = " "
    Frame_10 = " "
    Frame_11 = "X.XX:"
    Frame_12 = "X.XXX:"
    Frame_13 = "±"
    Frame_14 = "±"
    Frame_15 = "0.05"
    Frame_16 = "0.005"
    Frame_17 = "±"
    Frame_18 = "°"
    MyText27 = MyView.Texts.Add(Frame_1, 34, 63)
    MyText27.SetFontSize(0, 0, 11)
    MyText28 = MyView.Texts.Add(Frame_2, 34, 40)
    MyText28.SetFontSize(0, 0, 11)
    MyText29 = MyView.Texts.Add(Frame_11, 85, 63)
    MyText29.SetFontSize(0, 0, 11)
    MyText30 = MyView.Texts.Add(Frame_12, 85, 40)
    MyText30.SetFontSize(0, 0, 11)
    MyText31 = MyView.Texts.Add(Frame_3 + Frame_7, 60, 66)
    MyText31.SetFontSize(0, 0, 9)
    MyText32 = MyView.Texts.Add(Frame_4 + Frame_8, 60, 56)
    MyText32.SetFontSize(0, 0, 9)
    MyText33 = MyView.Texts.Add(Frame_5 + Frame_9, 60, 43)
    MyText33.SetFontSize(0, 0, 9)
    MyText34 = MyView.Texts.Add(Frame_6 + Frame_10, 60, 33)
    MyText34.SetFontSize(0, 0, 9)
    MyText35 = MyView.Texts.Add(Frame_13 + Frame_15, 117, 63)
    MyText35.SetFontSize(0, 0, 10)
    MyText36 = MyView.Texts.Add(Frame_14 + Frame_16, 117, 40)
    MyText36.SetFontSize(0, 0, 10)
    MyText37 = MyView.Texts.Add(Frame_17 + Frame_18, 176, 54)
    MyText37.SetFontSize(0, 0, 12)
    MyText38 = MyView.Texts.Add(partname, 256, 90)
    MyText38.SetFontSize(0, 0, 17)
    MyText39 = MyView.Texts.Add("USA035", 256, 66)  # 圖號
    MyText39.SetFontSize(0, 0, 17)
    MyText40 = MyView.Texts.Add("D:\小慈\實驗室", 256, 43)  # 路徑
    MyText40.SetFontSize(0, 0, 16)
    MyText41 = MyView.Texts.Add(Frame_Material, 428, 90)  # 自動材質
    MyText41.SetFontSize(0, 0, 17)
    MyText42 = MyView.Texts.Add(Frame_Heat_treatment, 428, 66)  # 自動熱處理
    MyText42.SetFontSize(0, 0, 17)
    MyText43 = MyView.Texts.Add(Frame_Thickness_1, 428, 43)  # 自動板厚
    MyText43.SetFontSize(0, 0, 17)
    MyText44 = MyView.Texts.Add("1：1", 555, 90)  # 圖檔比例
    MyText44.SetFontSize(0, 0, 17)
    MyText46 = MyView.Texts.Add("第三角", 555, 43)  # 投影方向
    MyText46.SetFontSize(0, 0, 17)
    MyText47 = MyView.Texts.Add(" ", 662, 90)  # 設計專員
    MyText47.SetFontSize(0, 0, 17)
    MyText48 = MyView.Texts.Add(" ", 721, 90)  # 設計專員
    MyText48.SetFontSize(0, 0, 17)
    MyText49 = MyView.Texts.Add(" ", 662, 66)  # 檢查專員
    MyText49.SetFontSize(0, 0, 17)
    MyText50 = MyView.Texts.Add(" ", 721, 66)  # 檢查專員
    MyText50.SetFontSize(0, 0, 17)
    MyText51 = MyView.Texts.Add(" ", 662, 43)  # 認可專員
    MyText51.SetFontSize(0, 0, 17)
    MyText52 = MyView.Texts.Add(" ", 721, 43)  # 認可專員
    MyText52.SetFontSize(0, 0, 17)
    MyText53 = MyView.Texts.Add("研發部", 859, 90)  # 圖發部門
    MyText53.SetFontSize(0, 0, 17)
    MyText54 = MyView.Texts.Add("654452", 859, 66)  # 客產編號
    MyText54.SetFontSize(0, 0, 17)
    now_time = time.strftime('%Y/%m/%d', time.localtime())
    MyText55 = MyView.Texts.Add(now_time, 859, 43)  # 圖印日期
    MyText55.SetFontSize(0, 0, 17)
    MyText56 = MyView.Texts.Add(drafting_page, 1056, 43)
    MyText56.SetFontSize(0, 0, 15)
    MyText57 = MyView.Texts.Add(drafting_total_page, 1134, 43)
    MyText57.SetFontSize(0, 0, 15)
    MyText58 = MyView.Texts.Add(" ", 933, 758)  # 版本
    MyText58.SetFontSize(0, 0, 17)
    MyText59 = MyView.Texts.Add(" ", 975, 758)  # 版本說明
    MyText59.SetFontSize(0, 0, 17)
    MyText60 = MyView.Texts.Add(now_time, 1050, 758)  # 日期
    MyText60.SetFontSize(0, 0, 15)
    MyText61 = MyView.Texts.Add("專案人員", 1114, 758)  # 設計人員
    MyText61.SetFontSize(0, 0, 15)
    # ---------------------------------------------------------↓投影↓---------------------------------------------------
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    product1 = partDocument1.getItem("Part1")  # 零件名稱
    drawingViewGenerativeBehavior1.Document = product1
    drawingViewGenerativeBehavior1.DefineFrontView(1, 0, 0, 0, 1, 0)
    drawingView1.X = mainviewx
    drawingView1.Y = mainviewy
    drawingView1.Scale = Scaleradio
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    drawingViewGenerativeBehavior1.Update()
    drawingView1.Activate()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingSheet1.ProjectionMethod = 1
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    # ---------------------------------↓存svg檔↓---------------------------------------------
    if "lower_die_" in partname or "lower_pad_" in partname or "upper_die_set" in partname or "Stripper_" in partname or "Splint_" in partname or "Stop_plate_" in partname or "up_plate_" in partname :
        if "insert" not in partname:
            svg(partname)
    # ---------------------------------↑存svg檔↑---------------------------------------------
    # ---------------------------------右視圖-------------------------------------------------------------
    drawingView2 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingViewGenerativeBehavior2.DefineProjectionView(drawingViewGenerativeBehavior1, 0)
    drawingViewGenerativeLinks1 = drawingView2.GenerativeLinks
    drawingViewGenerativeLinks2 = drawingView1.GenerativeLinks
    drawingViewGenerativeLinks2.CopyLinksTo(drawingViewGenerativeLinks1)
    drawingView2.X = rightviewx
    drawingView2.Y = mainviewy
    double1 = drawingView1.Scale
    drawingView2.Scale = Scaleradio
    drawingViewGenerativeBehavior2 = drawingView2.GenerativeBehavior
    drawingViewGenerativeBehavior2.Update()
    drawingView2.ReferenceView = drawingView1
    drawingView2.AlignedWithReferenceView()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    # ---------------------------------下視圖-------------------------------------------------------------
    drawingView4 = drawingViews1.Add("AutomaticNaming")
    drawingViewGenerativeBehavior4 = drawingView4.GenerativeBehavior
    drawingViewGenerativeBehavior4.DefineProjectionView(drawingViewGenerativeBehavior1, 3)
    drawingViewGenerativeLinks4 = drawingView4.GenerativeLinks
    drawingViewGenerativeLinks2 = drawingView1.GenerativeLinks
    drawingViewGenerativeLinks2.CopyLinksTo(drawingViewGenerativeLinks4)
    drawingView4.X = mainviewx
    drawingView4.Y = downviewy
    double3 = drawingView1.Scale
    drawingView4.Scale = Scaleradio
    drawingViewGenerativeBehavior4 = drawingView4.GenerativeBehavior
    drawingViewGenerativeBehavior4.Update()
    drawingView4.ReferenceView = drawingView1
    drawingView4.AlignedWithReferenceView()
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.ActiveSheet
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.ActiveView
    drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior
    # ================↓切換孔標籤模組↓================
    if label_determine_1 > 0:  # 下模座
        lower_die_set_label(machining_shape_point_number, mainviewx, Length_max, mainviewy, Width_max,
                            machining_shape_point_X, machining_shape_point_Y, out_Guide_posts_X, out_Guide_posts_Y,
                            out_Guide_posts_bolt_hole_X, out_Guide_posts_bolt_hole_Y, out_Guide_posts_Avoid_Error_X)
    if label_determine_2 > 0:  # 上模座
        upper_die_set_label()
    if label_determine_3 > 0:  # 下模板
        lower_die_label(machining_shape_point_number, mainviewx, Length_max, mainviewy, Width_max,
                        machining_shape_point_X,
                        machining_shape_point_Y)
    if label_determine_4 > 0:  # 導引沖入子
        pilot_punch_insert_label()
    if label_determine_5 > 0:  # 下靠刀沖頭入子
        cutting_cavity_d_insert_label(machining_explanation, machining_shape_point_number, mainviewx, mainviewy,
                                      Length_max,
                                      machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_6 > 0:  # 上靠刀沖頭入子
        cutting_cavity_u_insert_label(machining_explanation, machining_shape_point_number, mainviewx, mainviewy,
                                      Length_max,
                                      machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_10 > 0:  # 下背板
        lower_pad_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                        machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_11 > 0:  # 脫料板
        Stripper_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                       machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_12 > 0:  # 止擋板
        Stop_plate_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                         machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_13 > 0:  # 上夾板
        Splint_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                     machining_shape_point_X, Width_max, machining_shape_point_Y)
    if label_determine_14 > 0:  # 上墊板
        up_plate_label()
    # ================↑切換孔標籤模組↑================
    # ============================================↓存檔↓============================================
    time.sleep(0.5)
    partDocument1.save()
    partDocument1.Close()
    drawingDocument1.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".CATDrawing",
                                "CATDrawing")  # 更新儲存路徑(2D output)
    drawingDocument1.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".dwg", "dwg")  # 使用dwg存檔
    drawingDocument1.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".jpg", "jpg")  # 使用JPG存檔
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()
    drawingDocument1 = catapp.ActiveDocument
    drawingDocument1.Close()


def A00(partname, Height_max, Length_max, Width_max, Frame_Thickness_1, drafting_page, drafting_total_page):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    partDocument1.Close()
    pass


def Momo_machining_explanation(partname):  # 確認材料與熱處理
    # Frame_Material  自動材料
    # Frame_Heat_treatment  自動熱處理
    if partname == "upper_die_set":  # 上模座
        Frame_Material = gvar.strip_parameter_list[6]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[7]  # 熱處理
    if partname == "lower_die_set":  # 下模座
        Frame_Material = gvar.strip_parameter_list[33]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[34]  # 熱處理
    if "_cut_punch_" in partname:  # 沖頭
        Frame_Material = gvar.strip_parameter_list[37]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[38]  # 熱處理
    # ===================A沖=====================
    if "_SJAS_" in partname:
        Frame_Material = "SJAS"  # 材料
        Frame_Heat_treatment = "SJAS"  # 熱處理
    if "_SJAL_" in partname:
        Frame_Material = "SJAL"  # 材料
        Frame_Heat_treatment = "SJAL"  # 熱處理
    if "_A_SJAS_" in partname:
        Frame_Material = "A_SJAS"  # 材料
        Frame_Heat_treatment = "A_SJAS"  # 熱處理
    if "_A_SJAL_" in partname:
        Frame_Material = "A_SJAL"  # 材料
        Frame_Heat_treatment = "A_SJAL"  # 熱處理
    # ===========================================
    if "_cut_cavity_insert_" in partname or "_A_punch_insert_" in partname or "_A_punch_QR_Splint_insert_" in partname or "_A_punch_QR_Stripper_insert_" in partname:  # 入子
        Frame_Material = gvar.strip_parameter_list[35]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[36]  # 熱處理
    if "lower_die_" in partname:  # 下模板
        Frame_Material = gvar.strip_parameter_list[27]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[28]  # 熱處理
    if "lower_pad_" in partname:  # 下墊板
        Frame_Material = gvar.strip_parameter_list[30]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[31]  # 熱處理
    if "Stripper_" in partname:  # 脫料板
        Frame_Material = gvar.strip_parameter_list[21]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[22]  # 熱處理
    if "Stop_plate_" in partname:  # 止擋板
        Frame_Material = gvar.strip_parameter_list[18]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[19]  # 熱處理
    if "Splint_" in partname:  # 上夾板
        Frame_Material = gvar.strip_parameter_list[15]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[16]  # 熱處理
    if "up_plate_" in partname:  # 上墊板
        Frame_Material = gvar.strip_parameter_list[12]  # 材料
        Frame_Heat_treatment = gvar.strip_parameter_list[13]  # 熱處理
    return Frame_Material, Frame_Heat_treatment


def pilot_punch_insert_bolt_axis():
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    sketches1 = body1.Sketches
    sketch1 = sketches1.Item("Sketch.9")
    reference1 = part1.CreateReferenceFromObject(sketch1)
    hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
    hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
    hybridShapeExtremum1.Direction2 = hybridShapeDirection2
    hybridShapeExtremum1.ExtremumType2 = 1
    body2 = bodies1.Item("PartBody")
    body2.InsertHybridShape(hybridShapeExtremum1)
    part1.InWorkObject = hybridShapeExtremum1
    part1.InWorkObject.Name = "pilot_punch_insert_bolt_point"
    part1.Update()
    part1.Update()
    sketches2 = body2.Sketches
    originElements1 = part1.OriginElements
    reference2 = originElements1.PlaneXY
    sketch2 = sketches2.Add(reference2)
    arrayOfVariantOfDouble1 = [0, 0, 0, 1, 0, 0, 0, 1, 0]
    sketch2Variant = sketch2
    sketch2Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)
    part1.InWorkObject = sketch2
    factory2D1 = sketch2.OpenEdition()
    geometricElements1 = sketch2.GeometricElements
    axis2D1 = geometricElements1.Item("AbsoluteAxis")
    line2D1 = axis2D1.getItem("HDirection")
    line2D1.ReportName = 1
    line2D2 = axis2D1.getItem("VDirection")
    line2D2.ReportName = 2
    hybridShapes1 = body2.HybridShapes
    reference3 = hybridShapes1.Item("pilot_punch_insert_bolt_point")
    geometricElements2 = factory2D1.CreateProjections(reference3)
    geometry2D1 = geometricElements2.Item("Mark.1")
    geometry2D1.Construction = True
    reference4 = hybridShapes1.Item("X_min")
    geometricElements3 = factory2D1.CreateProjections(reference4)
    geometry2D2 = geometricElements3.Item("Mark.1")
    geometry2D2.Construction = True
    constraints1 = sketch2.Constraints
    reference5 = part1.CreateReferenceFromObject(geometry2D1)
    reference6 = part1.CreateReferenceFromObject(geometry2D2)
    reference7 = part1.CreateReferenceFromObject(line2D1)
    constraint1 = constraints1.AddTriEltCst(1, reference5, reference6, reference7)
    constraint1.mode = 1
    constraint1.Name = "bolt_hole_X_axis"
    geometricElements4 = factory2D1.CreateProjections(reference3)
    geometry2D3 = geometricElements4.Item("Mark.1")
    geometry2D3.Construction = True
    reference8 = hybridShapes1.Item("Y_min")
    geometricElements5 = factory2D1.CreateProjections(reference8)
    geometry2D4 = geometricElements5.Item("Mark.1")
    geometry2D4.Construction = True
    reference9 = part1.CreateReferenceFromObject(geometry2D3)
    reference10 = part1.CreateReferenceFromObject(geometry2D4)
    reference11 = part1.CreateReferenceFromObject(line2D2)
    constraint2 = constraints1.AddTriEltCst(1, reference9, reference10, reference11)
    constraint2.mode = 1
    constraint2.Name = "bolt_hole_Y_axis"
    sketch2.CloseEdition()
    part1.InWorkObject = sketch2
    part1.InWorkObject.Name = "bolt_label_Sketch"
    part1.Update()


def cut_punch_d_cutting_machining_shape_point(plate_line_cut_punch_d_cutting_machining_shape, total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_cut_punch_d_cutting_machining_shape):  # 線段總數
        line = str(line)
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_1_op" + str(total_op) + "0_cut_punch_d_cutting_" + str(line) + "_project_line")
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + line
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + line
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + line + "1"
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + line
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def cut_punch_u_cutting_machining_shape_point(plate_line_cut_punch_u_cutting_machining_shape, total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_cut_punch_u_cutting_machining_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_1_op" + str(total_op) + "0_cut_punch_u_cutting_" + str(line) + "_project_line")
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line) + "1"
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + machining_shape_point_number + 1


def QR_l_punch_machining_shape_point(plate_line_QR_l_punch_machining_shape, total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_QR_l_punch_machining_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_1_op" + str(total_op) + "0_QR_l_punch_" + str(line) + "_project_line")
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line) + "1"
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + machining_shape_point_number + 1


def stop_plate_QR_l_punch_machining_shape_point(QR_l_punch_machining_shape, total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + QR_l_punch_machining_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("op_" + str(total_op) + "0_QR_l_punch_" + str(line) + "_boundary")
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + machining_shape_point_number + 1


def plate_line_cut_line_machining_shape_point(plate_line_cut_line_machining_shape, total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_cut_line_machining_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item("plate_line_1_op" + str(total_op) + "0_cut_line_" + str(line))
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def plate_line_Stripper_insert_left_machining_shape_point(plate_line_Stripper_insert_l_machining_shape, total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_Stripper_insert_l_machining_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_1_op" + str(total_op) + "0_Stripper_insert_l_0" + str(line) + "_project_line")
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def plate_line_pilot_punch_insert_machining_shape_point(plate_line_pilot_punch_insert_machining_shape, total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_pilot_punch_insert_machining_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_1_op" + str(total_op) + "0_pilot_punch_insert_0" + str(line) + "_project_line")
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def plate_line_cutting_cavity_d_insert_machining_shape_point(plate_line_cutting_cavity_d_insert_machining_shape,
                                                             total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_cutting_cavity_d_insert_machining_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_1_op" + str(total_op) + "0_cutting_cavity_d_insert_0" + str(line) + "_project_line")
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def plate_line_cutting_cavity_u_insert_machining_shape_point(plate_line_cutting_cavity_u_insert_machining_shape,
                                                             total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_cutting_cavity_u_insert_machining_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_1_op" + str(total_op) + "0_cutting_cavity_u_insert_" + str(line) + "_project_line")
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def plate_line_forming_punch_machining_shape_point(plate_line_forming_punch_machining_shape, total_op):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_forming_punch_machining_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_1_op" + str(total_op) + "0_forming_punch_Project_" + str(line))
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Searc("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def plate_line_cutting_cavity_d_insert_shape_point(plate_line_cutting_cavity_d_insert_shape, total_op, total_plate):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_cutting_cavity_d_insert_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_" + str(total_plate) + "_op" + str(total_op) + "0_cut_punch_d_cutting_" + str(line))
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def plate_line_cutting_cavity_u_insert_shape_point(plate_line_cutting_cavity_u_insert_shape, total_op, total_plate):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_cutting_cavity_u_insert_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_" + total_plate + "_op" + str(total_op) + "0_cut_punch_u_cutting_" + str(line))
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def plate_line_Stripper_QR_l_punch_shape_point(plate_line_Stripper_QR_l_punch_shape, total_op, total_plate):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    for line in range(1, 1 + plate_line_Stripper_QR_l_punch_shape):  # 線段總數
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_" + total_plate + "_op" + str(total_op) + "0_QR_l_punch_" + str(line))
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 1
        body1 = bodies1.Item("PartBody")
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        selection1 = partDocument1.Selection
        selection1.Search("Name =machining_shape_point_*,all")
        machining_shape_point_number = selection1.Count
        if machining_shape_point_number == 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number == 1 and line == 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line+1)
        if machining_shape_point_number == 1 and line > 1:
            part1.InWorkObject.Name = "machining_shape_point_" + str(line)
        if machining_shape_point_number != 1 and machining_shape_point_number != 0:
            part1.InWorkObject.Name = "machining_shape_point_" + str(machining_shape_point_number + 1)


def machining_shape_point_axis(partname, machining_shape_point_number):
    catapp = win32.Dispatch('CATIA.Application')
    # =============================↓異形孔標註基準點↓=============================
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")  # --------------定義工作物件
    part1.InWorkObject = body1
    hybridShapeFactory1 = part1.HybridShapeFactory
    hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
    machining_shape_point_X = [0.0] * (machining_shape_point_number + 1)
    machining_shape_point_Y = [0.0] * (machining_shape_point_number + 1)
    # ------------↓選擇工作物件↓------------
    if partname == "lower_die_set":
        partbody_select = "PartBody"
    elif partname == "upper_die_set":
        partbody_select = "PartBody"
    elif partname == "Part1":
        partbody_select = "PartBody"
    elif partname == "op70_bending_punch_01":
        partbody_select = "PartBody"
    elif partname == "op70_bending_punch_02":
        partbody_select = "PartBody"
    elif partname == "op70_bending_punch_03":
        partbody_select = "PartBody"
    elif partname == "op70_bending_punch_up01":
        partbody_select = "PartBody"
    elif partname == "op70_bending_punch_up02":
        partbody_select = "PartBody"
    elif partname == "op70_bending_punch_up03":
        partbody_select = "PartBody"
    else:
        partbody_select = "Body.2"
    # -----------------↓建座標準點↓----------------
    body2 = bodies1.Item(partbody_select)  # -------定義工作物件
    shapes1 = body2.Shapes
    if partname == "lower_die_set":
        pad1 = shapes1.Item("Pad.6")
    elif partname == "upper_die_set":
        pad1 = shapes1.Item("Pad.6")
    else:
        pad1 = shapes1.Item("Pad.3")
    reference1 = part1.CreateReferenceFromObject(pad1)
    hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 0)
    hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
    hybridShapeExtremum1.Direction2 = hybridShapeDirection2
    hybridShapeExtremum1.ExtremumType2 = 0
    hybridShapeDirection3 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)
    hybridShapeExtremum1.Direction3 = hybridShapeDirection3
    hybridShapeExtremum1.ExtremumType3 = 0
    body2 = bodies1.Item("PartBody")
    body2.InsertHybridShape(hybridShapeExtremum1)
    part1.InWorkObject = hybridShapeExtremum1
    part1.Update()
    reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
    hybridShapePointExplicit1 = hybridShapeFactory1.AddNewPointDatum(reference2)
    body2.InsertHybridShape(hybridShapePointExplicit1)
    part1.InWorkObject = hybridShapePointExplicit1
    part1.InWorkObject.Name = "machining_shape_point_axis_base"
    part1.Update()
    hybridShapeFactory1.DeleteObjectforDatum(reference2)
    part1.Update()
    # =============================↑異形孔標註基準點↑=============================
    # ================================↓異形孔標註↓================================
    body1 = bodies1.Item("PartBody")
    sketches1 = body1.Sketches
    originElements1 = part1.OriginElements
    reference1 = originElements1.PlaneXY  # XY平面建草圖
    sketch1 = sketches1.Add(reference1)
    arrayOfVariantOfDouble1 = [0, 0, 0, 1, 0, 0, 0, 1, 0]
    sketch1Variant = sketch1
    sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)
    part1.InWorkObject = sketch1
    factory2D1 = sketch1.OpenEdition()
    geometricElements1 = sketch1.GeometricElements
    for line in range(1, 1 + machining_shape_point_number):  # 線段總數
        axis2D1 = geometricElements1.Item("AbsoluteAxis")
        line2D1 = axis2D1.getItem("HDirection")
        line2D1.ReportName = 1
        line2D2 = axis2D1.getItem("VDirection")
        line2D2.ReportName = 2
        hybridShapes1 = body1.HybridShapes
        print("machining_shape_point_" + str(line))
        reference2 = hybridShapes1.Item("machining_shape_point_" + str(line))
        geometricElements2 = factory2D1.CreateProjections(reference2)
        geometry2D1 = geometricElements2.Item("Mark.1")
        geometry2D1.Construction = True
        reference3 = hybridShapes1.Item("machining_shape_point_axis_base")
        geometricElements3 = factory2D1.CreateProjections(reference3)
        geometry2D2 = geometricElements3.Item("Mark.1")
        geometry2D2.Construction = True
        constraints1 = sketch1.Constraints
        reference4 = part1.CreateReferenceFromObject(geometry2D1)
        reference5 = part1.CreateReferenceFromObject(geometry2D2)
        reference6 = part1.CreateReferenceFromObject(line2D1)
        constraint1 = constraints1.AddTriEltCst(1, reference4, reference5, reference6)
        constraint1.mode = 1
        constraint1.Name = "X_axis_" + str(line)  # 更改X方向標註名稱
        geometricElements4 = factory2D1.CreateProjections(reference2)
        geometry2D3 = geometricElements4.Item("Mark.1")
        geometry2D3.Construction = True
        geometricElements5 = factory2D1.CreateProjections(reference3)
        geometry2D4 = geometricElements5.Item("Mark.1")
        geometry2D4.Construction = True
        reference7 = part1.CreateReferenceFromObject(geometry2D3)
        reference8 = part1.CreateReferenceFromObject(geometry2D4)
        reference9 = part1.CreateReferenceFromObject(line2D2)
        constraint2 = constraints1.AddTriEltCst(1, reference7, reference8, reference9)
        constraint2.mode = 1
        constraint2.Name = "Y_axis_" + str(line)  # 更改X方向標註名稱
    part1.InWorkObject.Name = "machining_shape_point_Sketch"  # 更改草圖名稱
    # ---------------------------↓建參數↓-------------------------------
    for line in range(1, 1 + machining_shape_point_number):  # 線段總數
        parameters1 = part1.Parameters
        length1 = parameters1.CreateDimension("", "LENGTH", 0)
        length1.rename("machining_shape_point_X_axis_" + str(line) + "")
        parameters2 = part1.Parameters
        length2 = parameters2.CreateDimension("", "LENGTH", 0)
        length2.rename("machining_shape_point_Y_axis_" + str(line) + "")
        relations1 = part1.Relations
        formula1 = relations1.Createformula("formula.1", "", length1,
                                            "PartBody\\machining_shape_point_Sketch\\X_axis_" + str(line) + "\\Offset ")
        formula1.rename("formula_machining_shape_point_X_axis_" + str(line) + "")
        relations2 = part1.Relations
        formula2 = relations2.Createformula("formula.2", "", length2,
                                            "PartBody\\machining_shape_point_Sketch\\Y_axis_" + str(line) + "\\Offset ")
        formula2.rename("formula_machining_shape_point_Y_axis_" + str(line) + "")
        machining_shape_point_X[line] = length1.Value  # 異形孔X座標
        machining_shape_point_Y[line] = length2.Value  # 異形孔Y座標
        # --------------------------↓座標存取↓---------------------------
    sketch1.CloseEdition()
    part1.Update()
    # ================================↑異形孔標註↑================================
    return machining_shape_point_X, machining_shape_point_Y


def svg(partname):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    drawingdocument = catapp.ActiveDocument
    drawingsheets = drawingdocument.Sheets
    drawingsheet = drawingsheets.Item('Sheet.1')
    drawingviews = drawingsheet.Views
    drawingview = drawingviews.Item('Front view')
    drawingview.Activate()
    # ----------------------------------------畫圓
    # hole_op10_pick = [[-20, -32.45, 6.64], [-20, 1.45, 6.647], [40, -12.584, 6.647],
    #                   [80, -25.505, 6.647], [200.007, -32.469, 6.647], [240, -33.5, 6.647]]  # [x,y]
    # Lower_pad_bolt_spacer = [[-20, -69.7, 13], [130, -69.7, 13], [-20, 69.3, 13], [130, 69.3, 13]]
    # Lower_pad_pin = [[60, -247, 10.5], [-30, 24.3, 10.5], [190, 24.3, -24.7], [190, 24.3, 10.5]]
    # Lower_pad_Innerguide = [[-50, -54.7, 22], [290, -54.7, 22], [-50, 54.3, 22], [275, 54.3, 22]]

    hole_all = {'hole_op10_pick': [[-20, -32.45, 6.64], [-20, 1.45, 6.647], [40, -12.584, 6.647],
                                   [80, -25.505, 6.647], [120, -21.5, 6.647], [200.007, -32.469, 6.647],
                                   [240, -33.5, 6.647]],
                'Lower_pad_bolt_spacer': [[-20, -69.7, 13], [130, -69.7, 13], [-20, 69.3, 13],
                                          [130, 69.3, 13]],
                'Lower_pad_pin': [[-60, -24.7, 10.5], [-30, 24.3, 10.5], [190, 24.3, 10.5], [190, -24.7, 10.5]],
                'Lower_pad_Innerguide': [[-50, -54.7, 22], [290, -54.7, 22], [-50, 54.3, 22], [275, 54.3, 22]]
                }
    color = {'hole_op10_pick': [0, 0, 0, 0],
             'Lower_pad_bolt_spacer': [0, 0, 0, 0],
             'Lower_pad_pin': [255, 0, 0, 0],
             'Lower_pad_Innerguide': [255, 255, 0, 0]}
    oSelection = catapp.ActiveDocument.Selection
    oSelection.Search('name=GeneratedItem')
    oSelection.VisProperties.SetRealColor(255, 0, 255, 0)
    padx = [-75, -75, 315, 315, -75]
    pady = [-79.7, 79.7, 79.7, -79.7, -79.7]
    for item in range(0, 4):
        factory2d = drawingview.Factory2D
        line = factory2d.CreateLine(padx[item], pady[item], padx[item + 1], pady[item + 1])
        line.name = "Line_test"
    for item in hole_all:
        print('現在孔：%s' % hole_all[item])
        for item_2 in hole_all[item]:
            factory2d = drawingview.Factory2D
            circle = factory2d.CreateClosedCircle(item_2[0], item_2[1], item_2[2] / 2)  # [x,y,半徑]
            circle.Name = '%s' % item
        object = catapp.ActiveDocument.Selection
        object.Search('Name=%s,all' % item)
        vis = catapp.ActiveDocument.Selection.VisProperties  # 抓取物件
        vis.SetRealColor(color[item][0], color[item][1], color[item][2], color[item][3])  # 顏色控制
    drawingdocument.ExportData("" + gvar.drafting_output_path + "\\" + partname + ".svg", "svg")  # 使用svg存檔
    for item in hole_all:
        print('現在孔：%s' % hole_all[item])
        for item_2 in hole_all[item]:
            factory2d = drawingview.Factory2D
            circle = factory2d.CreateClosedCircle(item_2[0], item_2[1], item_2[2] / 2)  # [x,y,半徑]
            circle.Name = '%s' % item
        object = catapp.ActiveDocument.Selection
        object.Search('Name=%s,all' % item)
        object.Delete()
    object.Search('Name=Line_test,all')
    object.Delete()
    oSelection.Search('name=GeneratedItem')
    oSelection.VisProperties.SetRealColor(0, 0, 0, 0)


def lower_die_label(machining_shape_point_number, mainviewx, Length_max, mainviewy, Width_max, machining_shape_point_X,
                    machining_shape_point_Y):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # -----------------------------------------------------------------------------------
    if machining_shape_point_number > 0:  # 形狀孔
        for j in range(1, 1 + machining_shape_point_number):
            MyText64 = MyView.Texts.Add("L1", mainviewx - (Length_max / 2) + machining_shape_point_X[j] - 15,
                                        mainviewy - (Width_max / 2) + machining_shape_point_Y[j] + 5)
            MyText64.SetFontSize(0, 0, 12)


def lower_die_set_label(machining_shape_point_number, mainviewx, Length_max, mainviewy, Width_max,
                        machining_shape_point_X, machining_shape_point_Y, out_Guide_posts_X, out_Guide_posts_Y,
                        out_Guide_posts_bolt_hole_X, out_Guide_posts_bolt_hole_Y, out_Guide_posts_Avoid_Error_X):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # -----------------------------------------------------------------------------------
    if machining_shape_point_number > 0:  # 形狀孔
        for j in range(1, 1 + machining_shape_point_number):
            MyText64 = MyView.Texts.Add("L1", mainviewx - (Length_max / 2) + machining_shape_point_X[j] - 15,
                                        mainviewy - (Width_max / 2) + machining_shape_point_Y[j] + 5)
            MyText64.SetFontSize(0, 0, 12)
    MyText68 = MyView.Texts.Add("F", mainviewx - (Length_max / 2) + out_Guide_posts_X,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y)
    MyText68.SetFontSize(0, 0, 12)
    MyText68 = MyView.Texts.Add("F", mainviewx - (Length_max / 2) + out_Guide_posts_X,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y)
    MyText68.SetFontSize(0, 0, 12)
    MyText68 = MyView.Texts.Add("F", mainviewx + (Length_max / 2) - out_Guide_posts_X,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y)
    MyText68.SetFontSize(0, 0, 12)
    MyText68 = MyView.Texts.Add("F", mainviewx + (Length_max / 2) - out_Guide_posts_Avoid_Error_X,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y)
    MyText68.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X - (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y - (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X - (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y + (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X + (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y - (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X + (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y + (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)

    MyText68 = MyView.Texts.Add("A",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X - (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y - (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText68.SetFontSize(0, 0, 12)
    MyText68 = MyView.Texts.Add("A",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X - (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y + (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText68.SetFontSize(0, 0, 12)
    MyText68 = MyView.Texts.Add("A",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X + (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y - (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText68.SetFontSize(0, 0, 12)
    MyText68 = MyView.Texts.Add("A",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X + (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y + (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText68.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A",
                                mainviewx + (Length_max / 2) - out_Guide_posts_X - (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y - (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A",
                                mainviewx + (Length_max / 2) - out_Guide_posts_X - (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y + (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A",
                                mainviewx + (Length_max / 2) - out_Guide_posts_X + (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y - (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A",
                                mainviewx + (Length_max / 2) - out_Guide_posts_X + (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y + (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)

    MyText65 = MyView.Texts.Add("A", mainviewx + (Length_max / 2) - out_Guide_posts_Avoid_Error_X - (
            out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y - (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A", mainviewx + (Length_max / 2) - out_Guide_posts_Avoid_Error_X - (
            out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y + (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A", mainviewx + (Length_max / 2) - out_Guide_posts_Avoid_Error_X + (
            out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y - (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("A", mainviewx + (Length_max / 2) - out_Guide_posts_Avoid_Error_X + (
            out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y + (
                                        out_Guide_posts_bolt_hole_Y / 2) + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("HP",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X - (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("HP",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X + (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("HP",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X - (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("HP",
                                mainviewx - (Length_max / 2) + out_Guide_posts_X + (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy + (Width_max / 2) - out_Guide_posts_Y + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("HP",
                                mainviewx + (Length_max / 2) - out_Guide_posts_X - (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("HP",
                                mainviewx + (Length_max / 2) - out_Guide_posts_X + (
                                        out_Guide_posts_bolt_hole_X / 2) + 5,
                                mainviewy - (Width_max / 2) + out_Guide_posts_Y + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("HP", mainviewx + (Length_max / 2) - out_Guide_posts_Avoid_Error_X - (
            out_Guide_posts_bolt_hole_X / 2) + 5, mainviewy + (Width_max / 2) - out_Guide_posts_Y + 15)
    MyText65.SetFontSize(0, 0, 12)
    MyText65 = MyView.Texts.Add("HP", mainviewx + (Length_max / 2) - out_Guide_posts_Avoid_Error_X + (
            out_Guide_posts_bolt_hole_X / 2) + 5, mainviewy + (Width_max / 2) - out_Guide_posts_Y + 15)
    MyText65.SetFontSize(0, 0, 12)


def upper_die_set_label():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # -----------------------------------------------------------------------------------


def pilot_punch_insert_label():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # -----------------------------------------------------------------------------------


def cutting_cavity_d_insert_label(machining_explanation, machining_shape_point_number, mainviewx, mainviewy, Length_max,
                                  machining_shape_point_X, Width_max, machining_shape_point_Y):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # '-----------------------------------------------------------------------------------
    if machining_explanation[1] != "":  # '形狀孔
        for j in range(1, 1 + machining_shape_point_number):
            MyText64 = MyView.Texts.Add("DL1", mainviewx - (Length_max / 2) + machining_shape_point_X[j] - 15,
                                        mainviewy - (Width_max / 2) + machining_shape_point_Y[j])
            MyText64.SetFontSize(0, 0, 8)
    if machining_explanation[2] != "":  # '螺栓孔
        for j in range(1, 1 + 1):
            MyText65 = MyView.Texts.Add("A", mainviewx - (Length_max / 2) + 15, mainviewy - (Width_max / 2) + 20)
            MyText65.SetFontSize(0, 0, 8)
            MyText65 = MyView.Texts.Add("A", mainviewx - (Length_max / 2) + 15, mainviewy + (Width_max / 2))
            MyText65.SetFontSize(0, 0, 8)
            MyText65 = MyView.Texts.Add("A", mainviewx + (Length_max / 2) - 5, mainviewy - (Width_max / 2) + 20)
            MyText65.SetFontSize(0, 0, 8)
            MyText65 = MyView.Texts.Add("A", mainviewx + (Length_max / 2) - 5, mainviewy + (Width_max / 2))
            MyText65.SetFontSize(0, 0, 8)


def cutting_cavity_u_insert_label(machining_explanation, machining_shape_point_number, mainviewx, mainviewy, Length_max,
                                  machining_shape_point_X, Width_max, machining_shape_point_Y):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # '-----------------------------------------------------------------------------------
    if machining_explanation[1] != "":  # '形狀孔
        for j in range(1, 1 + machining_shape_point_number):
            MyText64 = MyView.Texts.Add("DL1", mainviewx - (Length_max / 2) + machining_shape_point_X[j] - 15,
                                        mainviewy - (Width_max / 2) + machining_shape_point_Y[j])
            MyText64.SetFontSize(0, 0, 8)
    if machining_explanation[2] != "":  # '螺栓孔
        for j in range(1, 1 + 1):
            MyText65 = MyView.Texts.Add("A", mainviewx - (Length_max / 2) + 15, mainviewy - (Width_max / 2) + 20)
            MyText65.SetFontSize(0, 0, 8)
            MyText65 = MyView.Texts.Add("A", mainviewx - (Length_max / 2) + 15, mainviewy + (Width_max / 2))
            MyText65.SetFontSize(0, 0, 8)
            MyText65 = MyView.Texts.Add("A", mainviewx + (Length_max / 2) - 5, mainviewy - (Width_max / 2) + 20)
            MyText65.SetFontSize(0, 0, 8)
            MyText65 = MyView.Texts.Add("A", mainviewx + (Length_max / 2) - 5, mainviewy + (Width_max / 2))
            MyText65.SetFontSize(0, 0, 8)


def lower_pad_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                    machining_shape_point_X, Width_max, machining_shape_point_Y):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # '-----------------------------------------------------------------------------------
    if machining_shape_point_number > 0:  # 形狀孔
        for j in range(1, 1 + machining_shape_point_number):
            MyText64 = MyView.Texts.Add("L1", mainviewx - (Length_max / 2) + machining_shape_point_X[j] - 15,
                                        mainviewy - (Width_max / 2) + machining_shape_point_Y[j] + 5)
            MyText64.SetFontSize(0, 0, 12)


def Stripper_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                   machining_shape_point_X, Width_max, machining_shape_point_Y):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # '-----------------------------------------------------------------------------------
    if machining_shape_point_number > 0:  # 形狀孔
        for j in range(1, 1 + machining_shape_point_number):
            MyText64 = MyView.Texts.Add("L1", mainviewx - (Length_max / 2) + machining_shape_point_X[j] - 15,
                                        mainviewy - (Width_max / 2) + machining_shape_point_Y[j] + 5)
            MyText64.SetFontSize(0, 0, 12)


def Stop_plate_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                     machining_shape_point_X, Width_max, machining_shape_point_Y):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # '-----------------------------------------------------------------------------------
    if machining_shape_point_number > 0:  # 形狀孔
        for j in range(1, 1 + machining_shape_point_number):
            MyText64 = MyView.Texts.Add("L1", mainviewx - (Length_max / 2) + machining_shape_point_X[j] - 15,
                                        mainviewy - (Width_max / 2) + machining_shape_point_Y[j] + 5)
            MyText64.SetFontSize(0, 0, 12)


def Splint_label(machining_shape_point_number, mainviewx, mainviewy, Length_max,
                 machining_shape_point_X, Width_max, machining_shape_point_Y):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView
    # '-----------------------------------------------------------------------------------
    if machining_shape_point_number > 0:  # 形狀孔
        for j in range(1, 1 + machining_shape_point_number):
            MyText64 = MyView.Texts.Add("L1", mainviewx - (Length_max / 2) + machining_shape_point_X[j] - 15,
                                        mainviewy - (Width_max / 2) + machining_shape_point_Y[j] + 5)
            MyText64.SetFontSize(0, 0, 12)


def up_plate_label():
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    drawingDocument1 = catapp.ActiveDocument
    drawingSheets1 = drawingDocument1.Sheets
    drawingSheet1 = drawingSheets1.Item("Model")
    drawingViews1 = drawingSheet1.Views
    drawingView1 = drawingViews1.Item("Main View")
    drawingView1.Activate()
    oSheets = drawingDocument1.Sheets
    oSheet = drawingSheets1.Item("Sheet.1")
    MyView = oSheet.Views.ActiveView

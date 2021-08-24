import global_var as gvar
import win32com.client as win32
import defs
import time
import csv


def StripAnalyze():
    ratio_point_distance = [[0.0] * 20 for ii in range(999)]  # (number_cruve,point->point distance)
    line_length = [0.0] * 999  # line length
    Appearance_number = int()
    Data_element_line = [None] * 99  # 拆解後的element放置位置
    Data_element_number = [0] * 99  # 拆解後的element放置array位置
    connect_element_line = [[None] * 99 for ii in range(99)]  # 連接後的element放置位置
    Data_element_ratio_point = [[None] * 20 for ii in range(999)]  # ratio_point放置
    Data_close_line = [None] * 99  # close_line放置
    Appearance_circle = [None] * 99  # 外型工站
    Appearance_array_number = [0] * 99  # 外型工站 array 位置
    Central_key_axi = [None] * 99  # 中心鍵槽
    Central_axi = [None] * 99  # 中心軸
    Central_axi_number = [0] * 10  # 中心軸 array 位置
    Boots_part = [None] * 99  # 靴齒部
    Rivete_hole = [None] * 99  # 鉚接孔
    element_point = [None] * 30
    Base_Document = "Strip_Data"
    # ================================建立基本元素(產品中點,方向座標,測量參數)================================(Date_base_Dim)
    # ---------------------------------程式參數區
    Sketch_position = "Hybridbody"  # 副程式element放置位置   ("Body","Hybridbody")  (本體,依據)
    close_line_N = 0
    # ---------------------------------程式參數區
    # -------------------------------------------------------------------------------------------------設定Part模板建立
    # -----------------------------------------------------------------------------起手是宣告
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    try:
        document1 = catapp.ActiveDocument
        document1.Close()
    except:
        pass
    if dir(gvar.open_path + "\\Strip_Data.stp") != "":
        partDocument1 = documents1.Open(gvar.open_path + Base_Document + ".stp")
        time.sleep(0.5)
    # -----------------------------------------------------------------------------起手是宣告
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(Base_Document,
                                                                                             "PartBody",
                                                                                             "Geometrical Set.1")  # 環境設定
    part1 = ElementDocument.Part
    relations1 = part1.Relations  # 關聯指令起手宣告
    parameters1 = part1.Parameters  # 參數指令起手宣告
    hybridShapeFactory1 = part1.HybridShapeFactory  # 曲面指令起手宣告
    hybridShapes1 = ElementHybridBody.HybridShapes  # 依據指令起手宣告
    specsAndGeomWindow1 = catapp.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    viewer3D1.Reframe()
    part1.Update()
    # -----------------------------------------------------------------------------起手是宣告
    hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 2, 3)  # 建立方向性
    (ElementPoint5) = defs.BuildXYZpoint(0, 0, 0, "zero_original_point", ElementDocument,
                                         ElementHybridBody)  # (X座標,Y座標,Z座標) element_point(5) 為out
    (ElementPoint5) = defs.BuildXYZpoint(-500, 0, 0, "X_min_point", ElementDocument,
                                         ElementHybridBody)  # (X座標,Y座標,Z座標) element_point(5) 為out
    ElementReference11 = ElementPoint5
    (ElementPoint5) = defs.BuildXYZpoint(500, 0, 0, "X_max_point", ElementDocument,
                                         ElementHybridBody)  # (X座標,Y座標,Z座標) element_point(5) 為out
    ElementReference12 = ElementPoint5
    (ElementPoint5) = defs.BuildXYZpoint(0, -500, 0, "Y_min_point", ElementDocument,
                                         ElementHybridBody)  # (X座標,Y座標,Z座標) element_point(5) 為out
    ElementReference13 = ElementPoint5
    (ElementPoint5) = defs.BuildXYZpoint(0, 500, 0, "Y_max_point", ElementDocument,
                                         ElementHybridBody)  # (X座標,Y座標,Z座標) element_point(5) 為out
    ElementReference14 = ElementPoint5
    # -----------------------------------------------------------------------------建立測量參數
    length1 = parameters1.CreateDimension("", "LENGTH", 0)  # build parameter
    length2 = parameters1.CreateDimension("", "LENGTH", 0)  # build parameter
    length1.rename("MeasureDistance")
    length2.rename("Measureline")
    formula1 = relations1.CreateFormula("measure_formula", "", length1, "length( )")
    formula2 = relations1.CreateFormula("measure_line_formula", "", length2, "length( ) ")
    # -----------------------------------------------------------------------------建立測量參數
    # ================================建立基本元素(產品中點,方向座標,測量參數)================================(Date_base_Dim)
    # ============================================建立點資料===========================================(Date_built_point)
    Data_element_ratio_point = [[None] * 20 for ii in range(999)]
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set("Strip_Data",
                                                                                             "PartBody",
                                                                                             "Geometrical Set.1")  # 環境設定
    part1 = ElementDocument.Part
    relations1 = part1.Relations  # 關聯指令起手宣告
    parameters1 = part1.Parameters  # 參數指令起手宣告
    hybridShapeFactory1 = part1.HybridShapeFactory  # 曲面指令起手宣告
    hybridShapes1 = ElementHybridBody.HybridShapes  # 依據指令起手宣告
    # -----------------------------------------------------------------------------起手是宣告
    # -----------------------------------------------------------------------------搜尋零散元素
    selection1 = ElementProduct.Selection
    selection1.Search("Name=NONE*,all")
    S_count = selection1.Count
    # -----------------------------------------------------------------------------搜尋零散元素
    for now_element in range(1, S_count + 1):  # 針對每個元素開始進行測量,建點
        hybridShape1 = selection1.Item(now_element).Value
        Data_element_line[now_element] = selection1.Item(now_element).Value
        Data_element_number[now_element] = now_element
        hybridShape1.Name = hybridShape1.Name + "_" + str(now_element)
        # -----------------------------------build the ratio point
        Ratio = [0, 0.5, 1]
        for ratio_value in Ratio:
            (ElementPoint5) = defs.BuildPointChose("Cruve_ratio", hybridShape1, ElementDocument, ElementHybridBody,
                                                   Sketch_position,
                                                   ElementHybridBody)  # (建點形式("Center_Cruve","Cruve_ratio"),依據1(弧線)) element_Reference(11)依據2 element_Reference(12)依據3   element_point(5) 為out
            ElementPoint5.Ratio.Value = ratio_value
            ElementPoint5.Name = hybridShape1.Name + "_point_" + str(ratio_value)
            ElementPoint5.Orientation = True
            element_point[int(ratio_value * 4)] = ElementPoint5
            Data_element_ratio_point[now_element][1 + int(ratio_value * 2)] = ElementPoint5
            try:
                part1.Update()
            except:
                ElementPoint5.point = ElementReference11
                part1.Update()
        # -----------------------------------build the ratio point
        # -----------------------------------measure point and point distance
        for now_measure_element in Ratio:
            part1.Update()
            formula1.Modify("distance(`" + ElementHybridBody.Name + "\\" + element_point[
                0].Name + "`, `" + ElementHybridBody.Name + "\\" + element_point[
                                int(now_measure_element * 4)].Name + "`) ")
            S_point_distance_parameter = ("`" + ElementHybridBody.Name + "\\" + element_point[0].Name + "`")
            E_point_distance_parameter = (
                    "`" + ElementHybridBody.Name + "\\" + element_point[int(now_measure_element * 4)].Name + "`")
            line_distance_parameter = ("`" + ElementHybridBody.Name + "\\" + hybridShape1.Name + "`")
            formula1.Modify(
                "length(" + line_distance_parameter + "," + S_point_distance_parameter + "," + E_point_distance_parameter + ") ")
            formula2.Modify("length(" + line_distance_parameter + ") ")
            part1.UpdateObject(formula1)  # 單步更新 formula
            part1.UpdateObject(formula2)  # 單步更新 formula
            ratio_point_distance[now_element][int(now_measure_element * 2)] = length1.Value
            line_length[now_element] = length2.Value
        # --------------------確定點方向是否正確
        if abs(ratio_point_distance[now_element][2]) < 0.1:
            element_point[0].Orientation = False
            element_point[2].Orientation = False
            element_point[4].Orientation = False
            part1.UpdateObject(formula1)  # 單步更新 formula
            part1.UpdateObject(formula2)  # 單步更新 formula
        # --------------------確定點方向是否正確
        # --------------------確定是否為封閉線段  (需先判斷完方向)
        if (abs(length2.Value) - abs(length1.Value)) > 0.01 and abs(length1.Value) != 0:
            # element_point(0).point = element_Reference(20)
            # element_point(2).point = element_Reference(20)
            # element_point(4).point = element_Reference(20)
            pass
        elif abs(length1.Value) == 0:
            close_line_N = close_line_N + 1
            Data_close_line[close_line_N] = hybridShape1
        # --------------------確定是否為封閉線段  (需先判斷完方向)
        part1.UpdateObject(formula1)  # 單步更新 formula
        part1.UpdateObject(formula2)  # 單步更新 formula
        part1.Update()
        ratio_point_distance[now_element][2] = length1.Value
        # -----------------------------------measure point and point distance
    selection1.Clear()
    part1.Update()
    # ============================================建立點資料===========================================(Date_built_point)
    # =====================================點數據蒐集,判斷外型、中心孔====================================(Date_point_analysis)
    # ---------------------------------程式參數區
    Sketch_position = "Hybridbody"  # 副程式element放置位置   ("Body","Hybridbody")  (本體,依據)
    mid_point_distance = [0.0] * 999
    zero_point_distance = [0.0] * 999
    one_point_distance = [0.0] * 999
    close_line_N = 0
    count_element = 0
    # ---------------------------------程式參數區
    # -----------------------------------------------------------------------------起手是宣告
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(Base_Document,
                                                                                             "PartBody",
                                                                                             "Geometrical Set.1")  # 環境設定
    part1 = ElementDocument.Part
    relations1 = part1.Relations  # 關聯指令起手宣告
    parameters1 = part1.Parameters  # 參數指令起手宣告
    hybridShapeFactory1 = part1.HybridShapeFactory  # 曲面指令起手宣告
    hybridShapes1 = ElementHybridBody.HybridShapes  # 依據指令起手宣告
    # -----------------------------------------------------------------------------起手是宣告
    # -----------------------------------------------------------------------------沿用已建立測量參數
    length1.rename("MeasureDistance")
    length2.rename("Measureline")
    # -----------------------------------------------------------------------------沿用已建立測量參數
    array_N = [None] * (S_count + 1)
    for now_element in range(1, S_count + 1):  # 針對每個元素開始進行測量,建點
        time.sleep(0.1)
        S_point_distance_parameter = "`" + ElementHybridBody.Name + "\\zero_original_point`"
        E_point_distance_parameter_1 = "`" + ElementHybridBody.Name + "\\" + Data_element_ratio_point[now_element][
            2].Name + "`"
        E_point_distance_parameter_2 = "`" + ElementHybridBody.Name + "\\" + Data_element_ratio_point[now_element][
            1].Name + "`"
        formula1.Modify("distance(" + S_point_distance_parameter + "," + E_point_distance_parameter_1 + ") ")
        part1.UpdateObject(formula1)  # 單步更新 formula
        mid_point_distance[now_element] = length1.Value
        formula1.Modify(
            "distance(" + S_point_distance_parameter + "," + E_point_distance_parameter_2 + ") ")
        part1.UpdateObject(formula1)  # 單步更新 formula
        zero_point_distance[now_element] = length1.Value
        array_N[now_element] = now_element
    # -------------------------------------------------------------------------------------------------排列測量出來的數據
    # ------------------------------------------------------------    由大排到小
    for i in range(1, S_count):  # 變數= start~倒數第二個
        Min = i  # 定義較小數之指標
        for j in range((i + 1), S_count + 1):  # 變數= i+1~end
            if mid_point_distance[Min] < mid_point_distance[j]:
                Min = j
        store_number = mid_point_distance[i]
        mid_point_distance[i] = mid_point_distance[Min]
        mid_point_distance[Min] = store_number
        store_number_N = array_N[i]
        array_N[i] = array_N[Min]
        array_N[Min] = store_number_N
    # ------------------------------------------------------------    由大排到小
    # -------------------------------------------------------------------------------------------------排列測量出來的數據
    # -------------------------------------------------------------------------------------------------judge element if appearance  and central hole
    Central_axi[1] = Data_element_line[array_N[S_count]]  # central element 距離最短
    Appearance_number = 0  # 計數參數初始  有幾個外型元素
    count_element = 1  # 計數參數初始  計執行到第幾個元素
    # ------------------------------------------------------------    外型元素有幾個符合資格
    while mid_point_distance[1] <= mid_point_distance[count_element]:  # 抓出所有最大值
        if mid_point_distance[count_element] == zero_point_distance[
            array_N[count_element]]:  # 確認最大值是否為外型(原型)  origial point->ratio0.5=origial point->ratio0
            Appearance_circle[count_element] = Data_element_line[array_N[count_element]]
            Appearance_array_number[count_element] = array_N[count_element]
        count_element = count_element + 1
        Appearance_number = Appearance_number + 1  # Confirm the number of appearance
    # ------------------------------------------------------------   外型元素有幾個符合資格
    # =====================================點數據蒐集,判斷外型、中心孔====================================(Date_point_analysis)
    if abs(ratio_point_distance[array_N[S_count]][2]) > 0.1:
        # =====================================鍵槽元素判斷====================================(search_contral_key_element)
        # -----------------------------------------------------------------------------起手是宣告
        (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(Base_Document,
                                                                                                 "PartBody",
                                                                                                 "Geometrical Set.1")  # 環境設定
        part1 = ElementDocument.Part
        relations1 = part1.Relations  # 關聯指令起手宣告
        parameters1 = part1.Parameters  # 參數指令起手宣告
        hybridShapeFactory1 = part1.HybridShapeFactory  # 曲面指令起手宣告
        hybridShapes1 = ElementHybridBody.HybridShapes  # 依據指令起手宣告
        selection1 = ElementProduct.Selection
        # -----------------------------------------------------------------------------起手是宣告
        #   參數array_N(S_count)為ratio0.5->original_point距離排序(大到小)過後的編碼   例子:array_N(1)=5    距離最長的為元素5
        #   參數S_count  總共有多少個元素
        # ------------------------------------------------------------------judge the element of key
        # -------------------------------------------------------------------initialization parameter
        count_element = 1  # 計數參數初始  計執行到第幾個元素
        count_close_element = 1  # count the element of cortral key
        close_elemant_number = array_N[S_count]  # contral_element_number
        aims_position = array_N[S_count]  # contral_element_number of position
        now_element = array_N[S_count]
        ratio_0_1_parameter = [0] * 3
        ratio_0_1_parameter[1] = 1
        ratio_0_1_parameter[2] = 3
        ratio_array_count = 0
        Central_key_part_line = [None] * 99  # 中心鍵槽
        # -------------------------------------------------------------------initialization parameter
        S_point_distance_parameter = "`" + ElementHybridBody.Name + "\\" + \
                                     Data_element_ratio_point[close_elemant_number][
                                         ratio_0_1_parameter[1]].Name + "`"  # start measure elemet
        # -----------------------------------------------------------------------------------------------------------------------------DO迴圈確認線段連接的終止條件
        for Termination_condition_close_element_number_end_pint in range(1, S_count + 1):
            # --------------------------------初始參數(重新紀錄)
            count_element = 1
            ratio_array_count = 0
            # --------------------------------初始參數(重新紀錄)
            # ---------------------------------------------------------------------DO迴圈確認點距離為0
            for distance_is_zero_termination in range(1, (S_count * 2) + 1):
                # -------------------------------------------------確認矩陣值ratio_array_count不會超過2
                ratio_array_count = ratio_array_count + 1
                if ratio_array_count == 3:
                    count_element = count_element + 1
                    ratio_array_count = 1
                # -------------------------------------------------確認矩陣值ratio_array_count不會超過2
                # -------------------------------------------------make a measure
                E_point_distance_parameter_1 = "`" + ElementHybridBody.Name + "\\" + Data_element_ratio_point[
                    count_element][ratio_0_1_parameter[ratio_array_count]].Name + "`"
                formula1.Modify(
                    "distance(" + S_point_distance_parameter + "," + E_point_distance_parameter_1 + ") ")
                part1.UpdateObject(formula1)  # 單步更新 formula
                # -------------------------------------------------make a measure+
                if abs(length1.Value) < 0.1:  # 抓出所有最大值
                    break
            # ---------------------------------------------------------------------DO迴圈確認點距離為0
            now_element = count_element  # 連接下一個元素
            # -------------------------------紀錄連接元素(尚未到達終止條件)
            if now_element != aims_position:
                selection1.Add(Data_element_line[now_element])  # 抓取目前的元素(CATIA呈現)
                Central_key_part_line[count_close_element] = Data_element_line[now_element]  # 存放元素到該參數
                count_close_element = count_close_element + 1
                # -------------------------------更換下一點進行測量連接
                if ratio_array_count == 1:
                    ratio_array_count = 2
                elif ratio_array_count == 2:
                    ratio_array_count = 1
                # -------------------------------更換下一點進行測量連接
                S_point_distance_parameter = "`" + ElementHybridBody.Name + "\\" + \
                                             Data_element_ratio_point[now_element][
                                                 ratio_0_1_parameter[ratio_array_count]].Name + "`"  # 下一個起點
                (S_count, array_N, Data_element_line, Data_element_number) = array_delete_now_F(array_N,
                                                                                                now_element,
                                                                                                S_count,
                                                                                                Data_element_ratio_point,
                                                                                                Data_element_line,
                                                                                                Data_element_number)  # 針對此模組資料庫進行刪除
                (aims_position) = test_search(Data_element_number, close_elemant_number, aims_position,
                                              S_count)  # 搜尋模組建立(陣列,搜尋值,位置,陣列空間大小)
            # -------------------------------紀錄連接元素(尚未到達終止條件)
            if now_element == aims_position:  # 抓出所有最大值
                break
        count_close_element = count_close_element - 1  # 計數執行會多一
        # ----------------------------------------------------------------DO迴圈確認線段連接的終止條件

        element_Reference1 = Central_key_part_line[1]  # join element of firts
        delete_E = [None] * 99
        # -------------------------------進行結合元素結合
        for join_N in range(2, count_close_element + 1):  # 第一個與第二個元素結合->最後一個    (兩個結合完才能換下一個)
            (element_Reference1) = defs.JoinElement(element_Reference1, Central_key_part_line[join_N],
                                                    ElementDocument,
                                                    ElementBody, ElementHybridBody,
                                                    Sketch_position)  # 元素組合(元素1,元素2)out=element_Reference(1)
            (element_Reference5) = defs.break_relationship(element_Reference1, "line", ElementDocument, ElementBody,
                                                           ElementHybridBody,
                                                           Sketch_position)  # (需打斷關係之元素,element_type="point" or "line")   element_Reference(5)為out
            defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
            delete_E[join_N] = element_Reference5  # 紀錄需要刪除的元素
            element_Reference1 = element_Reference5
        # -------------------------------進行結合元素結合
        # -------------------------------刪除非完連接線段全
        for join_D in range(2, count_close_element):
            defs.delete_object(delete_E[join_D], ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
        # -------------------------------刪除非完連接線段全
        element_Reference1.Name = "open_curve"
        Central_axi[1].Name = "circle_line"
        array_N_position = 1
        while array_N[S_count] != Data_element_number[array_N_position]:
            array_N_position = array_N_position + 1
        (S_count, array_N, Data_element_line, Data_element_number) = array_delete_now_F(array_N, array_N_position,
                                                                                        S_count,
                                                                                        Data_element_ratio_point,
                                                                                        Data_element_line,
                                                                                        Data_element_number)  # 針對此模組資料庫進行刪除
        # --------------------------------------------------------------------------------------judge the element of key
        # =====================================鍵槽元素判斷====================================(search_contral_key_element)
    else:
        Central_axi[1].Name = "circle_line"
        (S_count, array_N, Data_element_line, Data_element_number) = array_delete_now_F(array_N, array_N[S_count],
                                                                                        S_count,
                                                                                        Data_element_ratio_point,
                                                                                        Data_element_line,
                                                                                        Data_element_number)  # 針對此模組資料庫進行刪除
    # =====================================靴齒部線段判斷====================================(search_Appearance_connect_element)
    Boots_part_line = [[None] * 99 for ii in range(99)]  # 靴齒部線段
    aims_position = [0] * 999
    # -----------------------------------------------------------------------------起手是宣告
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(Base_Document,
                                                                                             "PartBody",
                                                                                             "Geometrical Set.1")  # 環境設定
    part1 = ElementDocument.Part
    relations1 = part1.Relations  # 關聯指令起手宣告
    parameters1 = part1.Parameters  # 參數指令起手宣告
    hybridShapeFactory1 = part1.HybridShapeFactory  # 曲面指令起手宣告
    hybridShapes1 = ElementHybridBody.HybridShapes  # 依據指令起手宣告
    selection1 = ElementProduct.Selection
    # -----------------------------------------------------------------------------起手是宣告
    # ---------------------------------------------judge element if appearance
    for Appearance_count in range(1, Appearance_number + 1):
        (aims_position[Appearance_count]) = test_search(Data_element_number,
                                                        Appearance_array_number[Appearance_count],
                                                        aims_position[Appearance_count],
                                                        S_count)  # 搜尋模組建立(陣列,搜尋值,位置,陣列空間大小)
    # -------------------------------------------------------------------initialization parameter
    count_element = 1  # 計數參數初始  計執行到第幾個元素
    count_close_element = 1  # count the element of cortral key
    loop_stop = 0
    now_element = aims_position[1]
    S_point_distance_parameter = "`" + ElementHybridBody.Name + "\\" + Data_element_ratio_point[now_element][
        ratio_0_1_parameter[1]].Name + "`"  # start measure elemet
    # -------------------------------------------------------------------initialization parameter
    for boots_number in range(1, Appearance_number + 1):
        time.sleep(0.1)
        count_close_element = 1
        # -----------------------------------------------------------------------------------------------------------------------------DO迴圈確認線段連接的終止條件
        for aa in range(1, S_count + 1):  # Termination condition close_element_number end pint
            # --------------------------------初始參數(重新紀錄)
            count_element = 1
            ratio_array_count = 0
            # --------------------------------初始參數(重新紀錄)
            # ---------------------------------------------------------------------DO迴圈確認點距離為0
            for aaa in range(1, (S_count * 2) + 1):
                # -------------------------------------------------確認矩陣值ratio_array_count不會超過2
                ratio_array_count = ratio_array_count + 1
                if ratio_array_count == 3:
                    count_element = count_element + 1
                    ratio_array_count = 1
                # -------------------------------------------------確認矩陣值ratio_array_count不會超過2
                E_point_distance_parameter_1 = "`" + ElementHybridBody.Name + "\\" + \
                                               Data_element_ratio_point[count_element][
                                                   ratio_0_1_parameter[ratio_array_count]].Name + "`"
                formula1.Modify(
                    "distance(" + S_point_distance_parameter + "," + E_point_distance_parameter_1 + ") ")
                part1.UpdateObject(formula1)  # 單步更新 formula
                if length1.Value < 0.1:  # 抓出所有最大值
                    break
            # ---------------------------------------------------------------------DO迴圈確認點距離為0
            # -------------------------------------------------紀錄元素and位置
            Boots_part_line[boots_number][count_close_element] = Data_element_line[count_element]  # 紀錄元素
            # -------------------------------------------------紀錄元素and位置
            # -------------------------------------------------檢測是否到達外型元素
            for Appearance_count in range(1, Appearance_number + 1):
                if count_element == aims_position[Appearance_count]:
                    loop_stop = 1
            # -------------------------------------------------檢測是否到達外型元素
            # -------------------------------------------------紀錄元素與數量
            if loop_stop != 1:
                selection1.Add(Boots_part_line[boots_number][count_close_element])
                count_close_element = count_close_element + 1
            # -------------------------------------------------紀錄元素與數量
            # -------------------------------------------------更換下一點進行測量連接
            if ratio_array_count == 1:
                ratio_array_count = 2
            elif ratio_array_count == 2:
                ratio_array_count = 1
            # -------------------------------------------------更換下一點進行測量連接
            now_element = count_element  # 轉換連接依據
            S_point_distance_parameter = "`" + ElementHybridBody.Name + "\\" + \
                                         Data_element_ratio_point[now_element][
                                             ratio_0_1_parameter[ratio_array_count]].Name + "`"
            if loop_stop != 1:
                (S_count, array_N, Data_element_line, Data_element_number) = array_delete_now_F(array_N,
                                                                                                now_element,
                                                                                                S_count,
                                                                                                Data_element_ratio_point,
                                                                                                Data_element_line,
                                                                                                Data_element_number)  # 針對此模組資料庫進行刪除
                for Appearance_count in range(1, Appearance_number + 1):
                    (aims_position[Appearance_count]) = test_search(Data_element_number,
                                                                    Appearance_array_number[Appearance_count],
                                                                    aims_position[Appearance_count],
                                                                    S_count)  # 搜尋模組建立(陣列,搜尋值,位置,陣列空間大小)
            if loop_stop == 1:  # 抓出所有最大值
                break
        # -----------------------------------------------------------------------------------------------------------------------------DO迴圈確認線段連接的終止條件
        count_close_element = count_close_element - 1
        loop_stop = 0
        element_Reference1 = Boots_part_line[boots_number][1]
        delete_E = [None] * 99
        # -------------------------------進行結合元素結合
        for join_N in range(2, count_close_element + 1):  # 第一個與第二個元素結合->最後一個    (兩個結合完才能換下一個)
            (element_Reference1) = defs.JoinElement(element_Reference1, Boots_part_line[boots_number][join_N],
                                                    ElementDocument, ElementBody, ElementHybridBody,
                                                    Sketch_position)  # 元素組合(元素1,元素2)out=element_Reference(1)
            (element_Reference5) = defs.break_relationship(element_Reference1, "line", ElementDocument, ElementBody,
                                                           ElementHybridBody,
                                                           Sketch_position)  # (需打斷關係之元素,element_type="point" or "line")   element_Reference(5)為out
            defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
            delete_E[join_N] = element_Reference5  # 記錄需刪除的元素
            element_Reference1 = element_Reference5
        element_Reference1.Name = "open_curve"
        # -------------------------------刪除非完全連接線段
        for join_D in range(2, count_close_element):
            defs.delete_object(delete_E[join_D], ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
        # -------------------------------刪除非完全連接線段
    # -------------------------------------------------------------------------------------------------judge element if appearance
    for Appearance_count in range(1, Appearance_number + 1):
        Appearance_circle[Appearance_count].Name = "contour_circle_line"
        selection1.Add(Appearance_circle[Appearance_count])
        (S_count, array_N, Data_element_line, Data_element_number) = array_delete_now_F(array_N, 1, S_count,
                                                                                        Data_element_ratio_point,
                                                                                        Data_element_line,
                                                                                        Data_element_number)  # 針對此模組資料庫進行刪除(刪除位置)
    for Appearance_count in range(1, 5):
        Data_element_line[Appearance_count].Name = "cut_line"
    selection1.Search("Name=NONE*,all")
    selection1.Delete()
    # =====================================靴齒部線段判斷====================================(search_Appearance_connect_element)
    ElementDocument.ExportData(gvar.open_path + "Strip_Data-2.stp", "stp")
    ElementDocument.Close()
    return mid_point_distance[1]


def array_delete_now_F(array_N, delete_array_N, S_count, Data_element_ratio_point, Data_element_line,
                       Data_element_number):  # 針對此模組資料庫進行刪除
    array_N_position = 1
    for count_araay_D in range(delete_array_N, S_count):
        Data_element_ratio_point[count_araay_D][1] = Data_element_ratio_point[count_araay_D + 1][1]
        Data_element_ratio_point[count_araay_D][2] = Data_element_ratio_point[count_araay_D + 1][2]
        Data_element_ratio_point[count_araay_D][3] = Data_element_ratio_point[count_araay_D + 1][3]
    while array_N[array_N_position] != Data_element_number[delete_array_N]:
        array_N_position = array_N_position + 1
    (array_N) = array_delete_element(array_N, S_count, array_N_position)  # 刪除陣列元素(陣列參數,空間大小,刪除的編號)
    (Data_element_line) = array_delete_element(Data_element_line, S_count,
                                               delete_array_N)  # 刪除陣列元素(陣列參數,空間大小,刪除的編號)
    (Data_element_number) = array_delete_element(Data_element_number, S_count,
                                                 delete_array_N)  # 刪除陣列元素(陣列參數,空間大小,刪除的編號)
    S_count = S_count - 1
    return S_count, array_N, Data_element_line, Data_element_number


def array_delete_element(array_parameter, room, delete_number):
    for count_araay_D in range(delete_number, room):
        array_parameter[count_araay_D] = array_parameter[count_araay_D + 1]
    return array_parameter


def test_search(array_Name, search_number, position_N, array_limit):  # 搜尋模組建立(陣列,搜尋值,位置,陣列空間大小)
    array_N_position = 1
    while array_Name[array_N_position] != search_number and array_limit >= array_N_position:
        array_N_position = array_N_position + 1
    position_N = array_N_position
    return position_N


def StripBuild(R_value):
    temporary_point = [[None] * 5 for ii in range(99)]
    contour_number = int()
    contour_circle_R = [None] * 30
    open_curve_distance = [None] * 99  # open_curve環狀分類
    element_point = [None] * 30
    total_op_number = int(gvar.strip_parameter_list[2])
    pitch = int(gvar.strip_parameter_list[4])
    Sketch_position = "Hybridbody"
    OP_Empty = [0] * 10
    OP_Empty[1] = total_op_number
    OP_Empty[2] = 2
    # 第N站為空站↑(0:取消)
    testnum = int()
    Y_limit = -20  # 板材的邊界(距離產品多少)  出現地方"建立邊界點(圖在第一象限)"
    X_limit = -20  # 板材的邊界(距離產品多少)  出現地方"建立邊界點(圖在第一象限)"
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    try:
        document1 = catapp.ActiveDocument
        document1.Close()
    except:
        pass
    partDocument1 = documents1.Open(gvar.open_path + "Strip_Data-2.stp")
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set("Strip_Data-2", "PartBody",
                                                                                             "Geometrical Set.1")  # 環境設定
    ElementHybridBody.Name = "die"
    (zero_original_point) = defs.BuildXYZpoint(0, 0, 0, "zero_original_point", ElementDocument,
                                               ElementHybridBody)  # (X座標,Y座標,Z座標) element_point(5) 為out---------------------------------------------------------------刪掉
    selection1 = ElementDocument.Selection
    selection1.Clear()
    selection1.Search("Name=contour_circle_line,all")
    contour_circle_line_number = selection1.Count
    part1 = ElementDocument.Part
    originElements1 = part1.OriginElements
    hybridShape2 = originElements1.PlaneXY
    hybridShapes1 = ElementHybridBody.HybridShapes  # 依據指令起手宣告
    (element_sketch1) = defs.BuildSketch("Calculation_sketch", hybridShape2, ElementDocument, Sketch_position,
                                         ElementBody,
                                         ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    # -------------第一次修改名稱(尚未依照位置排序)
    for element_number in range(1, 1 + contour_circle_line_number):
        hybridShape1 = ElementHybridBody.HybridShapes.Item("contour_circle_line")  # 宣告平面
        hybridShape1.Name = "contour_circle_line_" + str(element_number)
        element_point[2] = hybridShape1
        # 將外框半徑確認
        (R_value) = defs.SketchBuildCallout(element_sketch1, "Radius", "Callout", R_value, ElementDocument,
                                            element_point[2], element_point[3])  #
        contour_circle_R[element_number] = R_value
    defs.delete_object(element_sketch1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    # -------------第一次修改名稱(尚未依照位置排序)
    # -------------建立contour_circle(打斷關聯後)
    (element_point[5]) = defs.BuildPointChose("Center_Cruve", hybridShape1, ElementDocument, ElementHybridBody,
                                              Sketch_position,
                                              ElementBody)  # (建點形式("Center_Cruve"),依據1(弧線)) element_Reference(11)依據2 element_Reference(12)依據3   element_point(5) 為out
    element_point[5].Name = "Circle_Center_Point"
    part1.Update()
    (element_sketch1) = defs.BuildSketch("Contour_Circle_sketch", hybridShape2, ElementDocument, Sketch_position,
                                         ElementBody,
                                         ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    (element_Reference30) = defs.SketchHidePoint(element_sketch1, element_point[5], 0, 0, "True", ElementDocument,
                                                 element_sketch1)  # (草圖陳述句,依據點之陳述句,+X,+Y,是否為實體("True","False")) output element_Reference(30)->point
    (element_Reference11) = defs.SketchCircle(element_sketch1, element_Reference30, R_value,
                                              ElementDocument)  # (草圖陳述句,依據點之陳述句,半徑)
    element_Reference11 = element_sketch1
    element_Reference12 = hybridShape2
    (element_line5) = defs.ProjectionLine(element_Reference11, element_Reference12, ElementDocument, ElementHybridBody,
                                          ElementBody, Sketch_position,
                                          "False")  # 投影線段  element_Reference(11)=投影之元素 #element_Reference(12)=plane  element_line(5) 為out
    element_line5.Name = "Contour_circle_line"
    defs.delete_object(element_sketch1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    defs.delete_object(element_point[5], ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    # -------------建立contour_circle(打斷關聯後)
    # -------------建立邊界點(圖在第一象限)
    (element_sketch1) = defs.BuildSketch("Calculation_sketch", hybridShape2, ElementDocument, Sketch_position,
                                         ElementBody,
                                         ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    (element_point[21], element_point[22], element_point[23], element_point[24]) = defs.ElementExtremumFourPoint(
        element_line5, ElementDocument, ElementBody, ElementHybridBody)  # 建立極值點   (極值方向1,極值方向2,極值方向3,需要幾個方向,在哪個元素上的極值)
    # 擺放第幾個元素的極值點
    temporary_point[element_number][1] = element_point[21]  # element_number_X_min
    temporary_point[element_number][2] = element_point[22]  # element_number_X_max
    temporary_point[element_number][3] = element_point[23]  # element_number_Y_min
    temporary_point[element_number][4] = element_point[24]  # element_number_Y_max
    part1.Update()
    (Contour_X_value) = defs.SketchBuildCallout(element_sketch1, "Horizontal", "Callout", 0, ElementDocument,
                                                element_point[21], element_point[
                                                    22])  # 建立標註(草圖陳述句,標註方向("Horizontal","Vertical","free","Radius"),式標OR拘束("Binding","Callout"),改變OR讀取之數值,element_point(now_point)and+1)
    (Contour_Y_value) = defs.SketchBuildCallout(element_sketch1, "Vertical", "Callout", 0, ElementDocument,
                                                element_point[23], element_point[
                                                    24])  # 建立標註(草圖陳述句,標註方向("Horizontal","Vertical","free","Radius"),式標OR拘束("Binding","Callout"),改變OR讀取之數值,element_point(now_point)and+1)
    defs.delete_object(element_sketch1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    # ------------------原點建立
    element_Reference10 = element_point[23]  # element_number_Y_min
    (element_point[5]) = defs.BuildPoint(element_Reference10, ElementDocument, ElementBody, ElementHybridBody,
                                         Sketch_position)  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
    element_point[5].X.Value = -Contour_X_value / 2 + X_limit
    element_point[5].Y.Value = Y_limit
    part1.Update()
    (element_Reference5) = defs.break_relationship(element_point[5], "point", ElementDocument, ElementBody,
                                                   ElementHybridBody,
                                                   Sketch_position)  # (需打斷關係之元素,element_type="point" or "line")   element_Reference(5)為out
    element_Reference5.Name = "origin_point"
    defs.hide(element_Reference5, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
    defs.delete_object(element_point[5], ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    # ------------------原點建立
    # -------------建立邊界點(圖在第一象限)
    # -------------open_curve建立中心點
    element_point[1] = element_Reference5
    # ----確認open_curve的數量
    selection1.Clear()
    selection1.Search("Name=open_curve,all")
    open_curve_number = selection1.Count
    # ----確認open_curve的數量
    for element_number in range(1, 1 + open_curve_number):
        E_open_curve = ElementHybridBody.HybridShapes.Item("open_curve")  # 宣告平面
        E_open_curve.Name = "open_curve_" + str(element_number)  # 建立中心點
        (element_point[5]) = Search_Graphics_center(E_open_curve, hybridShape2, element_point[1], ElementDocument,
                                                    ElementBody, ElementHybridBody, Sketch_position,
                                                    ElementProduct)  # out_put=element_point(5)
        element_point[5].Name = "open_curve_center_point_" + str(element_number)
        element_point[10 + element_number] = element_point[5]
    an = [""] * 99  # 存角度
    (circle_line_type) = circle_measure()  # 確定是否全圓 out circle_line_type=true or false
    selection1.Clear()
    selection1.Search("Name=point*,all")
    selection1.Delete()
    line_type = 1
    if circle_line_type == False:
        element_number = 1
        E_open_curve = ElementHybridBody.HybridShapes.Item("open_curve_" + str(element_number))
        E_open_curve.Name = "open_curve_" + str(line_type) + "_" + str(element_number)
        E_open_curve = ElementHybridBody.HybridShapes.Item("circle_line")  # 宣告平面
        (element_point[5]) = defs.BuildPointChose("Cruve_ratio", E_open_curve, ElementDocument, ElementHybridBody,
                                                  Sketch_position,
                                                  ElementBody)  # (建點形式("Center_Cruve","Cruve_ratio"),依據1(弧線,弧線))element_point(5) 為out
        element_point[5].Ratio.Value = 0
        element_point[5].Name = "open_curve_" + str(line_type) + "_" + str(element_number) + "_A"
        defs.hide(element_point[5], ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
        element_point[18] = element_point[5]
        (element_point[5]) = defs.BuildPointChose("Cruve_ratio", E_open_curve, ElementDocument, ElementHybridBody,
                                                  Sketch_position,
                                                  ElementBody)  # (建點形式("Center_Cruve","Cruve_ratio"),依據1(弧線,弧線))element_point(5) 為out
        element_point[5].Ratio.Value = 1
        element_point[5].Name = "open_curve_" + str(line_type) + "_" + str(element_number) + "_B"
        defs.hide(element_point[5], ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
        element_point[16] = element_point[5]
        part1.Update()
        (angel_out_value) = defs.AngleMeasure(zero_original_point, element_point[16], ElementDocument,
                                              ElementHybridBody)  # 中心點水平逆時針測量(中心點,測量點,輸出角度)
        an[element_number] = angel_out_value
        (element_sketch1) = defs.BuildSketch("Calculation_sketch_1", hybridShape2, ElementDocument, Sketch_position,
                                             ElementBody,
                                             ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
        defs.hide(element_sketch1, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
        (element_Reference5, element_Reference6, element_point[17], element_point[19]) = tryangle(an[element_number],
                                                                                                  element_sketch1,
                                                                                                  circle_line_type,
                                                                                                  line_type,
                                                                                                  ElementDocument,
                                                                                                  ElementHybridBody)
        defs.SketchBuildCallout(element_sketch1, "free", "Binding", 0, ElementDocument, element_point[16],
                                element_point[
                                    17])  # 建立標註(草圖陳述句,標註方向("Horizontal","Vertical","free","Radius"),式標OR拘束("Binding","Callout"),改變OR讀取之數值,element_point(now_point)and+1)
        defs.SketchBuildCallout(element_sketch1, "free", "Binding", 0, ElementDocument, element_point[18],
                                element_point[
                                    19])  # 建立標註(草圖陳述句,標註方向("Horizontal","Vertical","free","Radius"),式標OR拘束("Binding","Callout"),改變OR讀取之數值,element_point(now_point)and+1)
        selection1.Add(element_Reference5)
        selection1.Add(element_Reference6)
        selection1.Delete()
        E_open_curve = ElementHybridBody.HybridShapes.Item("open_curve_" + str(line_type) + "_" + str(element_number))
        (element_Reference1) = defs.JoinElement(E_open_curve, element_sketch1, ElementDocument, ElementBody,
                                                ElementHybridBody,
                                                Sketch_position)  # (元素1,元素2)   out=element_Reference(1)
        (element_Reference5) = defs.break_relationship(element_Reference1, "join", ElementDocument, ElementBody,
                                                       ElementHybridBody,
                                                       Sketch_position)  # (需打斷關係之元素,element_type="point" or "line")   element_Reference(5)為out
        defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
        element_Reference5.Name = "finish_" + E_open_curve.Name
        selection1.Add(element_sketch1)
        selection1.Delete()
        parameters1 = part1.Parameters
        relations1 = part1.Relations
        length6 = parameters1.CreateDimension("", "LENGTH", 0)
        formula1 = relations1.Createformula("formula.1", "", length6,
                                            "length(`die\\" + element_Reference5.Name + "`  )")
        length6.rename("keyway_circle_line_" + str(element_number))
        gvar.SumKeywayCircleLine = gvar.SumKeywayCircleLine + length6.Value
        line_type = line_type + 1
    final_number = open_curve_number - line_type + 1
    for element_number in range(1, 1 + final_number):
        if line_type != 1:
            E_open_curve = ElementHybridBody.HybridShapes.Item("open_curve_" + str(element_number + line_type - 1))
        else:
            E_open_curve = ElementHybridBody.HybridShapes.Item("open_curve_" + str(element_number))
        E_open_curve.Name = "open_curve_" + str(line_type) + "_" + str(element_number)
        selection1.Add(E_open_curve)
        (element_point[5]) = defs.BuildPointChose("Cruve_ratio", E_open_curve, ElementDocument, ElementHybridBody,
                                                  Sketch_position,
                                                  ElementBody)  # (建點形式("Center_Cruve","Cruve_ratio"),依據1(弧線,弧線))  element_point(5) 為out
        element_point[5].Ratio.Value = 0
        element_point[5].Name = "open_curve_" + str(line_type) + "_" + str(element_number) + "_A"
        defs.hide(element_point[5], ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
        element_point[16] = element_point[5]
        (element_point[5]) = defs.BuildPointChose("Cruve_ratio", E_open_curve, ElementDocument, ElementHybridBody,
                                                  Sketch_position,
                                                  ElementBody)  # (建點形式("Center_Cruve","Cruve_ratio"),依據1(弧線,弧線))  element_point(5) 為out
        element_point[5].Ratio.Value = 1
        element_point[5].Name = "open_curve_" + str(line_type) + "_" + str(element_number) + "_B"
        defs.hide(element_point[5], ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
        element_point[18] = element_point[5]
        (angel_out_value_1) = defs.AngleMeasure(zero_original_point, element_point[16], ElementDocument,
                                                ElementHybridBody)  # 中心點水平逆時針測量(中心點,測量點,輸出角度)
        (angel_out_value_2) = defs.AngleMeasure(zero_original_point, element_point[18], ElementDocument,
                                                ElementHybridBody)  # 中心點水平逆時針測量(中心點,測量點,輸出角度)
        if angel_out_value_1 < angel_out_value_2:
            element_point[19] = element_point[16]
            element_point[16] = element_point[18]
            element_point[18] = element_point[19]
        an[element_number] = angel_out_value_1
        (element_sketch1) = defs.BuildSketch("Calculation_sketch_" + str(element_number), hybridShape2, ElementDocument,
                                             Sketch_position, ElementBody,
                                             ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
        defs.hide(element_sketch1, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
        (element_Reference5, element_Reference6, element_point[17], element_point[19]) = tryangle(-(an[element_number]),
                                                                                                  element_sketch1,
                                                                                                  circle_line_type,
                                                                                                  line_type,
                                                                                                  ElementDocument,
                                                                                                  ElementHybridBody)
        defs.SketchBuildCallout(element_sketch1, "free", "Binding", 0, ElementDocument, element_point[16],
                                element_point[
                                    17])  # 建立標註(草圖陳述句,標註方向("Horizontal","Vertical","free","Radius"),式標OR拘束("Binding","Callout"),改變OR讀取之數值,element_point(now_point)and+1)
        defs.SketchBuildCallout(element_sketch1, "free", "Binding", 0, ElementDocument, element_point[18],
                                element_point[
                                    19])  # 建立標註(草圖陳述句,標註方向("Horizontal","Vertical","free","Radius"),式標OR拘束("Binding","Callout"),改變OR讀取之數值,element_point(now_point)and+1)
        selection1.Clear()
        selection1.Add(element_Reference5)
        selection1.Add(element_Reference6)
        selection1.Delete()
        E_open_curve = ElementHybridBody.HybridShapes.Item("open_curve_" + str(line_type) + "_" + str(element_number))
        (element_Reference1) = defs.JoinElement(E_open_curve, element_sketch1, ElementDocument, ElementBody,
                                                ElementHybridBody,
                                                Sketch_position)  # (元素1,元素2)   out=element_Reference(1)
        (element_Reference5) = defs.break_relationship(element_Reference1, "join", ElementDocument, ElementBody,
                                                       ElementHybridBody,
                                                       Sketch_position)  # (需打斷關係之元素,element_type="point" or "line")   element_Reference(5)為out
        defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
        element_Reference5.Name = "finish_" + E_open_curve.Name
        selection1.Add(element_sketch1)
        selection1.Delete()
        parameters1 = part1.Parameters
        relations1 = part1.Relations
        length6 = parameters1.CreateDimension("", "LENGTH", 0)
        formula1 = relations1.Createformula("formula.1", "", length6,
                                            "length(`die\\" + element_Reference5.Name + "`  )")
        length6.rename("boots_part_circle_line_" + str(element_number))
        gvar.SumBootsPartCircleLine = gvar.SumBootsPartCircleLine + length6.Value
    # -------------open_curve環狀分類
    selection1.Clear()
    selection1.Search("Name=cut_line,all")
    cut_line_number = selection1.Count
    (element_sketch1) = defs.BuildSketch("Calculation_sketch_1", hybridShape2, ElementDocument, Sketch_position,
                                         ElementBody,
                                         ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    for element_number in range(1, 1 + cut_line_number):
        hybridShape1 = ElementHybridBody.HybridShapes.Item("cut_line")  # 宣告平面
        hybridShape1.Name = "cut_line_" + str(element_number)
        parameters1 = part1.Parameters
        relations1 = part1.Relations
        length6 = parameters1.CreateDimension("", "LENGTH", 0)
        formula1 = relations1.Createformula("formula.1", "", length6, "length(`die\\" + hybridShape1.Name + "`  )")
        length6.rename("rivet_hole_circle_line_" + str(element_number))
        gvar.SumRivetHoleCircleLine = gvar.SumRivetHoleCircleLine + length6.Value
    hybridShape1 = ElementHybridBody.HybridShapes.Item("circle_line")  # 宣告平面
    element_point[2] = hybridShape1
    (R_value) = defs.SketchBuildCallout(element_sketch1, "Radius", "Callout", R_value, ElementDocument,
                                        element_point[2], element_point[
                                            3])  # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"or"Radius"
    contour_circle_R[element_number] = R_value
    defs.delete_object(element_sketch1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    # -------------建立contour_circle(打斷關聯後)
    (element_point[5]) = defs.BuildPointChose("Center_Cruve", hybridShape1, ElementDocument, ElementHybridBody,
                                              Sketch_position,
                                              ElementBody)  # (建點形式("Center_Cruve"),依據1(弧線)) element_Reference(11)依據2 element_Reference(12)依據3   element_point(5) 為out
    element_point[5].Name = "Circle_Center_Point"
    part1.Update()
    (element_sketch1) = defs.BuildSketch("Contour_Circle_sketch", hybridShape2, ElementDocument, Sketch_position,
                                         ElementBody,
                                         ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    (element_Reference30) = defs.SketchHidePoint(element_sketch1, element_point[5], 0, 0, "True", ElementDocument,
                                                 element_sketch1)  # (草圖陳述句,依據點之陳述句,+X,+Y,是否為實體("True","False")) output element_Reference(30)->point
    (element_Reference11) = defs.SketchCircle(element_sketch1, element_Reference30, R_value,
                                              ElementDocument)  # (草圖陳述句,依據點之陳述句,半徑)
    element_Reference11 = element_sketch1
    element_Reference12 = hybridShape2
    (element_line5) = defs.ProjectionLine(element_Reference11, element_Reference12, ElementDocument, ElementHybridBody,
                                          ElementBody, Sketch_position,
                                          "False")  # 投影線段  element_Reference(11)=投影之元素 #element_Reference(12)=plane  element_line(5) 為out
    element_line5.Name = "cut_circle_line"
    parameters1 = part1.Parameters
    relations1 = part1.Relations
    length6 = parameters1.CreateDimension("", "LENGTH", 0)
    formula1 = relations1.Createformula("formula.1", "", length6, "length(`die\\" + hybridShape1.Name + "`  )")
    length6.rename("central_pocket_circle_line_1")
    gvar.SumCentralPocketCircleLine = gvar.SumCentralPocketCircleLine + length6.Value
    defs.delete_object(element_sketch1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    defs.delete_object(element_point[5], ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    # -------------建立contour_circle(打斷關聯後)
    element_Reference10 = element_point[1]
    (element_point[5]) = defs.BuildPoint(element_Reference10, ElementDocument, ElementBody, ElementHybridBody,
                                         Sketch_position)  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
    strip_width = (Contour_Y_value + 2)  # 料條寬 (非浮生銷參數)   原料帶寬+10
    strip_length = total_op_number * int(gvar.strip_parameter_list[4]) + 10  # 料條長 (非浮生銷參數)
    if total_op_number % 2 != 0:
        element_point[5].X.Value = Contour_X_value / 2 - X_limit + (int(total_op_number / 2) * pitch)
    else:
        element_point[5].X.Value = (Contour_X_value / 2) - X_limit + (int(total_op_number / 2) * pitch) - (0.5 * pitch)
    element_point[5].Y.Value = Contour_Y_value / 2 - Y_limit
    element_point[5].Name = "plate_centor_point"
    (element_sketch1) = defs.BuildSketch("strip_sketch", hybridShape2, ElementDocument, Sketch_position, ElementBody,
                                         ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    (element_point[6], element_line1, element_line2, element_line3, element_line4) = defs.SketchRectangle(
        element_sketch1, strip_length, strip_width, ElementDocument,
        ElementHybridBody)  # (草圖陳述句,長,寬)  element_point(1)=>output中心點之陳述句  [需經環境副程式跑過]
    element_point[1] = element_Reference10
    defs.SketchBuildCallout(element_sketch1, "free", "Binding", 0, ElementDocument, element_point[5], element_point[
        6])  # 建立標註(草圖陳述句,標註方向("Horizontal","Vertical","free","Radius"),式標OR拘束("Binding","Callout"),改變OR讀取之數值,element_point(now_point)and+1)
    reference1 = part1.CreateReferenceFromName("")
    shapeFactory1 = part1.ShapeFactory
    part1.InWorkObject = ElementBody
    pad1 = shapeFactory1.AddNewPadFromRef(reference1, 0.5)
    reference2 = part1.CreateReferenceFromObject(element_sketch1)
    pad1.SetProfileElement(reference2)
    limit1 = pad1.FirstLimit
    length1 = limit1.dimension
    length1.Value = 0.35
    Sketch_position = "Body"
    part1.Update()  # 成品線段
    for i in range(1, 1 + contour_circle_line_number):
        E_open_curve = ElementHybridBody.HybridShapes.Item("contour_circle_line_" + str(i))
        (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * total_op_number + 5, "X", ElementDocument,
                                                     ElementHybridBody, ElementBody,
                                                     Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
        element_Reference1.Name = "contour_circle_line_number_" + str(i)
        defs.hide(element_Reference1, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
        time.sleep(0.5)
        selection1.Clear()
        selection1.Add(element_Reference1)
        selection1.Copy()
        selection1.Search("Name=" + ElementHybridBody.Name + ",all")
        selection1.Paste()
        defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
        parameters1 = part1.Parameters
        relations1 = part1.Relations
        length6 = parameters1.CreateDimension("", "LENGTH", 0)
        formula1 = relations1.Createformula("formula.1", "", length6, "length(`die\\" + E_open_curve.Name + "`  )")
        length6.rename("contour_circle_circle_line_" + str(element_number))
        gvar.SumContourCircleCircleLine = gvar.SumContourCircleCircleLine + length6.Value
    for i in range(1, 1 + final_number):
        E_open_curve = ElementHybridBody.HybridShapes.Item("open_curve_" + str(line_type) + "_" + str(i))
        (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * total_op_number + 5, "X", ElementDocument,
                                                     ElementHybridBody, ElementBody,
                                                     Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
        element_Reference1.Name = "open_curve_number_" + str(i)
        defs.hide(element_Reference1, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
        selection1.Clear()
        selection1.Add(element_Reference1)
        selection1.Copy()
        time.sleep(1)
        E_open_curve = ElementHybridBody.HybridShapes.Item("open_curve_" + str(line_type) + "_" + str(i))
        selection1.Clear()
        selection1.Search("Name=" + ElementHybridBody.Name + ",all")
        selection1.Paste()
        defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    for i in range(1, 1 + contour_circle_line_number + final_number - 1):
        Sketch_position = "Hybridbody"
        if i % 2 != 0 and i != 1:
            E_open_curve = ElementHybridBody.HybridShapes.Item("open_curve_number_" + str(testnum))
            testnum = testnum + 1
        elif i != 1:
            E_open_curve = ElementHybridBody.HybridShapes.Item("contour_circle_line_number_" + str(testnum))
        if i == 1:
            E_open_curve = ElementHybridBody.HybridShapes.Item("contour_circle_line_number_1")
            element_Reference1 = ElementHybridBody.HybridShapes.Item("open_curve_number_1")
            testnum = 2
        (element_Reference1) = defs.JoinElement(E_open_curve, element_Reference1, ElementDocument, ElementBody,
                                                ElementHybridBody,
                                                Sketch_position)  # (元素1,元素2)   out=element_Reference(1)
        element_Reference1.Name = "garbage" + str(i)
        if i == contour_circle_line_number + final_number - 1:
            # =============打斷關聯投影線條===============(J_program.DatumLine)
            part1 = ElementDocument.Part
            hybridShapeFactory1 = part1.HybridShapeFactory
            hybridBodies1 = part1.HybridBodies
            hybridBody1 = hybridBodies1.Item("die")
            hybridShapes1 = hybridBody1.HybridShapes
            hybridShapeAssemble1 = hybridShapes1.Item(element_Reference1.Name)
            reference1 = part1.CreateReferenceFromObject(hybridShapeAssemble1)
            originElements1 = part1.OriginElements
            hybridShapePlaneExplicit1 = originElements1.PlaneXY
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneExplicit1)
            hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference1, reference2)
            hybridShapeProject1.SolutionType = 0
            hybridShapeProject1.Normal = True
            hybridShapeProject1.SmoothingType = 0
            hybridBody1.AppendHybridShape(hybridShapeProject1)
            part1.InWorkObject = hybridShapeProject1
            part1.Update()
            reference3 = part1.CreateReferenceFromObject(hybridShapeProject1)
            hybridShapeCurveExplicit1 = hybridShapeFactory1.AddNewCurveDatum(reference3)
            hybridBody1.AppendHybridShape(hybridShapeCurveExplicit1)
            hybridShapeCurveExplicit1.Name = "blank"
            element_Reference1 = hybridShapeCurveExplicit1
            part1.InWorkObject = hybridShapeCurveExplicit1
            part1.Update()
            hybridShapeFactory1.DeleteObjectForDatum(reference3)
            # =============打斷關聯投影線條===============(J_program.DatumLine)
            selection1.Clear()
            selection1.Search("Name=garbage*,all")
            selection1.Delete()
            selection1.Clear()
            selection1.Search("Name=contour_circle_line_number_*,all")
            selection1.Delete()
            selection1.Clear()
            selection1.Search("Name=open_curve_number_*,all")
            selection1.Delete()
    part1.InWorkObject = ElementBody
    pad2 = shapeFactory1.AddNewPadFromRef(reference1, 0.5)
    reference3 = part1.CreateReferenceFromObject(element_Reference1)
    pad2.SetProfileElement(reference3)
    limit2 = pad1.FirstLimit
    length2 = limit2.dimension
    length2.Value = 0.35
    Sketch_position = "Body"
    part1.Update()
    specsAndGeomWindow1 = catapp.ActiveWindow
    viewer3D1 = specsAndGeomWindow1.ActiveViewer
    viewer3D1.Reframe()
    part1.Update()
    length6 = parameters1.Item("die\\strip_sketch\\Length.139\\Length")
    if length6.Value <= 40:
        guide_pin_diameter = 1.5
    if length6.Value > 40 and length6.Value <= 60:
        guide_pin_diameter = 2
    if length6.Value > 60 and length6.Value <= 80:
        guide_pin_diameter = 2.5
    if length6.Value > 80 and length6.Value <= 100:
        guide_pin_diameter = 3
    if length6.Value > 100 and length6.Value <= 120:
        guide_pin_diameter = 3.5
    if length6.Value > 120 and length6.Value <= 140:
        guide_pin_diameter = 4
    if length6.Value > 140 and length6.Value <= 160:
        guide_pin_diameter = 4.5
    if length6.Value > 160 and length6.Value <= 180:
        guide_pin_diameter = 5
    if length6.Value > 180 and length6.Value < 200:
        guide_pin_diameter = 5.5
    else:
        guide_pin_diameter = 6
    (OP_Empty) = Momo_station_change(OP_Empty, total_op_number, guide_pin_diameter)
    # --------------------------------改變孔位
    for hole_number_N in range(1, 1 + total_op_number + 1):  # 挖除定位孔\
        (element_point[5]) = defs.BuildPoint(element_Reference10, ElementDocument, ElementBody, ElementHybridBody,
                                             Sketch_position)  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
        element_point[5].X.Value = Contour_X_value / 2 - X_limit - pitch / 2 + (hole_number_N - 1) * pitch
        element_point[5].Y.Value = -Y_limit + 1 + guide_pin_diameter / 2  # +1做調整
        element_Reference11 = element_point[5]
        element_Reference12 = hybridShape2
        (hole_A) = defs.HoleSimpleD(4, 1, 1, ElementDocument, ElementBody, element_Reference11,
                                    element_Reference12)  # 直孔  (M,深度,out孔陳述式,方向) element_Reference(11)=point #element_Reference(12)=plane direction=>0=下 1=上
        if hole_number_N == 1:
            Sketch_position = "Hybridbody"
            (element_sketch1) = defs.BuildSketch("strip_sketch", hybridShape2, ElementDocument, Sketch_position,
                                                 ElementBody,
                                                 ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
            (element_Reference30) = defs.SketchHidePoint(element_sketch1, element_point[5], 0, 0, "True",
                                                         ElementDocument,
                                                         element_sketch1)  # (草圖陳述句,依據點之陳述句,+X,+Y,是否為實體("True","False")) output element_Reference(30)->point
            (element_Reference11) = defs.SketchCircle(element_sketch1, element_Reference30, 0.75,
                                                      ElementDocument)  # (草圖陳述句,依據點之陳述句,半徑)
            element_Reference11 = element_sketch1
            (element_line5) = defs.ProjectionLine(element_Reference11, element_Reference12, ElementDocument,
                                                  ElementHybridBody, ElementBody, Sketch_position,
                                                  'False')  # 投影線段  element_Reference(11)=投影之元素 #element_Reference(12)=plane  element_line(5) 為out
            element_line5.Name = "plate_line_1_op10_A_punch_1"
            defs.delete_object(element_sketch1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
            Sketch_position = "Body"
        (element_point[5]) = defs.BuildPoint(element_Reference10, ElementDocument, ElementBody, ElementHybridBody,
                                             Sketch_position)  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
        element_point[5].X.Value = Contour_X_value / 2 - X_limit - 20 + (hole_number_N - 1) * pitch
        element_point[5].Y.Value = Contour_Y_value - Y_limit - 1 - guide_pin_diameter / 2  # -1做調整
        element_Reference11 = element_point[5]
        element_Reference12 = hybridShape2
        (hole_B) = defs.HoleSimpleD(4, 1, 1, ElementDocument, ElementBody, element_Reference11,
                                    element_Reference12)  # 直孔  (M,深度,out孔陳述式,方向) element_Reference(11)=point #element_Reference(12)=plane direction=>0=下 1=上
        if hole_number_N == 1:
            Sketch_position = "Hybridbody"
            (element_sketch1) = defs.BuildSketch("strip_sketch", hybridShape2, ElementDocument, Sketch_position,
                                                 ElementBody,
                                                 ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
            (element_Reference30) = defs.SketchHidePoint(element_sketch1, element_point[5], 0, 0, "True",
                                                         ElementDocument,
                                                         element_sketch1)  # (草圖陳述句,依據點之陳述句,+X,+Y,是否為實體("True","False")) output element_Reference(30)->point
            (element_Reference11) = defs.SketchCircle(element_sketch1, element_Reference30, 0.75,
                                                      ElementDocument)  # (草圖陳述句,依據點之陳述句,半徑)
            element_Reference11 = element_sketch1
            (element_line5) = defs.ProjectionLine(element_Reference11, element_Reference12, ElementDocument,
                                                  ElementHybridBody, ElementBody, Sketch_position,
                                                  "False")  # 投影線段  element_Reference(11)=投影之元素 #element_Reference(12)=plane  element_line(5) 為out
            element_line5.Name = "plate_line_1_op10_A_punch_2"
            defs.delete_object(element_sketch1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
            Sketch_position = "Body"
    # 挖除鍵槽
    if line_type != 1:
        element_number = 1
        E_open_curve = ElementHybridBody.HybridShapes.Item("finish_open_curve_1_1")  #
        for OP_N in range(OP_Empty[4], 1 + total_op_number + 1):
            if OP_N == total_op_number + 1:
                (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1) + 5, "X", ElementDocument,
                                                             ElementHybridBody, ElementBody, Sketch_position)
            else:
                (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1), "X", ElementDocument,
                                                             ElementHybridBody, ElementBody,
                                                             Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
            # =============挖除線段===============
            part1 = ElementDocument.Part
            shapeFactory1 = part1.ShapeFactory
            reference1 = part1.CreateReferenceFromObject(element_Reference1)
            part1.InWorkObject = ElementBody
            pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
            element_Reference20 = pocket1
            # =============挖除線段===============
            if OP_N == OP_Empty[4]:
                if OP_N == total_op_number + 1:
                    (element_Reference1) = defs.TranslateElement(E_open_curve, pitch + pitch * OP_N + 5, "X",
                                                                 ElementDocument, ElementHybridBody,
                                                                 ElementBody, Sketch_position)
                else:
                    (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1), "X", ElementDocument,
                                                                 ElementHybridBody,
                                                                 ElementBody, Sketch_position)
                element_Reference1.Name = "plate_line_1_op" + str(OP_Empty[4]) + "0_cut_line_1"
                defs.hide(element_Reference1, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
                selection1.Clear()
                selection1.Add(element_Reference1)
                selection1.Copy()
                selection1.Search("Name=" + ElementHybridBody.Name + ",all")
                selection1.Paste()
                defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
            element_Reference20.FirstLimit.LimitMode = 2
        part1.Update()
    for element_number in range(1, 1 + final_number):  # 挖除靴齒部
        E_open_curve = ElementHybridBody.HybridShapes.Item(
            "finish_open_curve_" + str(line_type) + "_" + str(element_number))  #
        for OP_N in range(OP_Empty[5], 1 + total_op_number):
            (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1), "X", ElementDocument,
                                                         ElementHybridBody, ElementBody,
                                                         Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
            # =============挖除線段===============
            part1 = ElementDocument.Part
            shapeFactory1 = part1.ShapeFactory
            reference1 = part1.CreateReferenceFromObject(element_Reference1)
            part1.InWorkObject = ElementBody
            pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
            element_Reference20 = pocket1
            # =============挖除線段===============
            if OP_N == OP_Empty[5]:
                (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1), "X", ElementDocument,
                                                             ElementHybridBody, ElementBody,
                                                             Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
                element_Reference1.Name = "plate_line_1_op" + str(OP_Empty[5]) + "0_cut_line_" + str(element_number)
                defs.hide(element_Reference1, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
                selection1.Clear()
                selection1.Add(element_Reference1)
                selection1.Copy()
                selection1.Search("Name=" + ElementHybridBody.Name + ",all")
                selection1.Paste()
                defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
            element_Reference20.FirstLimit.LimitMode = 2
    part1.Update()
    for element_number in range(1, 1 + cut_line_number):  # 挖除鉚接孔
        E_open_curve = ElementHybridBody.HybridShapes.Item("cut_line_" + str(element_number))  # 宣告平面
        for OP_N in range(OP_Empty[6], 1 + total_op_number + 1):
            if OP_N == (total_op_number + 1):
                (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1) + 5, "X", ElementDocument,
                                                             ElementHybridBody, ElementBody, Sketch_position)
            else:
                (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1), "X", ElementDocument,
                                                             ElementHybridBody, ElementBody,
                                                             Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
            # =============挖除線段===============
            part1 = ElementDocument.Part
            shapeFactory1 = part1.ShapeFactory
            reference1 = part1.CreateReferenceFromObject(element_Reference1)
            part1.InWorkObject = ElementBody
            pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
            element_Reference20 = pocket1
            # =============挖除線段===============
            element_Reference20.FirstLimit.LimitMode = 2
            if OP_N == OP_Empty[6]:
                (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1), "X", ElementDocument,
                                                             ElementHybridBody, ElementBody,
                                                             Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
                element_Reference1.Name = "plate_line_1_op" + str(OP_Empty[6]) + "0_cut_line_" + str(element_number)
                defs.hide(element_Reference1, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
                selection1.Clear()
                selection1.Add(element_Reference1)
                selection1.Copy()
                selection1.Search("Name=" + ElementHybridBody.Name + ",all")
                selection1.Paste()
                defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
            part1.Update()
    for OP_N in range(OP_Empty[7], 1 + total_op_number + 1):  # 挖除中心孔
        E_open_curve = ElementHybridBody.HybridShapes.Item("cut_circle_line")  # 宣告平面
        if OP_N == (total_op_number + 1):
            (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1) + 5, "X", ElementDocument,
                                                         ElementHybridBody, ElementBody, Sketch_position)
        else:
            (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1), "X", ElementDocument,
                                                         ElementHybridBody, ElementBody,
                                                         Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
        # =============挖除線段===============
        part1 = ElementDocument.Part
        shapeFactory1 = part1.ShapeFactory
        reference1 = part1.CreateReferenceFromObject(element_Reference1)
        part1.InWorkObject = ElementBody
        pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
        element_Reference20 = pocket1
        # =============挖除線段===============
        element_Reference20.FirstLimit.LimitMode = 2
        if OP_N == OP_Empty[7]:
            (element_Reference1) = defs.TranslateElement(E_open_curve, pitch * (OP_N - 1), "X", ElementDocument,
                                                         ElementHybridBody, ElementBody,
                                                         Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
            element_Reference1.Name = "plate_line_1_op" + str(OP_Empty[7]) + "0_cut_line_1"
            defs.hide(element_Reference1, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
            selection1.Clear()
            selection1.Add(element_Reference1)
            selection1.Copy()
            selection1.Search("Name=" + ElementHybridBody.Name + ",all")
            selection1.Paste()
            defs.delete_object(element_Reference1, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
        part1.Update()
    # --------------------------------------------------------------------------------------------------------------------------------------------
    # 挖除外型線
    E_open_curve = ElementHybridBody.HybridShapes.Item("Contour_circle_line")  # 宣告平面
    (element_Reference1) = defs.TranslateElement(E_open_curve, (total_op_number - 1) * pitch, "X", ElementDocument,
                                                 ElementHybridBody, ElementBody,
                                                 Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
    # =============挖除線段===============
    part1 = ElementDocument.Part
    shapeFactory1 = part1.ShapeFactory
    reference1 = part1.CreateReferenceFromObject(element_Reference1)
    part1.InWorkObject = ElementBody
    pocket1 = shapeFactory1.AddNewPocketFromRef(reference1, 20)
    element_Reference20 = pocket1
    # =============挖除線段===============
    element_Reference20.FirstLimit.LimitMode = 2
    (element_Reference1) = defs.TranslateElement(E_open_curve, (total_op_number - 1) * pitch, "X", ElementDocument,
                                                 ElementHybridBody, ElementBody,
                                                 Sketch_position)  # (元素1,距離,方向)out=element_Reference(1)
    element_Reference1.Name = "plate_line_1_op70_cut_line_1"
    defs.hide(element_Reference1, ElementDocument)  # (hide_element) 惟須隱藏東西的陳述句
    selection1.Clear()
    selection1.Add(element_Reference1)
    selection1.Copy()
    selection1.Search("Name=" + ElementHybridBody.Name + ",all")
    selection1.Paste()
    defs.delete_object(element_Reference1, ElementProduct)
    part1.Update()
    ElementHybridBody.Name = "die"
    time.sleep(2)
    ElementDocument.SaveAs(gvar.open_path + "Strip_Data-2.CATPart")  # 存檔的檔案名稱
    ElementDocument.ExportData(gvar.file_path + "\\auto\\料帶備份\\Strip_Data-2.stp", "stp")  # 料帶備份
    ElementDocument.Close()
    return strip_length, strip_width


def tryangle(angle_number, element_sketch1, circle_line_type, line_type, ElementDocument, ElementHybridBody):
    part1 = ElementDocument.Part
    factory2D1 = element_sketch1.OpenEdition()
    geometricElements1 = element_sketch1.GeometricElements
    axis2D1 = geometricElements1.Item("AbsoluteAxis")
    line2D1 = axis2D1.getItem("HDirection")
    line2D1.ReportName = 1
    line2D2 = axis2D1.getItem("VDirection")
    line2D2.ReportName = 2
    point2D1 = factory2D1.CreatePoint(311.099756, 25.077236)
    point2D1.ReportName = 3
    point2D2 = factory2D1.CreatePoint(310.75246, 23.107621)
    point2D2.ReportName = 4
    line2D3 = factory2D1.CreateLine(311.099756, 25.077236, 310.75246, 23.107621)
    line2D3.ReportName = 5
    line2D3.StartPoint = point2D1
    line2D3.EndPoint = point2D2
    point2D3 = factory2D1.CreatePoint(308.782844, 23.454917)
    point2D3.ReportName = 6
    line2D4 = factory2D1.CreateLine(310.75246, 23.107621, 308.782844, 23.454917)
    line2D4.ReportName = 7
    line2D4.StartPoint = point2D2
    line2D4.EndPoint = point2D3
    point2D4 = factory2D1.CreatePoint(307.914603, 18.530878)
    point2D4.ReportName = 8
    line2D5 = factory2D1.CreateLine(308.782844, 23.454917, 307.914603, 18.530878)
    line2D5.ReportName = 9
    line2D5.StartPoint = point2D3
    line2D5.EndPoint = point2D4
    point2D5 = factory2D1.CreatePoint(325.015544, 15.515521)
    point2D5.ReportName = 10
    line2D6 = factory2D1.CreateLine(307.914603, 18.530878, 325.015544, 15.515521)
    line2D6.ReportName = 11
    line2D6.StartPoint = point2D4
    line2D6.EndPoint = point2D5
    point2D6 = factory2D1.CreatePoint(325.883784, 20.43956)
    point2D6.ReportName = 12
    line2D7 = factory2D1.CreateLine(325.015544, 15.515521, 325.883784, 20.43956)
    line2D7.ReportName = 13
    line2D7.StartPoint = point2D5
    line2D7.EndPoint = point2D6
    constraints1 = element_sketch1.Constraints
    reference2 = part1.CreateReferenceFromObject(point2D6)
    reference3 = part1.CreateReferenceFromObject(line2D4)
    constraint1 = constraints1.AddBiEltCst(2, reference2, reference3)
    constraint1.mode = 0
    point2D7 = factory2D1.CreatePoint(323.914169, 20.786856)
    point2D7.ReportName = 14
    line2D8 = factory2D1.CreateLine(325.883784, 20.43956, 323.914169, 20.786856)
    line2D8.ReportName = 15
    line2D8.StartPoint = point2D6
    line2D8.EndPoint = point2D7
    point2D8 = factory2D1.CreatePoint(324.261465, 22.756472)
    point2D8.ReportName = 16
    line2D9 = factory2D1.CreateLine(323.914169, 20.786856, 324.261465, 22.756472)
    line2D9.ReportName = 17
    line2D9.StartPoint = point2D7
    line2D9.EndPoint = point2D8
    point2D9 = factory2D1.CreatePoint(319.922297, 36.630088)
    point2D9.ReportName = 18
    point2D10 = factory2D1.CreatePoint(315.139627, 9.506219)
    point2D10.ReportName = 19
    line2D10 = factory2D1.CreateLine(319.922297, 36.630088, 315.139627, 9.506219)
    line2D10.ReportName = 20
    line2D10.Construction = True
    line2D10.StartPoint = point2D9
    line2D10.EndPoint = point2D10
    point2D11 = factory2D1.CreatePoint(309.157593, 33.844704)
    point2D11.ReportName = 21
    point2D12 = factory2D1.CreatePoint(343.239685, 33.844704)
    point2D12.ReportName = 22
    line2D11 = factory2D1.CreateLine(309.157593, 33.844704, 343.239685, 33.844704)
    line2D11.ReportName = 23
    line2D11.Construction = True
    line2D11.StartPoint = point2D11
    line2D11.EndPoint = point2D12
    reference4 = part1.CreateReferenceFromObject(line2D11)
    reference5 = part1.CreateReferenceFromObject(line2D1)
    constraint2 = constraints1.AddBiEltCst(10, reference4, reference5)
    constraint2.mode = 0
    reference6 = part1.CreateReferenceFromObject(line2D4)
    reference7 = part1.CreateReferenceFromObject(line2D3)
    constraint3 = constraints1.AddBiEltCst(11, reference6, reference7)
    constraint3.mode = 0
    reference8 = part1.CreateReferenceFromObject(line2D4)
    reference9 = part1.CreateReferenceFromObject(line2D5)
    constraint4 = constraints1.AddBiEltCst(11, reference8, reference9)
    constraint4.mode = 0
    reference10 = part1.CreateReferenceFromObject(line2D5)
    reference11 = part1.CreateReferenceFromObject(line2D6)
    constraint5 = constraints1.AddBiEltCst(11, reference10, reference11)
    constraint5.mode = 0
    reference12 = part1.CreateReferenceFromObject(line2D6)
    reference13 = part1.CreateReferenceFromObject(line2D7)
    constraint6 = constraints1.AddBiEltCst(11, reference12, reference13)
    constraint6.mode = 0
    reference14 = part1.CreateReferenceFromObject(line2D7)
    reference15 = part1.CreateReferenceFromObject(line2D8)
    constraint7 = constraints1.AddBiEltCst(11, reference14, reference15)
    constraint7.mode = 0
    reference16 = part1.CreateReferenceFromObject(line2D8)
    reference17 = part1.CreateReferenceFromObject(line2D9)
    constraint8 = constraints1.AddBiEltCst(11, reference16, reference17)
    constraint8.mode = 0
    reference18 = part1.CreateReferenceFromObject(line2D9)
    reference19 = part1.CreateReferenceFromObject(line2D3)
    reference20 = part1.CreateReferenceFromObject(line2D10)
    constraint9 = constraints1.AddTriEltCst(15, reference18, reference19, reference20)
    constraint9.mode = 0
    reference21 = part1.CreateReferenceFromObject(line2D5)
    reference22 = part1.CreateReferenceFromObject(line2D7)
    reference23 = part1.CreateReferenceFromObject(line2D10)
    constraint10 = constraints1.AddTriEltCst(15, reference21, reference22, reference23)
    constraint10.mode = 0
    reference24 = part1.CreateReferenceFromObject(line2D10)
    reference25 = part1.CreateReferenceFromObject(line2D7)
    constraint11 = constraints1.AddBiEltCst(1, reference24, reference25)
    constraint11.mode = 0
    length1 = constraint11.dimension
    length1.Value = 2.5
    reference26 = part1.CreateReferenceFromObject(line2D9)
    constraint12 = constraints1.AddMonoEltCst(5, reference26)
    constraint12.mode = 0
    length2 = constraint12.dimension
    length2.Value = 1
    reference27 = part1.CreateReferenceFromObject(line2D3)
    constraint13 = constraints1.AddMonoEltCst(5, reference27)
    constraint13.mode = 0
    length3 = constraint13.dimension
    length3.Value = 1
    reference28 = part1.CreateReferenceFromObject(line2D7)
    constraint14 = constraints1.AddMonoEltCst(5, reference28)
    constraint14.mode = 0
    length4 = constraint14.dimension
    length4.Value = 3
    reference31 = part1.CreateReferenceFromObject(point2D8)
    reference32 = part1.CreateReferenceFromObject(line2D10)
    constraint16 = constraints1.AddBiEltCst(1, reference31, reference32)
    constraint16.mode = 0
    length5 = constraint16.dimension
    length5.Value = 1
    reference29 = part1.CreateReferenceFromObject(line2D10)
    reference30 = part1.CreateReferenceFromObject(line2D11)
    constraint15 = constraints1.AddBiEltCst(6, reference29, reference30)
    constraint15.mode = 0
    constraint15.AngleSector = 0
    angle1 = constraint15.dimension
    angle1.Value = angle_number
    element_Reference5 = constraint16
    element_Reference6 = constraint15
    if circle_line_type == False and line_type == 1:
        element_point17 = point2D1
        element_point19 = point2D8
    else:
        element_point17 = point2D8
        element_point19 = point2D1
    element_sketch1.CloseEdition()
    part1.InWorkObject = ElementHybridBody
    part1.Update()
    part1.InWorkObject = element_sketch1
    return element_Reference5, element_Reference6, element_point17, element_point19


def Search_Graphics_center(G_element, plane_element, origin_point, ElementDocument, ElementBody, ElementHybridBody,
                           Sketch_position, ElementProduct):  # out_put=element_point(5)
    element_point = [None] * 30
    part1 = ElementDocument.Part
    X_X_distance = 0
    Y_Y_distance = 0
    X_min_distance = 0
    X_max_distance = 0
    Y_min_distance = 0
    Y_max_distance = 0
    temporary_number = int()
    (element_point[21], element_point[22], element_point[23], element_point[24]) = defs.ElementExtremumFourPoint(
        G_element, ElementDocument, ElementBody,
        ElementHybridBody)  # (建立元素) 建立4個極點 X_min element_point(21) X_max element_point(22) Y_min element_point(23) Y_max element_point(24)
    part1.Update()
    (element_sketch) = defs.BuildSketch("Calculation_sketch", plane_element, ElementDocument, Sketch_position,
                                        ElementBody,
                                        ElementHybridBody)  # (sketch名稱,產生sketch的平面)  element_sketch為產生出來的草圖
    # -------------------------------------------------找出X and Y的距離
    (X_X_distance) = defs.SketchBuildCallout(element_sketch, "Horizontal", "Callout", X_X_distance, ElementDocument,
                                             element_point[21], element_point[
                                                 22])  # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    (Y_Y_distance) = defs.SketchBuildCallout(element_sketch, "Vertical", "Callout", Y_Y_distance, ElementDocument,
                                             element_point[23], element_point[
                                                 24])  # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    # -------------------------------------------------找出X and Y的距離
    element_point[6] = origin_point
    # -------------------------------------------------X座標
    element_point[7] = element_point[21]
    (X_min_distance) = defs.SketchBuildCallout(element_sketch, "Horizontal", "Callout", X_min_distance, ElementDocument,
                                               element_point[6], element_point[
                                                   7])  # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    element_point[7] = element_point[22]
    (X_max_distance) = defs.SketchBuildCallout(element_sketch, "Horizontal", "Callout", X_max_distance, ElementDocument,
                                               element_point[6], element_point[
                                                   7])  # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    if X_min_distance > X_max_distance:
        temporary_number = X_min_distance
        X_min_distance = X_max_distance
        X_max_distance = temporary_number
    # -------------------------------------------------X座標
    # -------------------------------------------------Y座標
    element_point[7] = element_point[23]
    (Y_min_distance) = defs.SketchBuildCallout(element_sketch, "Vertical", "Callout", Y_min_distance, ElementDocument,
                                               element_point[6], element_point[
                                                   7])  # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    element_point[7] = element_point[24]
    (Y_max_distance) = defs.SketchBuildCallout(element_sketch, "Vertical", "Callout", Y_max_distance, ElementDocument,
                                               element_point[6], element_point[
                                                   7])  # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    if Y_min_distance > Y_max_distance:
        temporary_number = Y_min_distance
        Y_min_distance = Y_max_distance
        Y_max_distance = temporary_number
    # -------------------------------------------------Y座標
    element_Reference10 = origin_point
    (element_point[5]) = defs.BuildPoint(element_Reference10, ElementDocument, ElementBody, ElementHybridBody,
                                         Sketch_position)  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
    element_point[5].X.Value = X_min_distance + X_X_distance / 2
    element_point[5].Y.Value = Y_min_distance + Y_Y_distance / 2
    defs.delete_object(element_sketch, ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    for point_N in range(21, 25):
        defs.delete_object(element_point[point_N], ElementProduct)  # delete_element為需刪除元素之陳述句[需經過環境程式]
    return element_point[5]


def circle_measure():
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    hybridShapeCircleExplicit1 = parameters1.Item("circle_line")
    reference1 = part1.CreateReferenceFromObject(hybridShapeCircleExplicit1)
    hybridBodies1 = part1.HybridBodies
    hybridBody1 = hybridBodies1.Item("die")
    hybridShapeFactory1 = part1.HybridShapeFactory
    hybridShapePointOnCurve1 = hybridShapeFactory1.AddNewPointOnCurveFromPercent(reference1, 0, False)
    hybridBody1.AppendHybridShape(hybridShapePointOnCurve1)
    part1.InWorkObject = hybridShapePointOnCurve1
    try:
        part1.Update()
        circle_line_type = False
    except:
        circle_line_type = True
    return circle_line_type


def Momo_station_change(OP_Empty, total_op_number, guide_pin_diameter):
    OP_station_Quantity = total_op_number  # 工站數量
    OP_station_Pitch = int(gvar.strip_parameter_list[4])  # 工站距離(mm)
    Empty_station = OP_Empty[2]  # 空站站號
    Strip_Center = (OP_station_Quantity * OP_station_Pitch) / 2  # 料帶中心
    Force_total = float()
    Force_22 = float()
    Force_23 = float()
    Force_24 = float()
    Force_25 = float()
    Force_26 = float()
    Force_32 = float()
    Force_33 = float()
    Force_34 = float()
    Force_35 = float()
    Force_42 = float()
    Force_43 = float()
    Force_44 = float()
    Force_45 = float()
    Force_46 = float()
    Force_53 = float()
    Force_54 = float()
    Force_55 = float()
    Force_56 = float()
    # --------------分配工站位置
    OP_station_1 = 0 - Strip_Center  # 第一站位置=0-料帶中心
    OP_station_2 = (((2 * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center  # 第二站位置=(((2*2)-1)*(工站距離/2))-料帶中心
    OP_station_3 = (((3 * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center  # 第三站位置=(((3*2)-1)*(工站距離/2))-料帶中心
    OP_station_4 = (((4 * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center  # 第四站位置=(((4*2)-1)*(工站距離/2))-料帶中心
    OP_station_5 = (((5 * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center  # 第五站位置=(((5*2)-1)*(工站距離/2))-料帶中心
    OP_station_6 = (((6 * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center  # 第六站位置=(((6*2)-1)*(工站距離/2))-料帶中心
    OP_station_7 = (((7 * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center  # 第七站位置=(((7*2)-1)*(工站距離/2))-料帶中心
    OP_station_8 = (((8 * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center  # 第八站位置=(((8*2)-1)*(工站距離/2))-料帶中心
    OP_station_9 = (((9 * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center  # 第九站位置=(((9*2)-1)*(工站距離/2))-料帶中心
    OP_station_10 = (((10 * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center  # 第十站位置=(((10*2)-1)*(工站距離/2))-料帶中心
    # --------------各沖頭周長
    perimeter = [guide_pin_diameter * 3.14, gvar.SumBootsPartCircleLine, gvar.SumKeywayCircleLine,
                 gvar.SumRivetHoleCircleLine, gvar.SumCentralPocketCircleLine, gvar.SumContourCircleCircleLine, 0, 0, 0,
                 0]
    # --------------計算總周長
    for X in range(0, len(perimeter)):
        Force_total = Force_total + perimeter[X]
    # --------------
    # 第一站引導沖力量=引導沖周長*(0-料帶中心)
    Force_11 = perimeter[0] * (0 - Strip_Center)
    # 設定perimeter_Max是最後一站=Perimeter(工站數量-1-空站數量)，因Array()從0開始
    perimeter_Max = perimeter[OP_station_Quantity - 1 - Empty_station]
    # 最後一站為下料，力量=下料周長*((((工站數量*2)-1)*(工站距離/2))-料帶中心)
    Force_Max = perimeter_Max * ((((OP_station_Quantity * 2) - 1) * (OP_station_Pitch / 2)) - Strip_Center)
    time.sleep(2)
    # ---------------------------------------------------第一迴圈第二站配全部(除空站)
    for Station in range(1, 1 + (OP_station_Quantity - 2 - Empty_station)):  # Array()從0開始，第二站到倒數第二站，扣除空站
        Force = perimeter[Station] * OP_station_2
        # 判斷為哪個部分配第二工站
        if Station == 1:
            Force_22 = Force  # 靴齒部
        if Station == 2:
            Force_32 = Force  # 鍵槽
        if Station == 3:
            Force_42 = Force  # 柳接孔
        if Station == 4:
            Force_52 = Force  # 中心軸
    # ---------------------------------------------------第一迴圈結束
    # ---------------------------------------------------第二迴圈第三站配全部(除空站)
    for Station in range(1, 1 + (OP_station_Quantity - 2 - Empty_station)):  # Array()從0開始，第二站到倒數第二站，扣除空站
        Force = perimeter[Station] * OP_station_3
        # 判斷為哪個部分配第三工站
        if Station == 1:
            Force_23 = Force  # 靴齒部
        if Station == 2:
            Force_33 = Force  # 鍵槽
        if Station == 3:
            Force_43 = Force  # 柳接孔
        if Station == 4:
            Force_53 = Force  # 中心軸
    # ---------------------------------------------------第二迴圈結束
    # ---------------------------------------------------第三迴圈第四站配全部(除空站)
    for Station in range(1, 1 + (OP_station_Quantity - 2 - Empty_station)):  # Array()從0開始，第二站到倒數第二站，扣除空站
        Force = perimeter[Station] * OP_station_4
        # 判斷為哪個部分配第四工站
        if Station == 1:
            Force_24 = Force  # 靴齒部
        if Station == 2:
            Force_34 = Force  # 鍵槽
        if Station == 3:
            Force_44 = Force  # 柳接孔
        if Station == 4:
            Force_54 = Force  # 中心軸
    # ---------------------------------------------------第三迴圈結束
    # ---------------------------------------------------第四迴圈第五站配全部(除空站)
    for Station in range(1, 1 + (OP_station_Quantity - 2 - Empty_station)):  # Array()從0開始，第二站到倒數第二站，扣除空站
        Force = perimeter[Station] * OP_station_5
        # 判斷為哪個部分配第五工站
        if Station == 1:
            Force_25 = Force  # 靴齒部
        if Station == 2:
            Force_35 = Force  # 鍵槽
        if Station == 3:
            Force_45 = Force  # 柳接孔
        if Station == 4:
            Force_55 = Force  # 中心軸
    # ---------------------------------------------------第四迴圈結束
    # ---------------------------------------------------第五迴圈第六站配全部(除空站)
    if Empty_station == 1:
        for Station in range(1, 1 + (OP_station_Quantity - 2 - Empty_station)):  # Array()從0開始，第二站到倒數第二站，扣除空站
            Force = perimeter[Station] * OP_station_6
            # 判斷為哪個部分配第六工站
            if Station == 1:
                Force_26 = Force  # 靴齒部
            if Station == 2:
                Force_36 = Force  # 鍵槽
            if Station == 3:
                Force_46 = Force  # 柳接孔
            if Station == 4:
                Force_56 = Force  # 中心軸
    # ---------------------------------------------------第五迴圈結束
    # ---------------------------------------------------計算開始
    # Force_類別+站別
    # 類別：1.引導沖，2.靴齒部，3.鍵槽，4.柳接孔，5.中心軸，Max.下料
    # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    if OP_station_Quantity == 7 and Empty_station > 0:  # 如果為七個工站包含一個空站
        Force_end_1 = (
                              Force_11 + 0 + Force_23 + Force_34 + Force_45 + Force_56 + Force_Max) / Force_total  # 引導沖，空站　，靴齒部，鍵槽　，柳接孔，中心軸，下料
        Force_end_2 = (
                              Force_11 + Force_22 + 0 + Force_34 + Force_45 + Force_56 + Force_Max) / Force_total  # 引導沖，靴齒部，空站　，鍵槽　，柳接孔，中心軸，下料
        Force_end_3 = (
                              Force_11 + Force_22 + Force_33 + 0 + Force_45 + Force_56 + Force_Max) / Force_total  # 引導沖，靴齒部，鍵槽　，空站　，柳接孔，中心軸，下料
        Force_end_4 = (
                              Force_11 + Force_22 + Force_33 + Force_44 + 0 + Force_56 + Force_Max) / Force_total  # 引導沖，靴齒部，鍵槽　，柳接孔，空站　，中心軸，下料
        Force_end_5 = (
                              Force_11 + Force_22 + Force_33 + Force_44 + Force_55 + 0 + Force_Max) / Force_total  # 引導沖，靴齒部，鍵槽　，柳接孔，中心軸，空站　，下料
        Force_end_6 = (
                              Force_11 + 0 + Force_33 + Force_24 + Force_45 + Force_56 + Force_Max) / Force_total  # 引導沖，空站　，鍵槽　，靴齒部，柳接孔，中心軸，下料
        Force_end_7 = (
                              Force_11 + Force_32 + 0 + Force_24 + Force_45 + Force_56 + Force_Max) / Force_total  # 引導沖，鍵槽　，空站　，靴齒部，柳接孔，中心軸，下料
        Force_end_8 = (
                              Force_11 + Force_32 + Force_23 + 0 + Force_45 + Force_56 + Force_Max) / Force_total  # 引導沖，鍵槽　，靴齒部，空站　，柳接孔，中心軸，下料
        Force_end_9 = (
                              Force_11 + Force_32 + Force_23 + Force_44 + 0 + Force_56 + Force_Max) / Force_total  # 引導沖，鍵槽　，靴齒部，柳接孔，空站　，中心軸，下料
        Force_end_10 = (
                               Force_11 + Force_32 + Force_23 + Force_44 + Force_55 + 0 + Force_Max) / Force_total  # 引導沖，鍵槽　，靴齒部，柳接孔，中心軸，空站　，下料
        Force_end_11 = (
                               Force_11 + 0 + Force_33 + Force_44 + Force_25 + Force_56 + Force_Max) / Force_total  # 引導沖，空站　，鍵槽　，柳接孔，靴齒部，中心軸，下料
        Force_end_12 = (
                               Force_11 + Force_32 + 0 + Force_44 + Force_25 + Force_56 + Force_Max) / Force_total  # 引導沖，鍵槽　，空站　，柳接孔，靴齒部，中心軸，下料
        Force_end_13 = (
                               Force_11 + Force_32 + Force_43 + 0 + Force_25 + Force_56 + Force_Max) / Force_total  # 引導沖，鍵槽　，柳接孔，空站　，靴齒部，中心軸，下料
        Force_end_14 = (
                               Force_11 + Force_32 + Force_43 + Force_24 + 0 + Force_56 + Force_Max) / Force_total  # 引導沖，鍵槽　，柳接孔，靴齒部，空站　，中心軸，下料
        Force_end_15 = (
                               Force_11 + Force_32 + Force_43 + Force_24 + Force_55 + 0 + Force_Max) / Force_total  # 引導沖，鍵槽　，柳接孔，靴齒部，中心軸，空站　，下料
        Force_end_16 = (
                               Force_11 + 0 + Force_33 + Force_44 + Force_55 + Force_26 + Force_Max) / Force_total  # 引導沖，空站　，鍵槽　，柳接孔，中心軸，靴齒部，下料
        Force_end_17 = (
                               Force_11 + Force_32 + 0 + Force_44 + Force_55 + Force_26 + Force_Max) / Force_total  # 引導沖，鍵槽　，空站　，柳接孔，中心軸，靴齒部，下料
        Force_end_18 = (
                               Force_11 + Force_32 + Force_43 + 0 + Force_55 + Force_26 + Force_Max) / Force_total  # 引導沖，鍵槽　，柳接孔，空站　，中心軸，靴齒部，下料
        Force_end_19 = (
                               Force_11 + Force_32 + Force_43 + Force_54 + 0 + Force_26 + Force_Max) / Force_total  # 引導沖，鍵槽　，柳接孔，中心軸，空站　，靴齒部，下料
        Force_end_20 = (
                               Force_11 + Force_32 + Force_43 + Force_54 + Force_25 + 0 + Force_Max) / Force_total  # 引導沖，鍵槽　，柳接孔，中心軸，靴齒部，空站　，下料
        Force_end_21 = (
                               Force_11 + 0 + Force_23 + Force_44 + Force_35 + Force_56 + Force_Max) / Force_total  # 引導沖，空站　，靴齒部，柳接孔，鍵槽　，中心軸，下料
        Force_end_22 = (
                               Force_11 + Force_22 + 0 + Force_44 + Force_35 + Force_56 + Force_Max) / Force_total  # 引導沖，靴齒部，空站　，柳接孔，鍵槽　，中心軸，下料
        Force_end_23 = (
                               Force_11 + Force_22 + Force_43 + 0 + Force_35 + Force_56 + Force_Max) / Force_total  # 引導沖，靴齒部，柳接孔，空站　，鍵槽　，中心軸，下料
        Force_end_24 = (
                               Force_11 + Force_22 + Force_43 + Force_34 + 0 + Force_56 + Force_Max) / Force_total  # 引導沖，靴齒部，柳接孔，鍵槽　，空站　，中心軸，下料
        Force_end_25 = (
                               Force_11 + Force_22 + Force_43 + Force_34 + Force_55 + 0 + Force_Max) / Force_total  # 引導沖，靴齒部，柳接孔，鍵槽　，中心軸，空站　，下料
        Force_end_26 = (
                               Force_11 + 0 + Force_43 + Force_24 + Force_35 + Force_56 + Force_Max) / Force_total  # 引導沖，空站　，柳接孔，靴齒部，鍵槽　，中心軸，下料
        Force_end_27 = (
                               Force_11 + Force_42 + 0 + Force_24 + Force_35 + Force_56 + Force_Max) / Force_total  # 引導沖，柳接孔，空站　，靴齒部，鍵槽　，中心軸，下料
        Force_end_28 = (
                               Force_11 + Force_42 + Force_23 + 0 + Force_35 + Force_56 + Force_Max) / Force_total  # 引導沖，柳接孔，靴齒部，空站　，鍵槽　，中心軸，下料
        Force_end_29 = (
                               Force_11 + Force_42 + Force_23 + Force_34 + 0 + Force_56 + Force_Max) / Force_total  # 引導沖，柳接孔，靴齒部，鍵槽　，空站　，中心軸，下料
        Force_end_30 = (
                               Force_11 + Force_42 + Force_23 + Force_34 + Force_55 + 0 + Force_Max) / Force_total  # 引導沖，柳接孔，靴齒部，鍵槽　，中心軸，空站　，下料
        Force_end_31 = (
                               Force_11 + 0 + Force_23 + Force_34 + Force_55 + Force_46 + Force_Max) / Force_total  # 引導沖，空站　，靴齒部，鍵槽　，中心軸，柳接孔，下料
        Force_end_32 = (
                               Force_11 + Force_22 + 0 + Force_34 + Force_55 + Force_46 + Force_Max) / Force_total  # 引導沖，靴齒部，空站　，鍵槽　，中心軸，柳接孔，下料
        Force_end_33 = (
                               Force_11 + Force_22 + Force_33 + 0 + Force_55 + Force_46 + Force_Max) / Force_total  # 引導沖，靴齒部，鍵槽　，空站　，中心軸，柳接孔，下料
        Force_end_34 = (
                               Force_11 + Force_22 + Force_33 + Force_54 + 0 + Force_46 + Force_Max) / Force_total  # 引導沖，靴齒部，鍵槽　，中心軸，空站　，柳接孔，下料
        Force_end_35 = (
                               Force_11 + Force_22 + Force_33 + Force_54 + Force_45 + 0 + Force_Max) / Force_total  # 引導沖，靴齒部，鍵槽　，中心軸，柳接孔，空站　，下料
        Force_end_36 = (
                               Force_11 + 0 + Force_33 + Force_54 + Force_25 + Force_46 + Force_Max) / Force_total  # 引導沖，空站　，鍵槽　，中心軸，靴齒部，柳接孔，下料
        Force_end_37 = (
                               Force_11 + Force_32 + 0 + Force_54 + Force_25 + Force_46 + Force_Max) / Force_total  # 引導沖，鍵槽　，空站　，中心軸，靴齒部，柳接孔，下料
        Force_end_38 = (
                               Force_11 + Force_32 + Force_53 + 0 + Force_25 + Force_46 + Force_Max) / Force_total  # 引導沖，鍵槽　，中心軸，空站　，靴齒部，柳接孔，下料
        Force_end_39 = (
                               Force_11 + Force_32 + Force_53 + Force_24 + 0 + Force_46 + Force_Max) / Force_total  # 引導沖，鍵槽　，中心軸，靴齒部，空站　，柳接孔，下料
        Force_end_40 = (
                               Force_11 + Force_32 + Force_53 + Force_24 + Force_45 + 0 + Force_Max) / Force_total  # 引導沖，鍵槽　，中心軸，靴齒部，柳接孔，空站　，下料
        Force_end_41 = (
                               Force_11 + 0 + Force_33 + Force_54 + Force_45 + Force_26 + Force_Max) / Force_total  # 引導沖，空站　，鍵槽　，中心軸，柳接孔，靴齒部，下料
        Force_end_42 = (
                               Force_11 + Force_32 + 0 + Force_54 + Force_45 + Force_26 + Force_Max) / Force_total  # 引導沖，鍵槽　，空站　，中心軸，柳接孔，靴齒部，下料
        Force_end_43 = (
                               Force_11 + Force_32 + Force_53 + 0 + Force_45 + Force_26 + Force_Max) / Force_total  # 引導沖，鍵槽　，中心軸，空站　，柳接孔，靴齒部，下料
        Force_end_44 = (
                               Force_11 + Force_32 + Force_53 + Force_44 + 0 + Force_26 + Force_Max) / Force_total  # 引導沖，鍵槽　，中心軸，柳接孔，空站　，靴齒部，下料
        Force_end_45 = (
                               Force_11 + Force_32 + Force_53 + Force_44 + Force_25 + 0 + Force_Max) / Force_total  # 引導沖，鍵槽　，中心軸，柳接孔，靴齒部，空站　，下料
        Force_end_46 = (
                               Force_11 + 0 + Force_33 + Force_24 + Force_55 + Force_46 + Force_Max) / Force_total  # 引導沖，空站　，鍵槽　，靴齒部，中心軸，柳接孔，下料
        Force_end_47 = (
                               Force_11 + Force_32 + 0 + Force_24 + Force_55 + Force_46 + Force_Max) / Force_total  # 引導沖，鍵槽　，空站　，靴齒部，中心軸，柳接孔，下料
        Force_end_48 = (
                               Force_11 + Force_32 + Force_23 + 0 + Force_55 + Force_46 + Force_Max) / Force_total  # 引導沖，鍵槽　，靴齒部，空站　，中心軸，柳接孔，下料
        Force_end_49 = (
                               Force_11 + Force_32 + Force_23 + Force_54 + 0 + Force_46 + Force_Max) / Force_total  # 引導沖，鍵槽　，靴齒部，中心軸，空站　，柳接孔，下料
        Force_end_50 = (
                               Force_11 + Force_32 + Force_23 + Force_54 + Force_45 + 0 + Force_Max) / Force_total  # 引導沖，鍵槽　，靴齒部，中心軸，柳接孔，空站　，下料
        Force_end_51 = (
                               Force_11 + 0 + Force_43 + Force_34 + Force_25 + Force_56 + Force_Max) / Force_total  # 引導沖，空站　，柳接孔，鍵槽　，靴齒部，中心軸，下料
        Force_end_52 = (
                               Force_11 + Force_42 + 0 + Force_34 + Force_25 + Force_56 + Force_Max) / Force_total  # 引導沖，柳接孔，空站　，鍵槽　，靴齒部，中心軸，下料
        Force_end_53 = (
                               Force_11 + Force_42 + Force_33 + 0 + Force_25 + Force_56 + Force_Max) / Force_total  # 引導沖，柳接孔，鍵槽　，空站　，靴齒部，中心軸，下料
        Force_end_54 = (
                               Force_11 + Force_42 + Force_33 + Force_24 + 0 + Force_56 + Force_Max) / Force_total  # 引導沖，柳接孔，鍵槽　，靴齒部，空站　，中心軸，下料
        Force_end_55 = (
                               Force_11 + Force_42 + Force_33 + Force_24 + Force_55 + 0 + Force_Max) / Force_total  # 引導沖，柳接孔，鍵槽　，靴齒部，中心軸，空站　，下料
        Force_end_56 = (
                               Force_11 + 0 + Force_43 + Force_34 + Force_55 + Force_26 + Force_Max) / Force_total  # 引導沖，空站　，柳接孔，鍵槽　，中心軸，靴齒部，下料
        Force_end_57 = (
                               Force_11 + Force_42 + 0 + Force_34 + Force_55 + Force_26 + Force_Max) / Force_total  # 引導沖，柳接孔，空站　，鍵槽　，中心軸，靴齒部，下料
        Force_end_58 = (
                               Force_11 + Force_42 + Force_33 + 0 + Force_55 + Force_26 + Force_Max) / Force_total  # 引導沖，柳接孔，鍵槽　，空站　，中心軸，靴齒部，下料
        Force_end_59 = (
                               Force_11 + Force_42 + Force_33 + Force_54 + 0 + Force_26 + Force_Max) / Force_total  # 引導沖，柳接孔，鍵槽　，中心軸，空站　，靴齒部，下料
        Force_end_60 = (
                               Force_11 + Force_42 + Force_33 + Force_54 + Force_25 + 0 + Force_Max) / Force_total  # 引導沖，柳接孔，鍵槽　，中心軸，靴齒部，空站　，下料
        Force_all = [Force_end_1, Force_end_2, Force_end_3, Force_end_4, Force_end_5, Force_end_6, Force_end_7,
                     Force_end_8, Force_end_9, Force_end_10, Force_end_11, Force_end_12, Force_end_13, Force_end_14,
                     Force_end_15, Force_end_16, Force_end_17, Force_end_18, Force_end_19, Force_end_20, Force_end_21,
                     Force_end_22, Force_end_23, Force_end_24, Force_end_25, Force_end_26, Force_end_27, Force_end_28,
                     Force_end_29, Force_end_30, Force_end_31, Force_end_32, Force_end_33, Force_end_34, Force_end_35,
                     Force_end_36, Force_end_37, Force_end_38, Force_end_39, Force_end_40, Force_end_41, Force_end_42,
                     Force_end_43, Force_end_44, Force_end_45, Force_end_46, Force_end_47, Force_end_48, Force_end_49,
                     Force_end_50, Force_end_51, Force_end_52, Force_end_53, Force_end_54, Force_end_55, Force_end_56,
                     Force_end_57, Force_end_58, Force_end_59, Force_end_60]
        # --------------------------------------------計算最小值
        Force_min = abs(Force_all[0])
        for end_number in range(0, len(Force_all)):
            if Force_min > abs(Force_all[end_number]):
                Force_min = abs(Force_all[end_number])
        # --------------------------------------------
        # --------------------------------------------計算最小值為哪個組合
        Force_number = Force_min
        for number in range(0, len(Force_all)):
            if Force_number == abs(Force_all[number]):
                Force_combination = number
        # --------------------------------------------
        # --------------------------------------------計算工站順序
        OP_Order = Force_combination + 1
        # --------------------------------------------
        # --------------------------------改變孔位
        # OP_Empty[4] = 鍵槽
        # OP_Empty[5] = 靴齒部
        # OP_Empty[6] = 鉚接孔
        # OP_Empty[7] = 中心孔
        # -------------------------------
        # ----------------------------------------------------------------------------------------輸出順序
        # -------------------------------------------第二站
        if OP_Order == 2 or OP_Order == 3 or OP_Order == 4 or OP_Order == 5 or OP_Order == 22 or OP_Order == 23 or OP_Order == 24 or OP_Order == 25 or OP_Order == 32 or OP_Order == 33 or OP_Order == 34 or OP_Order == 35:
            OP_Empty[5] = 2  # 挖除靴齒部
        if OP_Order == 7 or OP_Order == OP_Order == 8 or OP_Order == 9 or OP_Order == 10 or OP_Order == 12 or OP_Order == 13 or OP_Order == 14 or OP_Order == 15 or OP_Order == 17 or OP_Order == 18 or OP_Order == 19 or OP_Order == 20 or OP_Order == 37 or OP_Order == 38 or OP_Order == 39 or OP_Order == 40 or OP_Order == 42 or OP_Order == 43 or OP_Order == 44 or OP_Order == 45 or OP_Order == 47 or OP_Order == 48 or OP_Order == 49 or OP_Order == 50:
            OP_Empty[4] = 2  # 挖除鍵槽
        if OP_Order == 27 or OP_Order == 28 or OP_Order == 29 or OP_Order == 30 or OP_Order == 52 or OP_Order == 53 or OP_Order == 54 or OP_Order == 55 or OP_Order == 57 or OP_Order == 58 or OP_Order == 59 or OP_Order == 60:
            OP_Empty[6] = 2  # 挖除鉚接孔
        # -------------------------------------------第三站
        if OP_Order == 1 or OP_Order == 8 or OP_Order == 9 or OP_Order == 10 or OP_Order == 21 or OP_Order == 28 or OP_Order == 29 or OP_Order == 30 or OP_Order == 31 or OP_Order == 48 or OP_Order == 49 or OP_Order == 50:
            OP_Empty[5] = 3  # 挖除靴齒部
        if OP_Order == 3 or OP_Order == 4 or OP_Order == 5 or OP_Order == 6 or OP_Order == 11 or OP_Order == 16 or OP_Order == 33 or OP_Order == 34 or OP_Order == 35 or OP_Order == 36 or OP_Order == 41 or OP_Order == 46 or OP_Order == 53 or OP_Order == 54 or OP_Order == 55 or OP_Order == 58 or OP_Order == 59 or OP_Order == 60:
            OP_Empty[4] = 3  # 挖除鍵槽
        if OP_Order == 13 or OP_Order == 14 or OP_Order == 15 or OP_Order == 18 or OP_Order == 19 or OP_Order == 20 or OP_Order == 23 or OP_Order == 24 or OP_Order == 25 or OP_Order == 26 or OP_Order == 51 or OP_Order == 56:
            OP_Empty[6] = 3  # 挖除鉚接孔
        if OP_Order == 38 or OP_Order == 39 or OP_Order == 40 or OP_Order == 43 or OP_Order == 44 or OP_Order == 45:
            OP_Empty[7] = 3  # 挖除中心孔
        # -------------------------------------------第四站
        if OP_Order == 6 or OP_Order == 7 or OP_Order == 14 or OP_Order == 15 or OP_Order == 26 or OP_Order == 27 or OP_Order == 39 or OP_Order == 40 or OP_Order == 46 or OP_Order == 47 or OP_Order == 54 or OP_Order == 55:
            OP_Empty[5] = 4  # 挖除靴齒部
        if OP_Order == 1 or OP_Order == 2 or OP_Order == 24 or OP_Order == 25 or OP_Order == 29 or OP_Order == 30 or OP_Order == 31 or OP_Order == 32 or OP_Order == 51 or OP_Order == 52 or OP_Order == 56 or OP_Order == 57:
            OP_Empty[4] = 4  # 挖除鍵槽
        if OP_Order == 4 or OP_Order == 5 or OP_Order == 9 or OP_Order == 10 or OP_Order == 11 or OP_Order == 12 or OP_Order == 16 or OP_Order == 17 or OP_Order == 21 or OP_Order == 22 or OP_Order == 44 or OP_Order == 45:
            OP_Empty[6] = 4  # 挖除鉚接孔
        if OP_Order == 19 or OP_Order == 20 or OP_Order == 34 or OP_Order == 35 or OP_Order == 36 or OP_Order == 37 or OP_Order == 41 or OP_Order == 42 or OP_Order == 49 or OP_Order == 50 or OP_Order == 59 or OP_Order == 60:
            OP_Empty[7] = 4  # 挖除中心孔
        # -------------------------------------------第五站
        if OP_Order == 11 or OP_Order == 12 or OP_Order == 13 or OP_Order == 20 or OP_Order == 36 or OP_Order == 37 or OP_Order == 38 or OP_Order == 45 or OP_Order == 51 or OP_Order == 52 or OP_Order == 53 or OP_Order == 60:
            OP_Empty[5] = 5  # 挖除靴齒部
        if OP_Order == 21 or OP_Order == 22 or OP_Order == 23 or OP_Order == 26 or OP_Order == 27 or OP_Order == 28:
            OP_Empty[4] = 5  # 挖除鍵槽
        if OP_Order == 1 or OP_Order == 2 or OP_Order == 3 or OP_Order == 6 or OP_Order == 7 or OP_Order == 8 or OP_Order == 35 or OP_Order == 40 or OP_Order == 41 or OP_Order == 42 or OP_Order == 43 or OP_Order == 50:
            OP_Empty[6] = 5  # 挖除鉚接孔
        if OP_Order == 5 or OP_Order == 10 or OP_Order == 15 or OP_Order == 16 or OP_Order == 17 or OP_Order == 18 or OP_Order == 25 or OP_Order == 30 or OP_Order == 31 or OP_Order == 32 or OP_Order == 33 or OP_Order == 46 or OP_Order == 47 or OP_Order == 48 or OP_Order == 55 or OP_Order == 56 or OP_Order == 57 or OP_Order == 58:
            OP_Empty[7] = 5  # 挖除中心孔
        # -------------------------------------------第六站
        if OP_Order == 16 or OP_Order == 17 or OP_Order == 18 or OP_Order == 19 or OP_Order == 41 or OP_Order == 42 or OP_Order == 43 or OP_Order == 44 or OP_Order == 56 or OP_Order == 57 or OP_Order == 58 or OP_Order == 59:
            OP_Empty[5] = 6  # 挖除靴齒部
        if OP_Order == 31 or OP_Order == 32 or OP_Order == 33 or OP_Order == 34 or OP_Order == 36 or OP_Order == 37 or OP_Order == 38 or OP_Order == 39 or OP_Order == 46 or OP_Order == 47 or OP_Order == 48 or OP_Order == 49:
            OP_Empty[6] = 6  # 挖除鉚接孔
        if OP_Order == 1 or OP_Order == 2 or OP_Order == 3 or OP_Order == 4 or OP_Order == 6 or OP_Order == 7 or OP_Order == 8 or OP_Order == 9 or OP_Order == 11 or OP_Order == 12 or OP_Order == 13 or OP_Order == 14 or OP_Order == 21 or OP_Order == 22 or OP_Order == 23 or OP_Order == 24 or OP_Order == 26 or OP_Order == 27 or OP_Order == 28 or OP_Order == 29 or OP_Order == 51 or OP_Order == 52 or OP_Order == 53 or OP_Order == 54:
            OP_Empty[7] = 6  # 挖除中心孔
        # ----------------------------------------------------------------------------------------
    # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    if OP_station_Quantity == 6 and Empty_station == 0:  # 如果為六個工站沒有空站
        Force_end_1 = (
                              Force_11 + Force_22 + Force_33 + Force_44 + Force_55 + Force_Max) / Force_total  # 引導沖，靴齒部，鍵槽　，柳接孔，中心軸，下料
        Force_end_2 = (
                              Force_11 + Force_32 + Force_23 + Force_44 + Force_55 + Force_Max) / Force_total  # 引導沖，空站　，鍵槽　，靴齒部，柳接孔，中心軸，下料
        Force_end_3 = (
                              Force_11 + Force_32 + Force_43 + Force_24 + Force_55 + Force_Max) / Force_total  # 引導沖，鍵槽　，柳接孔，靴齒部，中心軸，下料
        Force_end_4 = (
                              Force_11 + Force_32 + Force_43 + Force_54 + Force_25 + Force_Max) / Force_total  # 引導沖，鍵槽　，柳接孔，中心軸，靴齒部，下料
        Force_end_5 = (
                              Force_11 + Force_22 + Force_43 + Force_34 + Force_55 + Force_Max) / Force_total  # 引導沖，靴齒部，柳接孔，鍵槽　，中心軸，下料
        Force_end_6 = (
                              Force_11 + Force_42 + Force_23 + Force_34 + Force_55 + Force_Max) / Force_total  # 引導沖，柳接孔，靴齒部，鍵槽　，中心軸，下料
        Force_end_7 = (
                              Force_11 + Force_22 + Force_33 + Force_54 + Force_45 + Force_Max) / Force_total  # 引導沖，靴齒部，鍵槽　，中心軸，柳接孔，下料
        Force_end_8 = (
                              Force_11 + Force_32 + Force_53 + Force_24 + Force_45 + Force_Max) / Force_total  # 引導沖，鍵槽　，中心軸，靴齒部，柳接孔，下料
        Force_end_9 = (
                              Force_11 + Force_32 + Force_53 + Force_44 + Force_25 + Force_Max) / Force_total  # 引導沖，鍵槽　，中心軸，柳接孔，靴齒部，下料
        Force_end_10 = (
                               Force_11 + Force_32 + Force_23 + Force_54 + Force_45 + Force_Max) / Force_total  # 引導沖，鍵槽　，靴齒部，中心軸，柳接孔，下料
        Force_end_11 = (
                               Force_11 + Force_42 + Force_33 + Force_24 + Force_55 + Force_Max) / Force_total  # 引導沖，柳接孔，鍵槽　，靴齒部，中心軸，下料
        Force_end_12 = (
                               Force_11 + Force_42 + Force_33 + Force_54 + Force_25 + Force_Max) / Force_total  # 引導沖，柳接孔，鍵槽　，中心軸，靴齒部，下料
        ##Force_all As Variant
        Force_all = [Force_end_1, Force_end_2, Force_end_3, Force_end_4, Force_end_5, Force_end_6, Force_end_7,
                     Force_end_8, Force_end_9, Force_end_10, Force_end_11, Force_end_12]
        # --------------------------------------------計算最小值
        Force_min = abs(Force_all(0))
        for end_number in range(0, len(Force_all)):
            if Force_min > abs(Force_all[end_number]):
                Force_min = abs(Force_all[end_number])
        # --------------------------------------------
        # --------------------------------------------計算最小值為哪個組合
        Force_number = Force_min
        for number in range(0, 1 + len(Force_all)):
            if Force_number == abs(Force_all[number]):
                Force_combination = number
        # --------------------------------------------
        # --------------------------------------------計算工站順序
        OP_Order = Force_combination + 1
        # --------------------------------------------
        # --------------------------------------------------------------------------輸出順序
        if OP_Order == 1:
            # --------------------------------改變孔位
            OP_Empty[5] = 2  # 挖除靴齒部
            OP_Empty[4] = 3  # 挖除鍵槽
            OP_Empty[6] = 4  # 挖除鉚接孔
            OP_Empty[7] = 5  # 挖除中心孔
        # -------------------------------
        if OP_Order == 2:
            # --------------------------------改變孔位
            OP_Empty[5] = 3  # 挖除靴齒部
            OP_Empty[4] = 2  # 挖除鍵槽
            OP_Empty[6] = 4  # 挖除鉚接孔
            OP_Empty[7] = 5  # 挖除中心孔
        # -------------------------------
        if OP_Order == 3:
            # --------------------------------改變孔位
            OP_Empty[5] = 4  # 挖除靴齒部
            OP_Empty[4] = 2  # 挖除鍵槽
            OP_Empty[6] = 3  # 挖除鉚接孔
            OP_Empty[7] = 5  # 挖除中心孔
        # -------------------------------
        if OP_Order == 4:
            # --------------------------------改變孔位
            OP_Empty[5] = 5  # 挖除靴齒部
            OP_Empty[4] = 2  # 挖除鍵槽
            OP_Empty[6] = 3  # 挖除鉚接孔
            OP_Empty[7] = 4  # 挖除中心孔
        # -------------------------------
        if OP_Order == 5:
            # --------------------------------改變孔位
            OP_Empty[5] = 2  # 挖除靴齒部
            OP_Empty[4] = 4  # 挖除鍵槽
            OP_Empty[6] = 3  # 挖除鉚接孔
            OP_Empty[7] = 5  # 挖除中心孔
        # -------------------------------
        if OP_Order == 6:
            # --------------------------------改變孔位
            OP_Empty[5] = 3  # 挖除靴齒部
            OP_Empty[4] = 4  # 挖除鍵槽
            OP_Empty[6] = 2  # 挖除鉚接孔
            OP_Empty[7] = 5  # 挖除中心孔
        # -------------------------------
        if OP_Order == 7:
            # --------------------------------改變孔位
            OP_Empty[5] = 2  # 挖除靴齒部
            OP_Empty[4] = 3  # 挖除鍵槽
            OP_Empty[6] = 5  # 挖除鉚接孔
            OP_Empty[7] = 4  # 挖除中心孔
        # -------------------------------
        if OP_Order == 8:
            # --------------------------------改變孔位
            OP_Empty[5] = 4  # 挖除靴齒部
            OP_Empty[4] = 2  # 挖除鍵槽
            OP_Empty[6] = 5  # 挖除鉚接孔
            OP_Empty[7] = 3  # 挖除中心孔
        # -------------------------------
        if OP_Order == 9:
            # --------------------------------改變孔位
            OP_Empty[5] = 5  # 挖除靴齒部
            OP_Empty[4] = 2  # 挖除鍵槽
            OP_Empty[6] = 4  # 挖除鉚接孔
            OP_Empty[7] = 3  # 挖除中心孔
        # -------------------------------
        if OP_Order == 10:
            # --------------------------------改變孔位
            OP_Empty[5] = 3  # 挖除靴齒部
            OP_Empty[4] = 2  # 挖除鍵槽
            OP_Empty[6] = 5  # 挖除鉚接孔
            OP_Empty[7] = 4  # 挖除中心孔
        # -------------------------------
        if OP_Order == 11:
            # --------------------------------改變孔位
            OP_Empty[5] = 4  # 挖除靴齒部
            OP_Empty[4] = 3  # 挖除鍵槽
            OP_Empty[6] = 2  # 挖除鉚接孔
            OP_Empty[7] = 5  # 挖除中心孔
        # -------------------------------
        if OP_Order == 12:
            # --------------------------------改變孔位
            OP_Empty[5] = 5  # 挖除靴齒部
            OP_Empty[4] = 3  # 挖除鍵槽
            OP_Empty[6] = 2  # 挖除鉚接孔
            OP_Empty[7] = 4  # 挖除中心孔
        # -------------------------------
    return OP_Empty


def DataBuild(StripLength, StripWidth):
    gvar.die_type = 'common'
    TotleOP = 7
    PlateLength = StripLength + 100
    PlateWide = StripWidth + 120
    DieLength = PlateLength + 60
    DieWide = PlateWide + 150
    SketchPosition = 'Hybridbody'
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    try:
        document1 = catapp.ActiveDocument
        document1.Close()
    except:
        pass
    PartDocuments = documents1.Open(gvar.open_path + 'Strip_Data-2.CATPart')
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    hybridBodies1 = part1.HybridBodies
    originElements1 = part1.OriginElements
    hybridShapePlaneExplicit1 = originElements1.PlaneXY
    file_name = "Strip_Data-2"
    body_name1 = "PartBody"
    hybridBody_name = "die"
    # =============環境設置===============
    (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set(file_name, body_name1,
                                                                                             hybridBody_name)
    # =============環境設置===============
    part1.InWorkObject = ElementHybridBody
    hybridShape1 = ElementHybridBody.HybridShapes.Item("plate_centor_point")
    # --------------------------------------------------------------------------------------------build lower seat
    (ElementSketch) = defs.BuildSketch("lower_die_Sketch", hybridShapePlaneExplicit1, ElementDocument, SketchPosition,
                                       ElementBody, ElementHybridBody)
    (ElementPoint1, ElementLine1, ElementLine2, ElementLine3, ElementLine4) = defs.SketchRectangle(ElementSketch,
                                                                                                   DieLength, DieWide,
                                                                                                   ElementDocument,
                                                                                                   ElementHybridBody)
    ElementPoint3 = hybridShape1
    # -----------------------------------線段改名最後標註使用
    ElementLine1.Name = "L1"
    ElementLine4.Name = "L2"
    ElementLine2.Name = "L3"
    # -----------------------------------線段改名最後標註使用
    mainsketch1 = ElementSketch
    ElementPoint4 = ElementPoint1
    defs.SketchBuildCallout(ElementSketch, "free", "Binding", 0, ElementDocument, ElementPoint3, ElementPoint4)
    (ElementSketch) = defs.BuildSketch("upper_die_Sketch", hybridShapePlaneExplicit1, ElementDocument, SketchPosition,
                                       ElementBody, ElementHybridBody)
    (ElementPoint1, ElementLine1, ElementLine2, ElementLine3, ElementLine4) = defs.SketchRectangle(ElementSketch,
                                                                                                   DieLength, DieWide,
                                                                                                   ElementDocument,
                                                                                                   ElementHybridBody)
    ElementPoint3 = hybridShape1
    ElementPoint4 = ElementPoint1
    # -----------------------------------線段改名最後標註使用
    ElementLine1.Name = "L1"
    ElementLine4.Name = "L2"
    ElementLine2.Name = "L3"
    # -----------------------------------線段改名最後標註使用
    mainsketch2 = ElementSketch
    defs.SketchBuildCallout(ElementSketch, 'free', 'Binding', 0, ElementDocument, ElementPoint3, ElementPoint4)
    (ElementReference1) = defs.ExtremumPoint("X_min", "Y_min", "Z_max", 2, ElementSketch, ElementDocument, ElementBody,
                                             ElementHybridBody)
    ElementReference11 = ElementReference1
    ElementReference11.Name = "seat_plate_min"  # 建立最小點
    (ElementSketch) = defs.BuildSketch("location_block_1", hybridShapePlaneExplicit1, ElementDocument, SketchPosition,
                                       ElementBody, ElementHybridBody)
    (ElementPoint1, ElementLine1, ElementLine2, ElementLine3, ElementLine4) = defs.SketchRectangle(ElementSketch,
                                                                                                   PlateLength,
                                                                                                   PlateWide,
                                                                                                   ElementDocument,
                                                                                                   ElementHybridBody)
    ElementPoint4 = ElementPoint1
    ElementPoint3 = hybridShape1
    defs.SketchBuildCallout(ElementSketch, 'free', 'Binding', 0, ElementDocument, ElementPoint3, ElementPoint4)
    ElementSketch.Name = "plate_line_1_sketch"
    # -----------------------------------線段改名最後標註使用
    ElementLine1.Name = "L1"
    ElementLine4.Name = "L2"
    ElementLine2.Name = "L3"
    # -----------------------------------線段改名最後標註使用
    mainsketch4 = ElementSketch
    part1.Update()
    (ElementSketch) = defs.BuildSketch("sketch.9", hybridShapePlaneExplicit1, ElementDocument, SketchPosition,
                                       ElementBody, ElementHybridBody)
    defs.MarkLineConstraint(ElementSketch, mainsketch1, mainsketch2, "L1", "L1", "lower_up_die_set_X", ElementDocument,
                            ElementHybridBody)
    defs.MarkLineConstraint(ElementSketch, mainsketch1, mainsketch2, "L2", "L2", "lower_up_die_set_Y", ElementDocument,
                            ElementHybridBody)
    defs.MarkLineConstraint(ElementSketch, mainsketch1, mainsketch4, "L2", "L2", "lower_die_set_lower_die_X",
                            ElementDocument, ElementHybridBody)
    defs.MarkLineConstraint(ElementSketch, mainsketch1, mainsketch4, "L1", "L1", "lower_die_set_lower_die_Y",
                            ElementDocument, ElementHybridBody)
    defs.MarkLineLengthConstraint(ElementSketch, mainsketch4, "L1", "plate_length_1", ElementDocument,
                                  ElementHybridBody)
    defs.MarkLineLengthConstraint(ElementSketch, mainsketch4, "L2", "strip_width", ElementDocument, ElementHybridBody)
    defs.MarkLineAxisConstraint(ElementSketch, mainsketch1, ElementSketch, "L2", "HD", "lower_die_set_pilot_punch",
                                ElementDocument, ElementHybridBody)
    defs.MarkLineAxisConstraint(ElementSketch, mainsketch2, ElementSketch, "L2", "HD", "upper_die_set_pilot_punch",
                                ElementDocument, ElementHybridBody)
    defs.MarkLineAxisPointConstraint(ElementSketch, ElementSketch, "HD", "lower_die_pilot_punch", ElementDocument,
                                     ElementHybridBody)
    # =============隱藏元素===============
    selection1 = ElementDocument.Selection
    visPropertySet1 = selection1.VisProperties
    selection1.Clear()
    time.sleep(1)
    selection1.add(ElementSketch)
    visPropertySet1 = visPropertySet1.Parent
    visPropertySet1.SetShow(1)
    selection1.Clear()
    # =============隱藏元素===============
    # =================產生線段===================
    defs.Project(mainsketch1.Name, "upper_die_seat_line", ElementDocument)
    defs.Project(mainsketch2.Name, "lower_die_seat_line", ElementDocument)
    defs.Project(mainsketch4.Name, "number_1_plate_line", ElementDocument)
    # =================產生線段===================
    final = False
    (final) = defs.Del1("common_plate_Sketch", True, final, ElementDocument)
    (final) = defs.Del1("Hole_Sketch", True, final, ElementDocument)
    for i in range(1, 7):
        final = False
        (final) = defs.Del1("sketch_number_" + str(i) + "_plate_line", True, final, ElementDocument)
        if final == True:
            break
    for i in range(1, 7):
        final = False
        (final) = defs.Del1("location_block_" + str(i), True, final, ElementDocument)
        if final == True:
            break
    for i in range(1, 7):
        final = False
        (final) = defs.Del1("upper_location_block_Sketch" + str(i), True, final, ElementDocument)
        if final == True:
            break
    part1 = ElementDocument.Part
    hybridShapeFactory1 = part1.HybridShapeFactory
    hybridBodies1 = part1.HybridBodies
    hybridBody1 = hybridBodies1.Item("die")
    sel1 = ElementDocument.Selection
    hybridShapes1 = hybridBody1.HybridShapes
    hybridShapeExtremum1 = hybridShapes1.Item("seat_plate_min")
    sel1.Add(hybridShapeExtremum1)
    sel1.Delete()
    SketchPosition = 'Body'
    # -----------------隱藏Data----------------------
    partDocument1 = catapp.ActiveDocument
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    # bodies1 = body1.Parent
    # bSTR1 = str(body1.Name)
    selection1.Add(body1)
    visPropertySet1 = visPropertySet1.Parent
    # bSTR2 = str(visPropertySet1.Name)
    # bSTR3 = str(visPropertySet1.Name)
    visPropertySet1.SetShow(1)
    selection1.Clear()
    selection2 = partDocument1.Selection
    visPropertySet2 = selection2.VisProperties
    hybridBodies1 = part1.HybridBodies
    hybridBody1 = hybridBodies1.Item("die")
    hybridBodies1 = hybridBody1.Parent
    bSTR4 = str(hybridBody1.Name)
    selection2.Add(hybridBody1)
    visPropertySet2 = visPropertySet2.Parent
    # bSTR5 = str(visPropertySet2.Name)
    # bSTR6 = str(visPropertySet2.Name)
    visPropertySet2.SetShow(1)
    selection2.Clear()
    # -----------------隱藏Data----------------------
    time.sleep(2)
    partDocument1.SaveAs(gvar.open_path + "Data1.CATPart")
    return SketchPosition


def DataSetting(SketchPosition, StripWidth):
    PlateLength = [0.0]
    plate_line_demise_surface_up_number_surch = [[0] * 99 for iiii in range(9999)]
    plate_line_half_cut_line_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_reinforcement_cut_line = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_up_inbolt_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_up_inbolt_demise_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_up_outbolt_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_demise_up_outbolt_demise_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_up_outbolt_side_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_up_outbolt_side_demise_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_down_inbolt_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_down_inbolt_demise_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_down_outbolt_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_down_outbolt_demise_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_down_outbolt_side_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_down_outbolt_side_demise_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_up_demise_line_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_leveling_block_down_demise_line_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_bend_up_forming_punch_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_bend_up_emboss_forming_insert_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_emboss_forming_punch_left_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_emboss_forming_punch_right_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_bend_up_punch_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_bend_up_insert_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_bend_up_forming_insert_surface_L_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_bend_up_forming_insert_surface_R_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_cut_punch_d_cutting_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_cut_punch_u_cutting_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_right_quickly_remove_cut_line_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_left_quickly_remove_cut_line_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_up_quickly_remove_cut_line_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_down_quickly_remove_cut_line_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_right_quickly_remove_bending_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_left_quickly_remove_bending_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_up_quickly_remove_bending_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_down_quickly_remove_bending_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_A_punch_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_cut_line_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_allotype_cut_line_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_forming_cavity_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_forming_punch_surface_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_shaping_cavity_surface_down_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_shaping_cavity_surface_up_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_shaping_punch_surface_down_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_shaping_punch_surface_up_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_Bending_cavity_number_up = [[0.0] * 30 for iiii in range(20)]
    plate_line_Bending_punch_number_up = [[0.0] * 30 for iiii in range(20)]
    plate_line_Bending_cavity_floating_number_up = [[0.0] * 30 for iiii in range(20)]
    plate_line_Bending_cavity_number_down = [[0.0] * 30 for iiii in range(20)]
    plate_line_Bending_punch_number_down = [[0.0] * 30 for iiii in range(20)]
    plate_line_Bending_punch_up_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_Bending_punch_down_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_unnomal_cut_line_T_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_unnomal_cut_line_I_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_unnomal_cut_line_M_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_bend_down_forming_insert_surface_L_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_bend_down_forming_insert_surface_R_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_shoulder_bendin_punch_number = [[0.0] * 30 for iiii in range(20)]
    plate_line_shoulder_bendin_cavity_number = [[0.0] * 30 for iiii in range(20)]
    shoulder_bending_grouping_parameter = [[0] * 30 for iiii in range(20)]
    plate_line_shoulder_up_point_A = [[0.0] * 30 for iiii in range(20)]
    plate_line_shoulder_up_point_B = [[0.0] * 30 for iiii in range(20)]
    plate_line_shoulder_down_point_A = [[0.0] * 30 for iiii in range(20)]
    plate_line_shoulder_down_point_B = [[0.0] * 30 for iiii in range(20)]
    plate_line_left_shoulder_emboss_up_point_A = [[0.0] * 30 for iiii in range(20)]
    plate_line_left_shoulder_emboss_up_point_B = [[0.0] * 30 for iiii in range(20)]
    plate_line_left_shoulder_emboss_down_point_A = [[0.0] * 30 for iiii in range(20)]
    plate_line_left_shoulder_emboss_down_point_B = [[0.0] * 30 for iiii in range(20)]
    plate_line_right_shoulder_emboss_up_point_A = [[0.0] * 30 for iiii in range(20)]
    plate_line_right_shoulder_emboss_up_point_B = [[0.0] * 30 for iiii in range(20)]
    plate_line_right_shoulder_emboss_down_point_A = [[0.0] * 30 for iiii in range(20)]
    plate_line_right_shoulder_emboss_down_point_B = [[0.0] * 30 for iiii in range(20)]
    plate_line_bending_punch_surface = [[0.0] * 30 for iiii in range(20)]
    plate_line_bending_cavity_surface = [[0.0] * 30 for iiii in range(20)]
    plate_line_pilot_punch_number = [0.0] * 20
    plate_line_stripper_pin_point_number = [0.0] * 20
    plate_line_LIFTER_point_number = [0.0] * 20
    plate_line_limiting_point_number = [0.0] * 20
    bb = [0.0] * 30
    plate_X_min = [0.0] * 30
    plate_X_max = [0.0] * 30
    plate_Y_min = [0.0] * 30
    plate_Y_max = [0.0] * 30
    plate_length_origin_die = [0.0] * 30
    plate_wide_origin_die = [0.0] * 30
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    TotalOpNumber = int(gvar.strip_parameter_list[2])
    try:
        partDocument1 = catapp.ActiveDocument
    except:
        partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
    selection1 = partDocument1.Selection
    part1 = partDocument1.Part
    # -------------------------------------定義座標差值
    parameters1 = part1.Parameters
    # -----------------------------↓(搜尋plate_line_)
    selection1.Clear()
    selection1.Search("Name=number_*_plate_line,all")
    gvar.PlateLineNumber = selection1.Count
    selection1.Clear()
    # -----------------------------↑
    for i in range(1, gvar.PlateLineNumber + 1):  # 第1迴圈_開始
        # *************************************修改測量方法
        length10 = parameters1.Item("plate_length_" + str(i))
        PlateLength.append(length10.Value)
        # *************************************修改測量方法
        for tt in range(1, TotalOpNumber + 1):  # 第2迴圈_開始
            op_number = 10 * tt
            # ------------------------------------------------↓模板讓位
            # *************************************修改測數量方法
            for j in range(1, 11):  # 第3迴圈_開始
                for demise_h in range(1, 11):  # 第4迴圈_開始
                    plate_op = i * 100 + tt
                    # -----------------------------↓(搜尋每塊模板以及每個工程的_demise_surface_u_*H_) 讓位
                    selection1.Clear()
                    selection1.Search(
                        "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_demise_surface_up_" + str(
                            demise_h) + "H_" + str(j) + ",all")
                    plate_line_demise_surface_up_number_surch[plate_op][j] = selection1.Count
                    if plate_line_demise_surface_up_number_surch[plate_op][j] > 0:
                        plate_line_demise_surface_up_number_surch[plate_op][j] = demise_h
                    selection1.Clear()
                    # -----------------------------↓(搜尋每塊模板以及每個工程的_demise_surface_d_*H_) 讓位
                    selection1.Clear()
                    selection1.Search(
                        "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_demise_surface_down_" + str(
                            demise_h) + "H_" + str(j) + ",all")
                    plate_line_demise_surface_up_number_surch[plate_op][j] = selection1.Count
                    if plate_line_demise_surface_up_number_surch[plate_op][j] > 0:
                        plate_line_demise_surface_up_number_surch[plate_op][j] = demise_h
                    selection1.Clear()
                    # 第4迴圈_結束
                # 第3迴圈_結束
            # *************************************修改測數量方法
            # ------------------------------------------------↑模板讓位
            # -----------------------------↓(搜尋每塊模板以及每個工程的_half_cut_line_)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_half_cut_line_*,all")
            plate_line_half_cut_line_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的Reinforcement_cut_line)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_Reinforcement_cut_line_*,all")
            plate_line_reinforcement_cut_line[i][tt] = selection1.Count
            selection1.Clear()
            # ------------------------------------------------↓整平塊
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_up_inbolt_) 上整平塊_內螺栓
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_leveling_block_up_inbolt_surface_*,all")
            plate_line_leveling_block_up_inbolt_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_up_inbolt_demise_) 上整平塊_內螺栓_中間讓位
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_leveling_block_up_inbolt_demise_surface_*,all")
            plate_line_leveling_block_up_inbolt_demise_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_up_outbolt_) 上整平塊_外螺栓
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_leveling_block_up_outbolt_surface_*,all")
            plate_line_leveling_block_up_outbolt_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_up_outbolt_demise_) 上整平塊_外螺栓_中間讓位
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(
                op_number) + "_leveling_block_up_outbolt_demise_surface_*,all")
            plate_line_leveling_block_demise_up_outbolt_demise_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_up_outbolt_side_) 上整平塊_外螺栓_旁邊
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_leveling_block_up_outbolt_side_surface_*,all")
            plate_line_leveling_block_up_outbolt_side_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_up_outbolt_side_demise_) 上整平塊_外螺栓_旁邊_中間讓位
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(
                op_number) + "_leveling_block_up_outbolt_side_demise_surface_*,all")
            plate_line_leveling_block_up_outbolt_side_demise_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_down_inbolt_) 下整平塊_內螺栓
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_leveling_block_down_inbolt_surface_*,all")
            plate_line_leveling_block_down_inbolt_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_down_inbolt_demise_) 下整平塊_內螺栓_中間讓位
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(
                op_number) + "_leveling_block_down_inbolt_demise_surface_*,all")
            plate_line_leveling_block_down_inbolt_demise_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_down_outbolt_) 下整平塊_外螺栓
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_leveling_block_down_outbolt_surface_*,all")
            plate_line_leveling_block_down_outbolt_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_down_outbolt_demise_) 下整平塊_外螺栓_中間讓位
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(
                op_number) + "_leveling_block_down_outbolt_demise_surface_*,all")
            plate_line_leveling_block_down_outbolt_demise_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_down_outbolt_side_) 下整平塊_外螺栓_旁邊
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(
                op_number) + "_leveling_block_down_outbolt_side_surface_*,all")
            plate_line_leveling_block_down_outbolt_side_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_down_outbolt_side_) 下整平塊_外螺栓_旁邊_中間讓位
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(
                op_number) + "_leveling_block_down_outbolt_side_demise_surface_*,all")
            plate_line_leveling_block_down_outbolt_side_demise_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_lower_) 上整平塊模板讓位
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_leveling_block_up_demise_line_*,all")
            plate_line_leveling_block_up_demise_line_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_leveling_block_lower_) 下整平塊模板讓位
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_leveling_block_down_demise_line_*,all")
            plate_line_leveling_block_down_demise_line_number[i][tt] = selection1.Count
            selection1.Clear()
            # ------------------------------------------------------------------------------------------------------------------------↑整平塊
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bend_up_forming_punch_surface_) 向上折彎整形沖頭
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bend_up_forming_punch_surface_*,all")
            plate_line_bend_up_forming_punch_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bend_up_emboss_insert_surface_) 向上折彎整形入子(打凸包浮塊)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bend_up_emboss_insert_surface_*,all")
            plate_line_bend_up_emboss_forming_insert_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_emboss_forming_punch_left_surface_) 打凸包沖頭_左
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_emboss_forming_punch_left_surface_*,all")
            plate_line_emboss_forming_punch_left_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_emboss_forming_punch_right_surface_) 打凸包沖頭_右
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_emboss_forming_punch_right_surface_*,all")
            plate_line_emboss_forming_punch_right_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bend_up_punch_surface_) 向上折彎_沖頭
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bend_up_punch_surface_*,all")
            plate_line_bend_up_punch_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bend_up_insert_surface_) 向上折彎_入塊
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bend_up_insert_surface_*,all")
            plate_line_bend_up_insert_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bend_up_forming_punch_surface_) 向上折彎_成形沖頭_左
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bend_up_forming_insert_surface_L_*,all")
            plate_line_bend_up_forming_insert_surface_L_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bend_up_forming_punch_surface_) 向上折彎_成形沖頭_右
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bend_up_forming_insert_surface_R_*,all")
            plate_line_bend_up_forming_insert_surface_R_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_cut_punch_d_cutting_) #切斷沖頭_下
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_cut_punch_d_cutting_*,all")
            plate_line_cut_punch_d_cutting_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_cut_punch_u_cutting_)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_cut_punch_u_cutting_*,all")  ##切斷沖頭_上
            plate_line_cut_punch_u_cutting_number[i][tt] = selection1.Count
            selection1.Clear()
            # --------------------------------------------------------------------------------------------↓快拆沖頭
            # -----------------------------↓(搜尋每塊模板以及每個工程的_right_quickly_remove_cut_line)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_right_quickly_remove_cut_line_*,all")
            plate_line_right_quickly_remove_cut_line_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_left_quickly_remove_cut_line)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_left_quickly_remove_cut_line_*,all")
            plate_line_left_quickly_remove_cut_line_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_up_quickly_remove_cut_line)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_up_quickly_remove_cut_line_*,all")
            plate_line_up_quickly_remove_cut_line_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_down_quickly_remove_cut_line)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_down_quickly_remove_cut_line_*,all")
            plate_line_down_quickly_remove_cut_line_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_right_quickly_remove_bending_surface)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_right_quickly_remove_bending_surface_*,all")
            plate_line_right_quickly_remove_bending_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_left_quickly_remove_bending_surface)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_left_quickly_remove_bending_surface_*,all")
            plate_line_left_quickly_remove_bending_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_up_quickly_remove_bending_surface)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_up_quickly_remove_bending_surface_*,all")
            plate_line_up_quickly_remove_bending_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_down_quickly_remove_bending_surface)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_down_quickly_remove_bending_surface_*,all")
            plate_line_down_quickly_remove_bending_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # --------------------------------------------------------------------------------------------↑快拆沖頭
            # -----------------------------↓(搜尋每塊模板以及每個工程的_A_punch)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_A_punch_*,all")
            plate_line_A_punch_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_cut_line_)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_cut_line_*,all")
            plate_line_cut_line_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_allotype_cut_line_)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_allotype_cut_line_*,all")
            plate_line_allotype_cut_line_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_forming_cavity_surface_)成形模穴
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_forming_cavity_surface_*,all")
            plate_line_forming_cavity_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_forming_punch_surface_)成形沖頭
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_forming_punch_surface_*,all")
            plate_line_forming_punch_surface_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_shaping_cavity_surface_down_)下塑形模穴
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shaping_cavity_surface_down_*,all")
            plate_line_shaping_cavity_surface_down_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_shaping_cavity_surface_up_)上塑形模穴
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shaping_cavity_surface_up_*,all")
            plate_line_shaping_cavity_surface_up_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_shaping_punch_surface_down_)下塑形沖頭
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shaping_punch_surface_down_*,all")
            plate_line_shaping_punch_surface_down_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_shaping_punch_surface_up_)上塑形沖頭
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shaping_punch_surface_up_*,all")
            plate_line_shaping_punch_surface_up_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_Bending_up_cavity_surface_)彎折模穴
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bending_up_cavity_surface_*,all")
            plate_line_Bending_cavity_number_up[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_Bending_up_punch_surface_)彎折沖頭
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bending_up_punch_surface_*,all")
            plate_line_Bending_punch_number_up[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_Bending_up_cavity_floating_)浮動彎折模穴
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bending_up_cavity_floating_*,all")
            plate_line_Bending_cavity_floating_number_up[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_Bending_down_cavity_surface_)下彎折模穴
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bending_down_cavity_surface_*,all")
            plate_line_Bending_cavity_number_down[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_Bending_down_punch_surface_)下彎折沖頭
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bending_down_punch_surface_*,all")
            plate_line_Bending_punch_number_down[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bending_surface_up_)上彎折面
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bending_surface_up_*,all")
            plate_line_Bending_punch_up_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bending_surface_down_)下彎折面
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bending_surface_down_*,all")
            plate_line_Bending_punch_down_number[i][tt] = selection1.Count
            selection1.Clear()
            # ---------------------------------------------------------------------↓異型沖
            # -----------------------------↓(搜尋每塊模板以及每個工程的_unnomal_cut_line_T_異型沖)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_unnomal_cut_line_T_*,all")
            plate_line_unnomal_cut_line_T_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_unnomal_cut_line_I_異型沖)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_unnomal_cut_line_I_*,all")
            plate_line_unnomal_cut_line_I_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_unnomal_cut_line_M_異型沖)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_unnomal_cut_line_M_*,all")
            plate_line_unnomal_cut_line_M_number[i][tt] = selection1.Count
            selection1.Clear()
            # --------------------------------------------------------------------↑異型沖
            # ---------------------------------------------------------------------↓向下折彎_成形
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bend_down_forming_insert_surface_L_) 向下折彎_成形沖頭_左
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bend_down_forming_insert_surface_L_*,all")
            plate_line_bend_down_forming_insert_surface_L_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bend_down_forming_insert_surface_R_) 向下折彎_成形沖頭_右
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bend_down_forming_insert_surface_R_*,all")
            plate_line_bend_down_forming_insert_surface_R_number[i][tt] = selection1.Count
            selection1.Clear()
            # ----------------------------------------------------------------------↑向下折彎_成形
            # ----------------------------------------------------------------↓向上折彎shoulder     2016-12-21
            # -----------------------------↓(搜尋每塊模板以及每個工程的_shoulder_bending_surface_up_) 向上折彎靠肩沖頭
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shoulder_bending_surface_up_*,all")
            plate_line_shoulder_bendin_punch_number[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bend_up_forming_punch_surface_down_) 向下折彎靠肩模穴
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shoulder_bending_surface_down_*,all")
            plate_line_shoulder_bendin_cavity_number[i][tt] = selection1.Count
            if plate_line_shoulder_bendin_cavity_number[i][tt] == 4:
                # =============環境設置===============
                (ElementProduct, ElementDocument, ElementBody, ElementHybridBody) = defs.environment_set("Data1",
                                                                                                         "PartBody",
                                                                                                         "die")
                # =============環境設置===============
                now_plate_line_number = i
                part1 = ElementDocument.Part
                g = now_plate_line_number
                data_length = [[0.0] * 5 for iiii in range(5)]
                center_data = [[0.0] * 5 for iiii in range(3)]
                # --------------------------------------------------------------------------------------------↓XY平面宣告
                originElements1 = part1.OriginElements
                hybridShapePlaneExplicit1 = originElements1.PlaneXY
                # --------------------------------------------------------------------------------------------↑XY平面宣告
                hybridShape2 = ElementHybridBody.HybridShapes.Item("lower_die_seat_line")
                # 下模座的線段
                (ElementReference1) = defs.OriginalPoint(hybridShape2, "bending_punch_" + str(i), ElementDocument,
                                                         ElementBody, ElementHybridBody)
                ElementPoint1 = ElementReference1
                (ElementSketch1) = defs.BuildSketch("position_sketch", hybridShapePlaneExplicit1, ElementDocument,
                                                    SketchPosition, ElementBody, ElementHybridBody)
                part1.InWorkObject = ElementSketch1  # 目前工作位置=草圖
                for ii in range(1, 5):
                    hybridShape1 = ElementHybridBody.HybridShapes.Item(
                        "plate_line_" + str(g) + "_op" + str(op_number) + "_shoulder_bending_surface_up_" + str(i))
                    # 下模座的線段
                    (ElementReference1) = defs.ExtremumPoint("X_min", 0, 0, 1, hybridShape1, ElementDocument,
                                                             ElementBody, ElementHybridBody)
                    ElementPoint2 = ElementReference1
                    part1.Update()
                    (data_length[1][ii]) = defs.SketchBuildCallout(ElementSketch1, "free", "Callout",
                                                                   data_length[1][ii], ElementDocument, ElementPoint1,
                                                                   ElementPoint2)
                    (ElementPoint2) = defs.ExtremumPoint("X_max", 0, 0, 1, hybridShape1, ElementDocument, ElementBody,
                                                         ElementHybridBody)
                    part1.Update()
                    (data_length[2][ii]) = defs.SketchBuildCallout(ElementSketch1, "free", "Callout",
                                                                   data_length[2][ii],
                                                                   ElementDocument,
                                                                   ElementPoint1, ElementPoint2)
                    (ElementPoint2) = defs.ExtremumPoint("Y_min", 0, 0, 1, hybridShape1, ElementDocument, ElementBody,
                                                         ElementHybridBody)
                    part1.Update()
                    (data_length[3][ii]) = defs.SketchBuildCallout(ElementSketch1, "free", "Callout",
                                                                   data_length[3][ii],
                                                                   ElementDocument,
                                                                   ElementPoint1, ElementPoint2)
                    (ElementPoint2) = defs.ExtremumPoint("Y_max", 0, 0, 1, hybridShape1, ElementDocument, ElementBody,
                                                         ElementHybridBody)
                    part1.Update()
                    (data_length[4][ii]) = defs.SketchBuildCallout(ElementSketch1, "free", "Callout",
                                                                   data_length[4][ii],
                                                                   ElementDocument,
                                                                   ElementPoint1, ElementPoint2)
                    center_data[1][ii] = (data_length[1][ii] + data_length[2][ii]) / 2
                    center_data[2][ii] = (data_length[3][ii] + data_length[4][ii]) / 2
                # --------------------------------------------------------------------Y分組 shoulder_bending_grouping_parameter(Y, X)(Y分組1=上2 = 下, X位置排序)
                for ii in range(1, 5):
                    if StripWidth / 2 > center_data[2][ii]:
                        if shoulder_bending_grouping_parameter[1][1] == 0:
                            shoulder_bending_grouping_parameter[1][1] = i
                        else:
                            shoulder_bending_grouping_parameter[1][2] = i
                    if StripWidth / 2 < center_data[2][ii]:
                        if shoulder_bending_grouping_parameter[2][1] == 0:
                            shoulder_bending_grouping_parameter[2][1] = i
                        else:
                            shoulder_bending_grouping_parameter[2][2] = i
                # --------------------------------------------------------------------Y分組
                # ----------------------------------X排序shoulder_bending_grouping_parameter(i, 2)表示第幾個
                for ii in range(1, 3):
                    if center_data[1][shoulder_bending_grouping_parameter[ii][1]] > center_data[1][
                        shoulder_bending_grouping_parameter[ii][2]]:
                        SSS = shoulder_bending_grouping_parameter[ii][1]
                        shoulder_bending_grouping_parameter[ii][1] = shoulder_bending_grouping_parameter[ii][2]
                        shoulder_bending_grouping_parameter[ii][2] = SSS
                    # --------------------------------------------------------------------X排序
                    part1.Update()
            selection1.Clear()
            # -----------------------------------------------------↑向上折彎shoulder     2016-12-2
            # ------------------------------------------------------------------------------------------------------------------------↓  2016-8-16靠肩衝頭
            # -----------------------------↓(搜尋每塊模板以及每個工程的_shoulder_up_point_*_A 上靠肩點A)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shoulder_up_point_*_A,all")
            plate_line_shoulder_up_point_A[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_shoulder_up_point_*_B 上靠肩點B)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shoulder_up_point_*_B,all")
            plate_line_shoulder_up_point_B[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_shoulder_down_point_*_A 下靠肩點A)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shoulder_down_point_*_A,all")
            plate_line_shoulder_down_point_A[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_shoulder_down_point_*_B下靠肩點B)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_shoulder_down_point_*_B,all")
            plate_line_shoulder_down_point_B[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_left_shoulder_emboss_up_point_*_A 左上靠肩點A)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_left_shoulder_emboss_up_point_*_A,all")
            plate_line_left_shoulder_emboss_up_point_A[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_left_shoulder_emboss_up_point_*_B 左上靠肩點B)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_left_shoulder_emboss_up_point_*_B,all")
            plate_line_left_shoulder_emboss_up_point_B[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_left_shoulder_emboss_down_point_*_A 左下靠肩點A)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_left_shoulder_emboss_down_point_*_A,all")
            plate_line_left_shoulder_emboss_down_point_A[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_left_shoulder_emboss_down_point_*_B 左下靠肩點B)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_left_shoulder_emboss_down_point_*_B,all")
            plate_line_left_shoulder_emboss_down_point_B[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_right_shoulder_emboss_up_point_*_A 右上靠肩點A)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_right_shoulder_emboss_up_point_*_A,all")
            plate_line_right_shoulder_emboss_up_point_A[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_right_shoulder_emboss_up_point_*_B 右上靠肩點B)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_right_shoulder_emboss_up_point_*_B,all")
            plate_line_right_shoulder_emboss_up_point_B[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_right_shoulder_emboss_down_point_*_A 右下靠肩點A)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_right_shoulder_emboss_down_point_*_A,all")
            plate_line_right_shoulder_emboss_down_point_A[i][tt] = selection1.Count
            selection1.Clear()
            # -----------------------------↓(搜尋每塊模板以及每個工程的_right_shoulder_emboss_down_point_*_B 右下靠肩點B)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_right_shoulder_emboss_down_point_*_B,all")
            plate_line_right_shoulder_emboss_down_point_B[i][tt] = selection1.Count
            selection1.Clear()
            # -------------------------------------------------------↑   2016-8-16靠肩衝頭
            # -------------------------------------------------------↓  2016-12-13 整形工站
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bending_punch_surface_ 彎折沖頭)
            selection1.Clear()
            selection1.Search("Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bending_punch_surface_*,all")
            plate_line_bending_punch_surface[i][tt] = selection1.Count
            # -----------------------------↓(搜尋每塊模板以及每個工程的_bending_cavity_surface_ 彎折模穴)
            selection1.Clear()
            selection1.Search(
                "Name=plate_line_" + str(i) + "_op" + str(op_number) + "_bending_cavity_surface_*,all")
            plate_line_bending_cavity_surface[i][tt] = selection1.Count
            # ------------------------------------------------------↑  2016-12-13 整形工站
            selection1.Clear()  # 第2迴圈_結束
        # -----------------------------↓(搜尋每塊模板以及每個工程的_pilot_punch_)導引沖
        selection1.Clear()
        selection1.Search("Name=plate_line_" + str(i) + "_pilot_punch_*,all")
        plate_line_pilot_punch_number[i] = selection1.Count
        selection1.Clear()
        # -----------------------------↓(搜尋每塊模板以及每個工程的_stripper_pin_)浮升銷
        selection1.Clear()
        selection1.Search("Name=plate_line_" + str(i) + "_stripper_pin_point_*,all")
        plate_line_stripper_pin_point_number[i] = selection1.Count
        selection1.Clear()
        # -----------------------------↓(搜尋每塊模板以及每個工程的_LIFTER_point_)???
        selection1.Clear()
        selection1.Search("Name=plate_line_" + str(i) + "_LIFTER_point_*,all")
        plate_line_LIFTER_point_number[i] = selection1.Count
        selection1.Clear()
        # -----------------------------↓(搜尋每塊模板以及每個工程的_limit_point_)限制點
        selection1.Clear()
        selection1.Search("Name=plate_line_" + str(i) + "_limiting_point_*,all")
        plate_line_limiting_point_number[i] = selection1.Count
        selection1.Clear()
        bb[i] = gvar.strip_parameter_list[26]  # 之後修改介面模板設定 看位置
        # *************************************刪掉
    # ----------------------------------------------------plate_Data
    Boundary = [None] * 30
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    # ----------------------------------------------------------建立向量參數
    hybridShapeFactory1 = part1.HybridShapeFactory
    X_min = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)  # X方向  <- = +  -> = -     (X,Y,Z)
    X_max = hybridShapeFactory1.AddNewDirectionByCoord(-1, 0, 0)  # Y方向  ↓ = +  ↑ = -     (X,Y,Z)
    Y_min = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)  # Y方向  ↓ = +  ↑ = -     (X,Y,Z)
    Y_max = hybridShapeFactory1.AddNewDirectionByCoord(0, -1, 0)  # Y方向  ↓ = +  ↑ = -     (X,Y,Z)
    zero = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 0)  # 無方向     (X,Y,Z)
    # ----------------------------------------------------------建立向量參數
    hybridBodies1 = part1.HybridBodies
    hybridBody1 = hybridBodies1.Item("die")
    hybridShapes1 = hybridBody1.HybridShapes
    hybridShape1 = hybridShapes1.Item("lower_die_seat_line")
    reference1 = part1.CreateReferenceFromObject(hybridShape1)
    # ----------------------------------------------------------建立極值點1
    hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, X_min, 0)
    hybridShapeExtremum1.Direction2 = Y_min
    hybridShapeExtremum1.ExtremumType2 = 0
    hybridBody1.AppendHybridShape(hybridShapeExtremum1)
    hybridShapeExtremum1.Name = "die_basis_point"
    # ----------------------------------------------------------建立極值點1
    hybridShape2 = hybridShapes1.Item("number_1_plate_line")
    reference2 = part1.CreateReferenceFromObject(hybridShape2)
    # ----------------------------------------------------------建立極值點2
    hybridShapeExtremum2 = hybridShapeFactory1.AddNewExtremum(reference2, X_min, 0)
    hybridShapeExtremum2.Direction2 = Y_min
    hybridShapeExtremum2.ExtremumType2 = 0
    hybridBody1.AppendHybridShape(hybridShapeExtremum2)
    hybridShapeExtremum2.Name = "measure_point"
    part1.Update()
    # ----------------------------------------------------------建立極值點2
    # ----------------------------------------------------------草圖建立
    sketches1 = hybridBody1.HybridSketches
    originElements1 = part1.OriginElements
    reference3 = originElements1.PlaneXY
    sketch1 = sketches1.Add(reference3)
    sketch1Variant = sketch1
    # sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)
    part1.InWorkObject = sketch1
    # -----------------------------------------------------------草圖建立
    # -----------------------------------------------------------草圖坐標軸建立
    factory2D1 = sketch1.OpenEdition()
    geometricElements1 = sketch1.GeometricElements
    axis2D1 = geometricElements1.Item("AbsoluteAxis")
    line2D1 = axis2D1.getItem("HDirection")
    line2D2 = axis2D1.getItem("VDirection")
    # ----------------------------------------------------------草圖坐標軸建立
    # ----------------------------------------------------------建立投影點
    reference4 = hybridShapes1.Item("die_basis_point")
    reference5 = hybridShapes1.Item("measure_point")
    geometricElements2 = factory2D1.CreateProjections(reference4)
    geometricElements3 = factory2D1.CreateProjections(reference5)
    geometry2D1 = geometricElements2.Item("Mark.1")
    geometry2D1.Construction = True
    geometry2D2 = geometricElements3.Item("Mark.1")
    geometry2D2.Construction = True
    # ----------------------------------------------------------建立投影點
    # ----------------------------------------------------------建立標註依據
    reference6 = part1.CreateReferenceFromObject(geometry2D1)
    reference7 = part1.CreateReferenceFromObject(geometry2D2)
    reference8 = part1.CreateReferenceFromObject(line2D1)  # 水平方向
    reference9 = part1.CreateReferenceFromObject(line2D2)  # 垂直方向
    # ----------------------------------------------------------建立標註依據
    # ----------------------------------------------------------建立標註
    constraints1 = sketch1.Constraints
    constraint1 = constraints1.AddTriEltCst(1, reference6, reference7, reference8)
    constraint1.mode = 1
    constraint1.Name = "H_measure"
    constraint2 = constraints1.AddTriEltCst(1, reference6, reference7, reference9)
    constraint2.mode = 1
    constraint2.Name = "V_measure"
    # ----------------------------------------------------------建立標註
    sketch1.CloseEdition()  # 離開草圖模式
    sketch1.Name = "plate_measure"
    part1.InWorkObject = hybridBody1
    part1.Update()
    for i in range(1, gvar.PlateLineNumber + 1):
        Boundary[i] = hybridShapes1.Item("number_" + str(i) + "_plate_line")
        for direct in range(1, 3):
            defs.ChangeExtremum(Boundary[i], direct)  # 直接宣告曲線
            if direct == 1:
                constraints1 = sketch1.Constraints
                constraint1 = constraints1.Item("H_measure")
                plate_X_min[i] = round(constraint1.dimension.Value, 2)
                constraint1 = constraints1.Item("V_measure")
                plate_Y_min[i] = round(constraint1.dimension.Value, 2)
            elif direct == 2:
                constraints1 = sketch1.Constraints
                constraint1 = constraints1.Item("H_measure")
                plate_X_max[i] = round(constraint1.dimension.Value, 2)
                constraint1 = constraints1.Item("V_measure")
                plate_Y_max[i] = round(constraint1.dimension.Value, 2)
        plate_length_origin_die[i] = round(plate_X_max[i] - plate_X_min[i])  # 計算此模板長
        plate_wide_origin_die[i] = round(plate_Y_max[i] - plate_Y_min[i])  # 計算此模板寬
    # ---------------------------------------------------plate_Data
    partDocument1.Close()
    gvar.StripDataList = [None, PlateLength, plate_line_demise_surface_up_number_surch, plate_line_half_cut_line_number,
                          plate_line_reinforcement_cut_line, plate_line_leveling_block_up_inbolt_surface_number,
                          plate_line_leveling_block_up_inbolt_demise_surface_number,
                          plate_line_leveling_block_up_outbolt_surface_number,
                          plate_line_leveling_block_demise_up_outbolt_demise_surface_number,
                          plate_line_leveling_block_up_outbolt_side_surface_number,
                          plate_line_leveling_block_up_outbolt_side_demise_surface_number,
                          plate_line_leveling_block_down_inbolt_surface_number,
                          plate_line_leveling_block_down_inbolt_demise_surface_number,
                          plate_line_leveling_block_down_outbolt_surface_number,
                          plate_line_leveling_block_down_outbolt_demise_surface_number,
                          plate_line_leveling_block_down_outbolt_side_surface_number,
                          plate_line_leveling_block_down_outbolt_side_demise_surface_number,
                          plate_line_leveling_block_up_demise_line_number,
                          plate_line_leveling_block_down_demise_line_number,
                          plate_line_bend_up_forming_punch_surface_number,
                          plate_line_bend_up_emboss_forming_insert_surface_number,
                          plate_line_emboss_forming_punch_left_surface_number,
                          plate_line_emboss_forming_punch_right_surface_number, plate_line_bend_up_punch_surface_number,
                          plate_line_bend_up_insert_surface_number, plate_line_bend_up_forming_insert_surface_L_number,
                          plate_line_bend_up_forming_insert_surface_R_number, plate_line_cut_punch_d_cutting_number,
                          plate_line_cut_punch_u_cutting_number, plate_line_right_quickly_remove_cut_line_number,
                          plate_line_left_quickly_remove_cut_line_number, plate_line_up_quickly_remove_cut_line_number,
                          plate_line_down_quickly_remove_cut_line_number,
                          plate_line_right_quickly_remove_bending_surface_number,
                          plate_line_left_quickly_remove_bending_surface_number,
                          plate_line_up_quickly_remove_bending_surface_number,
                          plate_line_down_quickly_remove_bending_surface_number, plate_line_A_punch_number,
                          plate_line_cut_line_number, plate_line_allotype_cut_line_number,
                          plate_line_forming_cavity_surface_number, plate_line_forming_punch_surface_number,
                          plate_line_shaping_cavity_surface_down_number, plate_line_shaping_cavity_surface_up_number,
                          plate_line_shaping_punch_surface_down_number, plate_line_shaping_punch_surface_up_number,
                          plate_line_Bending_cavity_number_up, plate_line_Bending_punch_number_up,
                          plate_line_Bending_cavity_floating_number_up, plate_line_Bending_cavity_number_down,
                          plate_line_Bending_punch_number_down, plate_line_Bending_punch_up_number,
                          plate_line_Bending_punch_down_number, plate_line_unnomal_cut_line_T_number,
                          plate_line_unnomal_cut_line_I_number, plate_line_unnomal_cut_line_M_number,
                          plate_line_bend_down_forming_insert_surface_L_number,
                          plate_line_bend_down_forming_insert_surface_R_number, plate_line_shoulder_bendin_punch_number,
                          plate_line_shoulder_bendin_cavity_number, shoulder_bending_grouping_parameter,
                          plate_line_shoulder_up_point_A, plate_line_shoulder_up_point_B,
                          plate_line_shoulder_down_point_A,
                          plate_line_shoulder_down_point_B, plate_line_left_shoulder_emboss_up_point_A,
                          plate_line_left_shoulder_emboss_up_point_B, plate_line_left_shoulder_emboss_down_point_A,
                          plate_line_left_shoulder_emboss_down_point_B, plate_line_right_shoulder_emboss_up_point_A,
                          plate_line_right_shoulder_emboss_up_point_B, plate_line_right_shoulder_emboss_down_point_A,
                          plate_line_right_shoulder_emboss_down_point_B, plate_line_bending_punch_surface,
                          plate_line_bending_cavity_surface, plate_line_pilot_punch_number,
                          plate_line_stripper_pin_point_number,
                          plate_line_LIFTER_point_number, plate_line_limiting_point_number, bb, plate_X_min,
                          plate_X_max,
                          plate_Y_min, plate_Y_max, plate_length_origin_die, plate_wide_origin_die]
    print(gvar.StripDataList[37][1][1])
    print(gvar.StripDataList[38][1][2])
    print(gvar.StripDataList[38][1][3])
    print(gvar.StripDataList[38][1][4])
    print(gvar.StripDataList[38][1][5])
    print(gvar.StripDataList[38][1][6])
    print(gvar.StripDataList[38][1][7])


def atest():
    catapp = win32.Dispatch('CATIA.Application')
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    Cut_Perimeter = [0.0] * 99
    Surface_Perimeter = [0.0] * 99
    ALL_Cut_Perimeter = [0.0] * 99
    ALL_Surface_Perimeter = [0.0] * 99
    # ==========↓查詢抗剪強度↓
    die_rule_file_name = "rule"
    Row_string_serch = gvar.strip_parameter_list[3]
    Column_string_serch = "SHEAR"
    excel_Sheet_name = "Material"
    (gvar.Shear_Strength) = defs.ExcelSearch('rule', 'Material', 'AIPS-1', 'SHEAR')
    # ==========↑查詢抗剪強度↑
    MeasureEdge_number = [0] * 99
    MeasureSurface_number = [0] * 99
    # ==========↓搜尋週長種類數量↓
    selection1 = partDocument1.Selection
    selection1.Clear()
    selection1.Search("Name=MeasureEdge.*,all")  # 剪切周長數量
    MeasureEdge_number[1] = selection1.Count
    selection1.Clear()
    selection1.Search("Name=MeasureSurface.*,all")  # 曲面周長數量
    MeasureSurface_number[1] = selection1.Count
    selection1.Clear()
    # ==========↑搜尋週長種類數量↑
    # ==========↓儲存各剪切週長↓
    if MeasureEdge_number[1] > 0:
        for i in range(1, MeasureEdge_number[1] + 1):
            parameters1 = part1.Parameters
            length1 = parameters1.Item("Part1\MeasureEdge." + str(i) + "\Length")
            Cut_Perimeter[i] = length1.Value
            part1.Update()
    # ==========↑儲存各剪切週長↑
    # ==========↓儲存各成形週長↓
    if MeasureSurface_number[1] > 0:
        for j in range(1, MeasureSurface_number[1] + 1):
            parameters2 = part1.Parameters
            length2 = parameters2.Item("Part1\MeasureSurface." + str(j) + "\Perimeter")
            Surface_Perimeter[j] = length2.Value
            part1.Update()
    # ==========↑儲存各成形週長↑
    # ==========↓計算剪切總週長↓
    a = 0
    for k in range(1, MeasureEdge_number[1] + 1):
        a = a + 1
        c = a - 1
        ALL_Cut_Perimeter[0] = 0
        ALL_Cut_Perimeter[a] = ALL_Cut_Perimeter[c] + Cut_Perimeter[k]
    # ==========↑計算剪切總週長↑
    # ==========↓計算成形總週長↓
    b = 0
    for M in range(1, MeasureSurface_number[1] + 1):
        b = b + 1
        c = b - 1
        ALL_Surface_Perimeter[0] = 0
        ALL_Surface_Perimeter[b] = ALL_Surface_Perimeter[c] + Surface_Perimeter[M]
    # ==========↑計算成形總週長↑
    # ======↓沖頭行程↓=======
    punch_travel = 0
    # ======↑沖頭行程↑======
    # ==========↓計算脫料力↓
    q = MeasureEdge_number[1]
    R = MeasureSurface_number[1]
    Cut_Punching_force = (ALL_Cut_Perimeter[q] * float(
        gvar.strip_parameter_list[1])) * gvar.Shear_Strength  # --------------沖裁力
    forming_Punching_force = (ALL_Surface_Perimeter[R] * float(
        gvar.strip_parameter_list[1])) * gvar.Shear_Strength  # --------------成形
    Stripping_force = (Cut_Punching_force + forming_Punching_force) * 0.08  # ----------脫料力
    # ==========↑計算脫料力↑
    partDocument1.Close()


def StripSystem():
    with open(gvar.strip_parameters_file_root) as csvFile:
        rows = csv.reader(csvFile)
        strip_parameter_list = list(tuple(rows)[0])
        gvar.strip_parameter_list = strip_parameter_list
    (R_Value) = StripAnalyze()
    (StripLength, StripWidth) = StripBuild(R_Value)
    # StripLength = 290
    # StripWidth = 39.4
    (SketchPosition) = DataBuild(StripLength, StripWidth)
    # SketchPosition = "Body"
    DataSetting(SketchPosition, StripWidth)
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
    atest()
# gvar.StripDataList[37][1][1] = 2
# gvar.StripDataList[38][1][2] = 4
# gvar.StripDataList[38][1][3] = 1
# gvar.StripDataList[38][1][4] = 1
# gvar.StripDataList[38][1][5] = 0
# gvar.StripDataList[38][1][6] = 4
# gvar.StripDataList[38][1][7] = 1

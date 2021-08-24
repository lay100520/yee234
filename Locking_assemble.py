import win32com.client as win32
import os
import global_var as gvar

def Lockingassemble(CB_data,BoltQuantity):  # 螺栓組立
    # bending_bolt()
    splint(CB_data,BoltQuantity)
    stop_plate(CB_data,BoltQuantity)
    lower_die(CB_data,BoltQuantity)
    # guide_plate()
    # emboss_forming_punch_left()
    # emboss_forming_punch_right()
    up_plate_Bolt()
    stop_plate_Bolt()
    # half_cut_punch_assemble()
    # Bend_shaping_form_bolt_assemble()
    # hide()
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()
    insert_assemble()  # 入子螺栓
    product1.Update()

def splint(CB_data,BoltQuantity):
    catapp = win32.Dispatch('CATIA.Application')
    plate_line_number =int( gvar.PlateLineNumber)
    for g in range(1, plate_line_number + 1):
        M = 0
        # =====================螺栓判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=*Bolt_" + str(CB_data[1][1]) + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================螺栓判斷(搜尋)===============================
        for i in range(1, BoltQuantity[1] + 1):
            M += 1
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + "Bolt_" + str(CB_data[1][1]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Product1/Splint_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Locking_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][1]) + "." + str(M) + "/!End_Point")
            try:
                constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)
            except:
                pass
            reference4 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Locking_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][1]) + "." + str(M) + "/!Start_Point")
            try:
                constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)
            except:
                pass

            product1.Update()


def stop_plate(CB_data,BoltQuantity):
    catapp = win32.Dispatch('CATIA.Application')
    plate_line_number = 1
    for g in range(1, plate_line_number + 1):
        M = 0
        # =====================螺栓判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=*Bolt_" + str(CB_data[1][2]) + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================螺栓判斷(搜尋)===============================
        for i in range(1, BoltQuantity[2] + 1):
            M += 1
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + "Bolt_" + str(CB_data[1][2]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Product1/Splint_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Locking_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][2]) + "." + str(M) + "/!End_Point")
            try:
                constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)
            except:
                pass
            reference4 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Locking_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][2]) + "." + str(M) + "/!Start_Point")
            try:
                constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)
            except:
                pass
            product1.Update()


def lower_die(CB_data,BoltQuantity):
    catapp = win32.Dispatch('CATIA.Application')
    plate_line_number = 1
    for g in range(1, plate_line_number + 1):
        M = 0
        # =====================螺栓判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=*Bolt_" + str(CB_data[1][3]) + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================螺栓判斷(搜尋)===============================
        for i in range(1, BoltQuantity[3] + 1):
            M += 1
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + "Bolt_" + str(CB_data[1][3]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Product1/Stop_plate_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Locking_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][3]) + "." + str(M) + "/!End_Point")
            try:
                constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)
            except:
                pass
            reference4 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Locking_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][3]) + "." + str(M) + "/!Start_Point")
            try:
                constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)
            except:
                pass
            product1.Update()


def guide_plate():
    catapp = win32.Dispatch('CATIA.Application')


def emboss_forming_punch_left():
    catapp = win32.Dispatch('CATIA.Application')


def emboss_forming_punch_right():
    catapp = win32.Dispatch('CATIA.Application')


def up_plate_Bolt():
    catapp = win32.Dispatch('CATIA.Application')
    # =====================數值設定=========================
    CB_length = str(46)
    CB_M = str(8)
    point_name1 = "Start_Point"
    point_name2 = "End_Point"
    # =====================數值設定=========================
    (CB_length) = lifter_guide_save_CB(CB_M, CB_length)  # 建檔副程式呼叫
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    # =====================搜尋現有的螺栓數量=========================
    partdoc = catapp.ActiveDocument
    selection1 = partdoc.Selection
    selection1.Clear()
    selection1.Search("Name=Bolt_CB_" + str(CB_M) + str(CB_length) + "_*")
    CB_Counter1 = selection1.Count
    selection1.Clear()
    CB_Counter1 += 1
    # =====================搜尋現有的螺栓數量=========================
    element_name1 = "Bolt_CB_" + str(CB_M) + "-" + str(CB_length)  # 數值設定
    now_plate_line_number = 2
    for g in range(1, now_plate_line_number + 1):
        up_pad_Bolt_Hole = [g] * 9
        up_pad_Bolt_Hole[g] = 0
        for i in range(1, int(up_pad_Bolt_Hole[g] / 2 + 1)):
            # =====================匯入檔案到組立=========================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + element_name1 + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1)
            # =====================匯入檔案到組立=========================
            for for_start in range(1, 2 + 1):
                assemble_name1 = "Product1/" + element_name1 + "." + CB_Counter1 + "/!" + point_name2
                assemble_name2 = str(
                    "Product1/up_plate_" + str(g) + ".1" + "/!up_pad_" + str(g) + "_Bolt_point_" + str(i * 2 - 1))
                assemble(assemble_name1, assemble_name2)
                assemble_name1 = "Product1/" + element_name1 + "." + CB_Counter1 + "/!" + point_name1
                assemble_name2 = str(
                    "Product1/up_plate_" + str(g) + ".1" + "/!up_pad_" + str(g) + "_Bolt_point_" + str(i * 2))
                assemble(assemble_name1, assemble_name2)
            CB_Counter1 += 1


def stop_plate_Bolt():
    catapp = win32.Dispatch('CATIA.Application')
    # =====================數值設定=========================
    CB_length = str(int(gvar.strip_parameter_list[20]) - 11 + 16)
    CB_M = str(8)
    point_name1 = "Start_Point"
    point_name2 = "End_Point"
    # =====================數值設定=========================
    (CB_length) = lifter_guide_save_CB(CB_M, CB_length)  # 建檔副程式呼叫
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    # =====================搜尋現有的螺栓數量=========================
    partdoc = catapp.ActiveDocument
    selection1 = partdoc.Selection
    selection1.Clear()
    selection1.Search("Name=Bolt_CB_" + str(CB_M) + str(CB_length) + "_*")
    CB_Counter1 = selection1.Count
    selection1.Clear()
    CB_Counter1 += 1
    # =====================搜尋現有的螺栓數量=========================
    element_name1 = "Bolt_CB_" + str(CB_M) + "-" + str(CB_length)  # 數值設定
    now_plate_line_number = 2
    for g in range(1, now_plate_line_number + 1):
        up_pad_Bolt_Hole = [g] * 9
        up_pad_Bolt_Hole[g] = 0
        for i in range(1, int(up_pad_Bolt_Hole[g] / 2 + 1)):
            # =====================匯入檔案到組立=========================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + element_name1 + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1)
            # =====================匯入檔案到組立=========================
            for for_start in range(1, 2):
                assemble_name1 = "Product1/" + element_name1 + "." + CB_Counter1 + "/!" + point_name2
                assemble_name2 = str(
                    "Product1/Stop_plate_" + str(g) + ".1" + "/!Stop_plate_" + str(g) + "_Bolt_point_" + str(
                        i * 2 - 1))
                assemble(assemble_name1, assemble_name2)
                assemble_name1 = "Product1/" + element_name1 + "." + CB_Counter1 + "/!" + point_name1
                assemble_name2 = str(
                    "Product1/Stop_plate_" + str(g) + ".1" + "/!Stop_plate_" + str(g) + "_Bolt_point_" + str(i * 2))
                assemble(assemble_name1, assemble_name2)
            CB_Counter1 += 1


def half_cut_punch_assemble():
    catapp = win32.Dispatch('CATIA.Application')


def Bend_shaping_form_bolt_assemble():
    catapp = win32.Dispatch('CATIA.Application')


def hide():
    catapp = win32.Dispatch('CATIA.Application')


def insert_assemble():  # 入子螺栓組立
    catapp = win32.Dispatch('CATIA.Application')
    lower_die_cavity_plate_height = 40  # 測試用數值
    # =====================數值設定=========================
    CB_length = lower_die_cavity_plate_height - 13
    CB_length += 16
    CB_M = 8
    # =====================數值設定=========================
    (CB_length) = lifter_guide_save_CB(CB_M, CB_length)  # 建檔副程式呼叫
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    # =====================搜尋現有的螺栓數量=========================
    partdoc = catapp.ActiveDocument
    selection1 = partdoc.Selection
    selection1.Clear()
    selection1.Search("Name=Bolt_CB_" + str(CB_M) + str(CB_length) + "_*")
    CB_Counter1 = selection1.Count
    selection1.Clear()
    CB_Counter1 += 1
    # =====================搜尋現有的螺栓數量=========================
    now_plate_line_number = 1
    for g in range(1, now_plate_line_number + 1):
        pad_Bolt_Hole = [g] * 6
        pad_Bolt_Hole[g] = 7
        for i in range(1, pad_Bolt_Hole[g] + 1):
            # =====================匯入檔案到組立=========================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + "Bolt_CB_" + str(CB_M) + "-" + str(CB_length) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            # =====================匯入檔案到組立=========================
            # =====================組立拘束宣告=========================
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/lower_pad_" + str(g) + ".1/!Product1/lower_pad_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            try:
                reference2 = product1.CreateReferenceFromName(
                    "Product1/Bolt_CB_" + str(CB_M) + "-" + str(CB_length) + "." + str(CB_Counter1) + "/!Start_Point")
                reference3 = product1.CreateReferenceFromName(
                    "Product1/lower_pad_" + str(g) + ".1" + "/!pad_" + str(g) + "_Bolt_point_" + str(i))
                constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            except:
                pass
            length1 = constraint2.dimension
            length1.Value = 0
            CB_Counter1 += 1
    # =====================組立拘束宣告=========================


def lifter_guide_save_CB(M, length):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    plate_bolt = 0  # 初始化值
    partDocument1 = documents1.Open(gvar.standard_path + "\\Bolt\\CB_" + str(M) + ".CATPart")  # 開啟檔案
    product1 = partDocument1.getItem("CB_" + str(M))
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    # =====================螺栓長度=========================
    strParam1 = parameters1.Item("CB_M_L")
    iSize = strParam1.GetEnumerateValuesSize()  # STRING 裡面選擇數量
    myArray = [iSize - 1] * 31
    myArray[iSize - 1] = "8-200"
    strParam1.GetEnumerateValues(myArray)  # 抓取 STRING 的數值放入
    # =====================找尋適合的螺栓=========================
    plate_bolt = ""
    while length != 0 and plate_bolt == "":
        length = int(length)
        length -= 1
        plate_bolt_test_name_1 = str(M) + "-" + str(length)
        for Array_count in range(1, iSize):
            if myArray[Array_count] == plate_bolt_test_name_1:
                plate_bolt = plate_bolt_test_name_1
    # =====================找尋適合的螺栓=========================
    CB_name = "CB_" + plate_bolt_test_name_1
    # =====================螺栓長度=========================
    # =====================找尋現有螺栓的尺寸=========================
    file_name = os.listdir(gvar.save_path)
    if "Bolt_" + CB_name + ".CATPatr" not in file_name:
        # =====================找尋現有螺栓的尺寸=========================
        strParam1 = parameters1.Item("CB_M_L")  # 參數宣告
        strParam1.Value = plate_bolt  # 變更
        product1.PartNumber = "Bolt_" + CB_name  # 改part名字(非檔名)
        part1.Update()
        partDocument1.SaveAs(gvar.save_path + "Bolt_" + CB_name + ".CATPart")
    part1.Update()
    partDocument1.Close()
    # =====================參數宣告及變更=========================
    return length  # 回傳數值


def assemble(assemble_name1, assemble_name2):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    # =====================組立拘束宣告=========================
    constraints1 = product1.Connections("CATIAConstraints")
    reference1 = product1.CreateReferenceFromName(assemble_name1)
    reference2 = product1.CreateReferenceFromName(assemble_name2)
    constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
    length1 = constraint1.dimension
    length1.Value = 0
    # =====================組立拘束宣告=========================
    product1.Update()

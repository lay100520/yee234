import win32com.client as win32
import global_var as gvar

def out_Guide_posts_locking_assemble(outer_Guiding_data):  # 外導柱螺栓
    (M) = out_Guide_posts_down_locking_assemble(outer_Guiding_data)  # 下模座螺栓
    # hide_1()
    (N) = out_Guide_posts_down_pin_assemble(outer_Guiding_data)  # 下模座合銷
    # hide_2()
    out_Guide_posts_up_locking_assemble(M,outer_Guiding_data)  # 上模座螺栓
    # # hide_3()
    out_Guide_posts_up_pin_assemble(N,outer_Guiding_data)  # 上模座合銷
    # # hide_4()
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def out_Guide_posts_down_locking_assemble(outer_Guiding_data):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    # if outer_Guiding_data(1, 1) == 20 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_6"
    # elif outer_Guiding_data(1, 1) == 25 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_8"
    # elif outer_Guiding_data(1, 1) == 32 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 38 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 50 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_12"
    # elif outer_Guiding_data(1, 1) == 20 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_8"
    # elif outer_Guiding_data(1, 1) == 25 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_8"
    # elif outer_Guiding_data(1, 1) == 32 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 38 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 50 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_12"
    M = 0
    # =====================螺栓判斷(搜尋)===============================
    selection1 = document.Selection
    selection1.Clear()
    selection1.Search("Name=" + str(outer_Guiding_data[4][1]) + "*")
    N = selection1.Count
    selection1.Clear()
    # =====================螺栓判斷(搜尋)===============================
    for g in range(1, 4 + 1):
        for i in range(1, 4 + 1):
            M += 1
            # ================匯入檔案================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + str(outer_Guiding_data[4][1]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            # ================匯入檔案================
            constraints1 = product1.Connections("CATIAConstraints")
            # ================進行拘束================
            reference1 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_down." + str(
                    g) + "/!Product1/" + str(outer_Guiding_data[3][1]) + "_down." + str(g) + "/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_down." + str(
                    g) + "/!Locking_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[4][1]) + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)
            reference4 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_down." + str(
                    g) + "/!Locking_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[4][1]) + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)
            # ================進行拘束================
    return M


def out_Guide_posts_down_pin_assemble(outer_Guiding_data):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    # if outer_Guiding_data(1, 1) == 20 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_6"
    # elif outer_Guiding_data(1, 1) == 25 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_8"
    # elif outer_Guiding_data(1, 1) == 32 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 38 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 50 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_12"
    # elif outer_Guiding_data(1, 1) == 20 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_8"
    # elif outer_Guiding_data(1, 1) == 25 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_8"
    # elif outer_Guiding_data(1, 1) == 32 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 38 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 50 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_12"
    N = 0
    # =====================pin判斷(搜尋)===============================
    selection1 = document.Selection
    selection1.Clear()
    selection1.Search("Name=" + str(outer_Guiding_data[5][1]))
    N = selection1.Count
    selection1.Clear()
    # =====================pin判斷(搜尋)===============================
    for g in range(1, 4 + 1):
        for i in range(1, 2 + 1):
            N += 1
            # ================匯入檔案================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + str(outer_Guiding_data[5][1]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            # ================匯入檔案================
            constraints1 = product1.Connections("CATIAConstraints")
            # ================進行拘束================
            reference1 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_down." + str(
                    g) + "/!Product1" + str(outer_Guiding_data[3][1]) + "_" + str(
                    outer_Guiding_data[1][1]) + "_down." + str(g) + "/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_down." + str(
                    g) + "/!Pin_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[5][1]) + "." + str(N) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)
            reference4 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_down." + str(
                    g) + "/!Pin_dir_point_" + str(i))
            # reference5 = product1.CreateReferenceFromName(
            #     "Product1/" + str(outer_Guiding_data[5][1]) + "." + str(N) + "/!End_Point")
            # constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)

            # WordCount_PinLength = len(outer_Guiding_data[5][1])
            # for j in range(0, WordCount_PinLength):
            #     word = outer_Guiding_data51[j]  # 提取Pin_data[2][1]中的值
            #     if word == "1":
            #         length2 = constraint3.dimension
            #         if WordCount_PinLength < 14:
            #             k = 2
            #         else:
            #             k = 3
            #         if int(int(outer_Guiding_data51[j + 3:20]) - 30) < 0:
            #             length2.Value = int((int(outer_Guiding_data51[j + k:10]) - 30) * -1)
            #         else:
            #             length2.Value = int(int(outer_Guiding_data51[j + k:10]) - 30)
            # ================進行拘束================
    return N


def out_Guide_posts_up_locking_assemble(M,outer_Guiding_data):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    # if outer_Guiding_data(1, 1) == 32 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 38 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 50 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_12"
    # elif outer_Guiding_data(1, 1) == 20 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_8"
    # elif outer_Guiding_data(1, 1) == 25 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "CB_8"
    # elif outer_Guiding_data(1, 1) ==32 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_810"
    # elif outer_Guiding_data(1, 1) == 38 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_10"
    # elif outer_Guiding_data(1, 1) == 50 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "CB_12"
    # =====================螺栓判斷(搜尋)===============================
    # selection1 = document.Selection
    # selection1.Clear()
    # selection1.Search("Name=" + str(outer_Guiding_data[4][2]) + "*")
    # M = selection1.Count
    # selection1.Clear()
    # =====================螺栓判斷(搜尋)===============================
    for g in range(1, 4 + 1):
        for i in range(1, 4 + 1):
            M += 1
            # ================匯入檔案================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + str(outer_Guiding_data[4][2]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            # ================匯入檔案================
            constraints1 = product1.Connections("CATIAConstraints")
            # ================進行拘束================
            reference1 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_up." + str(
                    g) + "/!Product1" + str(outer_Guiding_data[3][1]) + "_" + str(
                    outer_Guiding_data[1][1]) + "_up." + str(g) + "/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_up." + str(
                    g) + "/!Locking_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[4][2]) + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)
            reference4 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_up." + str(
                    g) + "/!Locking_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[4][2]) + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)
            # ================進行拘束================
            # product1.Update()


def out_Guide_posts_up_pin_assemble(N,outer_Guiding_data):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    # if outer_Guiding_data(1, 1) == 20 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "MSTM_6"
    # elif outer_Guiding_data(1, 1) == 25 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "MSTM_8"
    # elif outer_Guiding_data(1, 1) == 32 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "MSTM_8"
    # elif outer_Guiding_data(1, 1) == 38 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "MSTM_8"
    # elif outer_Guiding_data(1, 1) == 50 and outer_Guiding_data(3, 1) == "MYJP":
    #     a = "MSTM_10"
    # elif outer_Guiding_data(1, 1) ==20 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "MSTM_8"
    # elif outer_Guiding_data(1, 1) == 25 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "MSTM_8"
    # elif outer_Guiding_data(1, 1) == 32 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "MSTM_8"
    # elif outer_Guiding_data(1, 1) == 38 and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "MSTM_10"
    # elif outer_Guiding_data(1, 1) ==  and outer_Guiding_data(3, 1) == "MYKP":
    #     a = "MSTM_10"
    # =====================pin判斷(搜尋)===============================
    # selection1 = document.Selection
    # selection1.Clear()
    # selection1.Search("Name=*" + str(outer_Guiding_data[4][2]) + ".*")
    # N = selection1.Count
    # selection1.Clear()
    # =====================pin判斷(搜尋)===============================
    for g in range(1, 4 + 1):
        for i in range(1, 2 + 1):
            N += 1
            # ================匯入檔案================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + str(outer_Guiding_data[5][2]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            # ================匯入檔案================
            constraints1 = product1.Connections("CATIAConstraints")
            # ================進行拘束================
            reference1 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_up." + str(
                    g) + "/!Product1" + str(outer_Guiding_data[3][1]) + "_" + str(
                    outer_Guiding_data[1][1]) + "_up." + str(g) + "/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_up." + str(
                    g) + "/!Pin_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + str(outer_Guiding_data[5][2]) + "." + str(N) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)
            # reference4 = product1.CreateReferenceFromName(
            #     "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_up." + str(
            #         g) + "/!Pin_dir_point_" + str(i))
            # reference5 = product1.CreateReferenceFromName(
            #     "Product1/" + str(outer_Guiding_data[5][2]) + "." + str(N) + "/!End_Point")
            # constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)
            # WordCount_PinLength = len(outer_Guiding_data[5][2])
            # for j in range(0, WordCount_PinLength):
            #     word = outer_Guiding_data51[j]  # 提取Pin_data[2][1]中的值
            #     if word == "1":
            #         length2 = constraint3.dimension
            #         if WordCount_PinLength < 14:
            #             k = 2
            #         else:
            #             k = 3
            #         if int(int(outer_Guiding_data51[j + 3:20]) - 30) < 0:
            #             length2.Value = int((int(outer_Guiding_data51[j + k:10]) - 30) * -1)
            #         else:
            #             length2.Value = int(int(outer_Guiding_data51[j + k:10]) - 30)
            # ================進行拘束================
            # product1.Update()


def hide1():
    catapp = win32.Dispatch('CATIA.Application')


def hide2():
    catapp = win32.Dispatch('CATIA.Application')


def hide3():
    catapp = win32.Dispatch('CATIA.Application')


def hide4():
    catapp = win32.Dispatch('CATIA.Application')

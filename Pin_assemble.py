import win32com.client as win32
import global_var as gvar

def Pin_assemble(Pin_data,PinQuantity):  # 和銷
    splint(Pin_data,PinQuantity)
    stop_plate(Pin_data,PinQuantity)
    lower_die(Pin_data,PinQuantity)
    # guide_plate()
    # hide()
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def splint(Pin_data,PinQuantity):
    catapp = win32.Dispatch('CATIA.Application')
    b = "Pin_" + str(Pin_data[2][1])  # pin長度
    plate_line_number = 1
    first_number_position = int()
    for g in range(1, plate_line_number + 1):
        M = 0
        # =====================pin判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=" + b + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================pin判斷(搜尋)===============================
        for i in range(1, PinQuantity[1] + 1):
            M += 1
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + b + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Product1/Splint_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Pin_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0
            reference4 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Pin_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)
            WordCount_PinLength = len(Pin_data[2][1])
            for j in range(0, WordCount_PinLength):
                word = Pin_data[2][1][j]  # 提取Pin_data[2][1]中的值
                try:
                    a=int(word)
                    first_number_position = j
                    break
                except:
                    pass
            for j in range(WordCount_PinLength,0,-1):
                word = Pin_data[2][1][j]  # 提取Pin_data[2][1]中的值
                try:
                    a=int(word)
                    last_number_position = j
                    break
                except:
                    pass
            length2 = constraint3.dimension
            length2.Value = int(Pin_data[2][1][first_number_position + 3:last_number_position]) - float(gvar.strip_parameter_list[14]) * 0.5
            product1.Update()


def stop_plate(Pin_data,PinQuantity):
    catapp = win32.Dispatch('CATIA.Application')
    b = "Stripper_pin_" + str(Pin_data[2][2])  # 帶頭合銷_型號_直徑-長度
    plate_line_number = 1
    for g in range(1, plate_line_number + 1):
        M = 0
        # =====================pin判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=" + b + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================pin判斷(搜尋)===============================
        for i in range(1, PinQuantity[2] + 1):
            M += 1
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + b + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Product1/Stop_plate_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Pin_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 5
            reference4 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Pin_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)
            length2 = constraint3.dimension
            length2.Value = 0
            product1.Update()


def lower_die(Pin_data,PinQuantity):
    catapp = win32.Dispatch('CATIA.Application')
    plate_line_number = 1
    # a = Form19.Combo2 + Pin_data[1, 3]
    b = "Pin_" + str(Pin_data[2][3])
    for g in range(1, plate_line_number + 1):
        M = 0
        # =====================pin判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=" + b + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================pin判斷(搜尋)===============================
        for i in range(1, PinQuantity[3] + 1):
            M += 1
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + str(b) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Product1/lower_die_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Pin_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0
            reference4 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Pin_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)
            lower_die_cavity_plate_height = 40
            length2 = constraint3.dimension
            length2.Value = round(float(Pin_data[2][3]) - float(gvar.strip_parameter_list[26]) * 0.5,3)
            product1.Update()


def guide_plate():
    catapp = win32.Dispatch('CATIA.Application')


def hide():
    catapp = win32.Dispatch('CATIA.Application')

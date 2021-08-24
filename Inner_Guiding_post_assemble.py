import win32com.client as win32
import global_var as gvar

def Inner_Guiding_post_assemble(Inner_Guiding_data, InnerGuidingQuantity):  # 內導柱/套
    splint(Inner_Guiding_data, InnerGuidingQuantity)
    Stripper(Inner_Guiding_data, InnerGuidingQuantity)
    lower_die(Inner_Guiding_data, InnerGuidingQuantity)
    # hide()
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def splint(Inner_Guiding_data, InnerGuidingQuantity):
    catapp = win32.Dispatch('CATIA.Application')
    M = 0
    plate_line_number = gvar.PlateLineNumber
    for g in range(1, plate_line_number + 1):
        for i in range(1, InnerGuidingQuantity[1] + 1):
            M += 1
            Inner_Guiding_Post_Material = "SGPH"
            Inner_Guiding_Post_Diameter = 20
            a = Inner_Guiding_Post_Material + "_" + str(Inner_Guiding_Post_Diameter)
            b = str(Inner_Guiding_data[2][1])
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + a + "-" + b + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Product1/Splint_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Inner_Guiding_Post_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + a + "." + str(M) + "/!Inner_Guiding_Post_dir_point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0.3
            reference4 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Inner_Guiding_Post_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + a + "." + str(M) + "/!Inner_Guiding_Post_point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)
            length1 = constraint3.dimension
            length1.Value = 0


def Stripper(Inner_Guiding_data, InnerGuidingQuantity):
    catapp = win32.Dispatch('CATIA.Application')
    M = 0
    plate_line_number = gvar.PlateLineNumber
    for g in range(1, plate_line_number + 1):
        for i in range(1, InnerGuidingQuantity[1] + 1):
            M += 1
            Under_Inner_Guiding_Post_Material = "SGFZ"
            a = Under_Inner_Guiding_Post_Material + "_" + str(Inner_Guiding_data[1][1])
            b = 20  # Inner_Guiding_Post_Bush_up_data[2][1]
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + str(a) + "-" + str(b) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!Product1/Stripper_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!Inner_Guiding_Post_Bush_up_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + str(a) + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0.3
            reference4 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!Inner_Guiding_Post_Bush_up_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + str(a) + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)
            length1 = constraint3.dimension
            length1.Value = 0


def lower_die(Inner_Guiding_data, InnerGuidingQuantity):
    catapp = win32.Dispatch('CATIA.Application')
    M = InnerGuidingQuantity[1]
    plate_line_number = gvar.PlateLineNumber
    for g in range(1, plate_line_number + 1):
        for i in range(1, InnerGuidingQuantity[1] + 1):
            M += 1
            a = Inner_Guiding_data[1][1]
            b = 20  # Inner_Guiding_Post_Bush_up_data[2, 2]
            Under_Inner_Guiding_Post_Material = "SGFZ"
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = \
                gvar.save_path + Under_Inner_Guiding_Post_Material + "_" + str(a) + "-" + str(b) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Product1//Stop_plate_" + str(g) + ".1//")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Inner_Guiding_Post_Bush_down_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + Under_Inner_Guiding_Post_Material + "_" + str(a) + "." + str(M) + "/!Start_Point")
            reference4 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Inner_Guiding_Post_Bush_down_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + Under_Inner_Guiding_Post_Material + "_" + str(a) + "." + str(M) + "/!End_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference4, reference5)
            length1 = constraint2.dimension
            length1.Value = 0.3
            constraint3 = constraints1.AddBiEltCst(1, reference2, reference3)
            length2 = constraint3.dimension
            length2.Value = 0


def hide():
    catapp = win32.Dispatch('CATIA.Application')

import win32com.client as win32
import global_var as gvar

def SBT_assemble(SBT_data):  # 螺栓組立
    # stop_plate()
    Stripper(SBT_data)
    # hide()

    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def stop_plate():
    catapp = win32.Dispatch('CATIA.Application')


def Stripper(SBT_data):
    catapp = win32.Dispatch('CATIA.Application')
    M = 0
    plate_line_number = gvar.PlateLineNumber
    for g in range(1, plate_line_number + 1):
        # =====================螺栓判斷(搜尋)===============================
        # partdoc = catapp.ActiveDocument
        # selection1 = partdoc.Selection
        # selection1.Clear()
        # selection1.Search("Name=*" + Pin_Hole_ + "_*")
        # M = selection1.Count
        # selection1.Clear()
        # =====================螺栓判斷(搜尋)===============================
        for i in range(1, 2 + 1):
            M += 1
            a = SBT_data[1][1]
            b = SBT_data[2][1]
            # document = catapp.Documents
            # partDocument1 = document.Open(
            #     "C:\\Users\\PDAL\\Desktop\\auto\\Standard_Assembly\\MSTP " + str(a) + "-" + str(b) + ".CATPart")
            # part1 = partDocument1.part
            # length1 = part1.Parameters.Item("T")
            # g = now_plate_line_number
            # length1.Value = Bolt_data[2][1]
            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = gvar.save_path + str(SBT_data[7][1]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!Product1/Stripper_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)
            reference2 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!SBT_dir1_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + str(SBT_data[7][1]) + "." + str(M) + "/!SBT_dir_point")
            constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)
            reference4 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!SBT_dir2_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + str(SBT_data[7][1]) + "." + str(M) + "/!SBT_point")
            constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)


def hide():
    catapp = win32.Dispatch('CATIA.Application')



import win32com.client as win32
import time
import global_var as gvar

plate_line_number = gvar.PlateLineNumber
total_op_number = int(gvar.strip_parameter_list[2])
Working_parameter = 4.98
parameter_name = [0] * 3
plane_limit = [''] * 3


def Plate_Pilot_Punch_Hole():  # 引導沖孔挖孔
    for now_plate_line_number in range(1, plate_line_number + 1):
        g = now_plate_line_number
        for now_op_number in range(1, total_op_number + 1):
            n = now_op_number
            op_number = 10 * n
            if int(gvar.StripDataList[37][g][n]) > 0:
                parameter_name[1] = 0  # 加工說明
                # ==========================下模板==========================
                product_name = "lower_die_" + str(g)
                item_belong = "Body.2"
                plane_limit[1] = "down_die_plate_up_plane"
                plane_limit[2] = "down_die_plate_down_plane"
                pilot_punch_hole(product_name, item_belong, "lower_die",now_plate_line_number,now_op_number)
                lower_die_machining_instructions_Pilot_Punch_Hole = parameter_name[1]
                # ==========================下模板==========================
                parameter_name[1] = 0  # 加工說明
                # ==========================下墊板==========================
                product_name = "lower_pad_" + str(g)
                item_belong = "Body.2"
                plane_limit[1] = "up_plane"
                plane_limit[2] = "down_plane"
                pilot_punch_hole(product_name, item_belong, "lower_pad",now_plate_line_number,now_op_number)
                lower_pad_machining_instructions_Pilot_Punch_Hole = parameter_name[1]
                # ==========================下墊板==========================
                parameter_name[1] = 0  # 加工說明
                # ==========================下模座==========================
                product_name = "lower_die_set"
                item_belong = "PartBody"
                plane_limit[1] = "up_plane"
                plane_limit[2] = "down_plane"
                pilot_punch_hole(product_name, item_belong, "lower_die_set",now_plate_line_number,now_op_number)
                lower_die_set_machining_instructions_Pilot_Punch_Hole = parameter_name[1]
                # ==========================下模座==========================
                parameter_name[1] = 0  # 加工說明
                # ==========================脫料板==========================
                product_name = "Stripper_" + str(g)
                item_belong = "Body.2"
                plane_limit[1] = "down_die_plate_down_plane"
                plane_limit[2] = "down_die_plate_up_plane"
                pilot_punch_hole(product_name, item_belong, "stripper_plate",now_plate_line_number,now_op_number)
                Stripper_machining_instructions_Pilot_Punch = parameter_name[1]
                # ==========================脫料板==========================


def pilot_punch_hole(product_name, item_belong, mode,now_plate_line_number,now_op_number):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    partDocument1 = document.Open(gvar.save_path + product_name + ".CATPart")
    # ==========================新增body==========================
    part1 = partDocument1.Part
    bodies1 = part1.Bodies
    body1 = bodies1.Add()
    body1.Name = "Pilot_Body"
    part1.Update()
    # ==========================新增body==========================
    g = now_plate_line_number
    n = now_op_number
    op_number = 10 * n
    for i in range(1, int(gvar.StripDataList[37][g][n]) + 1):
        parameter_name[1] = parameter_name[1] + (total_op_number - now_op_number)
        # ==========================投影外型線==========================
        hybridShapeFactory1 = part1.HybridShapeFactory
        parameters1 = part1.Parameters
        hybridShapeCurveExplicit1 = parameters1.Item(
            "plate_line_" + str(g) + "_op" + str(op_number) + "_A_punch_" + str(i))
        reference1 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
        body2 = bodies1.Item(item_belong)
        hybridShapes1 = body2.HybridShapes
        hybridShapePlaneOffset1 = hybridShapes1.Item(plane_limit[1])
        reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
        hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference1, reference2)
        hybridShapeProject1.SolutionType = 0
        hybridShapeProject1.Normal = True
        hybridShapeProject1.SmoothingType = 0
        body1.InsertHybridShape(hybridShapeProject1)
        part1.InWorkObject = hybridShapeProject1
        part1.Update()
        # ==========================投影外型線==========================
        # ==========================長出==========================
        shapeFactory1 = part1.ShapeFactory
        reference3 = part1.CreateReferenceFromName("")
        pad1 = shapeFactory1.AddNewPadFromRef(reference3, 20)
        reference4 = part1.CreateReferenceFromObject(hybridShapeProject1)
        pad1.SetProfileElement(reference4)
        reference5 = part1.CreateReferenceFromObject(hybridShapeProject1)
        pad1.SetProfileElement(reference5)
        limit1 = pad1.FirstLimit
        limit1.LimitMode = 3
        hybridShapePlaneOffset2 = hybridShapes1.Item(plane_limit[2])
        reference6 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
        limit1.LimitingElement = reference6
        part1.Update()
        # ==========================長出==========================
        # ==========================offset==========================
        reference7 = part1.CreateReferenceFromName("")
        reference8 = part1.CreateReferenceFromName("")
        rectPattern1 = shapeFactory1.AddNewRectPattern(pad1, 2, 1, 20, 20, 1, 1, reference1, reference2, True, True, 0)
        # offset物件,數量(方向1),數量(方向2),間距(方向1),間距(方向2)
        rectPattern1.FirstRectangularPatternParameters = 0
        rectPattern1.SecondRectangularPatternParameters = 0
        linearRepartition1 = rectPattern1.FirstDirectionRepartition
        intParam1 = linearRepartition1.InstancesCount
        intParam1.Value = total_op_number - now_op_number
        length1 = linearRepartition1.Spacing
        length1.Value = 40  # 間距
        originElements1 = part1.OriginElements
        hybridShapePlaneExplicit1 = originElements1.PlaneXY
        reference9 = part1.CreateReferenceFromObject(hybridShapePlaneExplicit1)
        rectPattern1.SetFirstDirection(reference9)
        reference10 = part1.CreateReferenceFromObject(hybridShapePlaneExplicit1)
        rectPattern1.SetSecondDirection(reference10)
        part1.Update()
        # ==========================offset==========================
    # ==========================移除body(布林運算)==========================
    part1.InWorkObject = body2
    remove1 = shapeFactory1.AddNewRemove(body1)
    if mode == "lower_die":
        lower_die_pilot_punch_number = 0
        lower_die_pilot_punch_number += 1
        remove1.Name = "Lower-die-Pilot-punch-" + str(lower_die_pilot_punch_number)
    elif mode == "lower_pad":
        lower_pad_pilot_punch_number = 0
        lower_pad_pilot_punch_number += 1
        remove1.Name = "Lower-pad-Pilot-punch-" + str(lower_pad_pilot_punch_number)
    elif mode == "lower_die_set":
        lower_die_set_pilot_punch_number = 0
        lower_die_set_pilot_punch_number += 1
        remove1.Name = "Lower-die-set-Pilot-punch-" + str(lower_die_set_pilot_punch_number)
    elif mode == "stripper_plate":
        stripper_plate_pilot_punch_number = 0
        stripper_plate_pilot_punch_number += 1
        remove1.Name = "Stripper-plate-Pilot-punch-" + str(stripper_plate_pilot_punch_number)
    part1.Update()
    strParam1 = parameters1.Item("Properties\\B")
    strParam1.Value = "B: " + str(parameter_name[1]) + "-%%C" + str(Working_parameter) + str(
        0.01 * 2) + "割, 單+0.10(B型引導沖)"
    time.sleep(1)
    partDocument1.save()
    partDocument1.Close()


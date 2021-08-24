import win32com.client as win32
import time
import global_var as gvar

stripper_plate_height = int(gvar.strip_parameter_list[20])
Pilot_Punch_Stripper_punch_Diameter = 2  # 引導沖脫料釘直徑


def Plate_Stripper_Punch():
    stop_plate()
    Stripper()


def stop_plate():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    for i in range(1, gvar.PlateLineNumber + 1):
        partDocument1 = document.Open(gvar.save_path + "Stop_plate_" + str(i) + ".CATPart")
        stop_plate_stripper_punch_number = 0
        for n in range(1, int(gvar.StripDataList[76][0]) + 1):
            part1 = partDocument1.Part
            bodies1 = part1.Bodies
            body1 = bodies1.Item("Body.2")
            parameters1 = part1.Parameters
            # ==========================建點==========================
            hybridShapes1 = body1.HybridShapes
            hybridShapeFactory1 = part1.HybridShapeFactory
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
            hybridShapePointExplicit1 = parameters1.Item("plate_line_" + str(i) + "_stripper_pin_point_" + str(n))
            reference1 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
            hybridShapePointOnPlane1 = hybridShapeFactory1.AddNewPointOnPlaneWithReference(reference1, reference2, 0, 0)
            body1.InsertHybridShape(hybridShapePointOnPlane1)
            part1.InWorkObject = hybridShapePointOnPlane1
            part1.InWorkObject.Name = "plate_line_" + str(i) + "_stripper_pin_point_Project_" + str(n)
            part1.Update()
            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, -float(gvar.strip_parameter_list[17]))
            reference5 = part1.CreateReferenceFromObject(hybridShapePointOnPlane1)
            hybridShapePointCoord1.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "plate_line_" + str(i) + "_stripper_pin_point_Project_Project_" + str(n)
            part1.Update()
            # ==========================建點==========================
            # ==========================挖孔==========================
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(n)
            part1.Update()
            body2 = bodies1.Item("Body.2")
            shapeFactory1 = part1.ShapeFactory
            reference3 = part1.CreateReferenceFromObject(hybridShapePointOnPlane1)
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference3, reference4, 10)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.BottomType = 2
            hole1.ThreadingMode = 0
            hole1.CreateStandardThreadDesignTable(1)
            limit1.LimitMode = 3
            strParam1 = hole1.HoleThreadDescription
            strParam1.Value = "M" + str(Pilot_Punch_Stripper_punch_Diameter * 2)
            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_up_plane")
            reference6 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit1.LimitingElement = reference6
            stop_plate_stripper_punch_number += 1
            remove1 = shapeFactory1.AddNewRemove(body1)
            remove1.Name = "Stop-plate-Stripper-punch-" + str(stop_plate_stripper_punch_number)
            part1.Update()
            # ==========================挖孔==========================
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=*_point_*, All ")
        selection1.VisProperties.SetShow(1)
        selection1.Clear()
        selection1.Search("Name=*Sketch.*, All ")
        selection1.VisProperties.SetShow(1)
        time.sleep(1)
        partDocument1.save()
        partDocument1.Close()


def Stripper():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    for i in range(1, gvar.PlateLineNumber + 1):
        partDocument1 = document.Open(gvar.save_path + "Stripper_" + str(i) + ".CATPart")
        stripper_plate_stripper_punch_number = 0
        for n in range(1, int(gvar.StripDataList[76][0]) + 1):
            part1 = partDocument1.Part
            bodies1 = part1.Bodies
            body1 = bodies1.Item("Body.2")
            parameters1 = part1.Parameters
            # ==========================建點==========================
            hybridShapes1 = body1.HybridShapes
            hybridShapeFactory1 = part1.HybridShapeFactory
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
            hybridShapePointExplicit1 = parameters1.Item("plate_line_" + str(i) + "_stripper_pin_point_" + str(n))
            reference1 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePointExplicit1)
            hybridShapePointOnPlane1 = hybridShapeFactory1.AddNewPointOnPlaneWithReference(reference1, reference2, 0, 0)
            body1.InsertHybridShape(hybridShapePointOnPlane1)
            part1.InWorkObject = hybridShapePointOnPlane1
            part1.InWorkObject.Name = "plate_line_" + str(i) + "_stripper_pin_point_Project_" + str(n)
            part1.Update()
            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, -(stripper_plate_height - 12))
            reference3 = part1.CreateReferenceFromObject(hybridShapePointOnPlane1)
            hybridShapePointCoord1.PtRef = reference3
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "plate_line_" + str(i) + "_stripper_pin_point_Project_Project_" + str(n)
            part1.Update()
            # ==========================建點==========================
            # ==========================挖孔==========================
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(n)
            part1.Update()
            body2 = bodies1.Item("Body.2")
            shapeFactory1 = part1.ShapeFactory
            reference4 = part1.CreateReferenceFromObject(hybridShapePointOnPlane1)
            reference5 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference4, reference5, 10)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            length1 = hole1.Diameter
            length1.Value = Pilot_Punch_Stripper_punch_Diameter * 2
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            length2 = limit1.dimension
            length2.Value = stripper_plate_height - 12
            part1.Update()
            reference6 = part1.CreateReferenceFromObject(hybridShapePointOnPlane1)
            reference7 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole2 = shapeFactory1.AddNewHoleFromRefPoint(reference6, reference7, 10)
            hole2.Type = 0
            hole2.AnchorMode = 0
            hole2.BottomType = 0
            limit2 = hole2.BottomLimit
            limit2.LimitMode = 0
            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_up_plane")
            reference6 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit2.LimitingElement = reference6
            part1.Update()
            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            stripper_plate_stripper_punch_number += 1
            remove1.Name = "Stripper-plate-Stripper-punch-" + str(stripper_plate_stripper_punch_number)
            part1.Update()
            # ==========================挖孔==========================
        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=*_point_*, All ")
        selection1.VisProperties.SetShow(1)
        selection1.Clear()
        selection1.Search("Name=*Sketch.*, All ")
        selection1.VisProperties.SetShow(1)
        time.sleep(1)
        partDocument1.save()
        partDocument1.Close()


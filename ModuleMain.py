import global_var as gvar
import win32com.client as win32
import csv
import punch
import A_punch
import BinderPlate
import BinderPlateCut
import LowerDieInsert
import SplintInsert
import StripperInsert
import CutCavity
import CutPlate
import Stripper
import StopPlate
import Splint
import UpPlate
import LowerDieSet
import UpperDieSet
import Pilot_Punch
import Plate_Locking_Hole
import Plate_Pin_Hole
import Plate_Inner_Guiding_Post
import Plate_SBT_Hole
import Plate_Pilot_Punch_Hole
import Plate_Stripper_Punch
import Standard_Part
import time


def ModuleMain():
    with open(gvar.strip_parameters_file_root) as csvFile:
        rows = csv.reader(csvFile)
        strip_parameter_list = tuple(tuple(rows)[0])
        gvar.strip_parameter_list = strip_parameter_list
    # ------------------------------------------------參數初始化
    Sketch_position = "Hybridbody"  # 用F的程式接放在依據母群裡
    op_number = 0  # OP站從0開始
    all_part_number = 0  # 一開始無任何零件輸出,所以為0
    # ------------------------------------------------參數初始化
    # --------------------------------------------------------------起手式
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    j = gvar.PlateLineNumber
    A_punch_H = 9
    for now_plate_line_number in range(1, j + 1):
        punch.PunchMaking(now_plate_line_number)  # 沖頭
        (A_punch_H) = A_punch.APunchMaking(now_plate_line_number)  # A沖
        BinderPlateCut.BinderPlateCut(now_plate_line_number)  # 沖頭挖槽
        BinderPlate.BinderPlateSystem(now_plate_line_number)  # 沖頭壓板
        partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
        LowerDieInsert.LowerDieInsert(now_plate_line_number)  # 下模入子
        SplintInsert.SplintInsert(now_plate_line_number, A_punch_H)  # 上夾板入子
        StripperInsert.StripperInsert(now_plate_line_number, A_punch_H)  # 脫料板入子
        partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
        CutCavity.CutCavity(now_plate_line_number)  # 下模板
        partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
        CutPlate.CutPlate(now_plate_line_number)  # 下墊板
        partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
        Stripper.Stripper(now_plate_line_number)  # 脫料板
        partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
        StopPlate.StopPlate(now_plate_line_number)  # 止擋板
        partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
        Splint.Splint(now_plate_line_number)  # 上夾板
        partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
        UpPlate.UpPlate(now_plate_line_number)  # 上墊板
    partDocument1 = documents1.Open(gvar.open_path + "lower_die_set.CATPart")
    (lower_die_set_length, lower_die_set_width) = LowerDieSet.LowerDieSet()  # 下模座
    partDocument1 = documents1.Open(gvar.open_path + "upper_die_set.CATPart")
    UpperDieSet.UpperDieSet(lower_die_set_length, lower_die_set_width)  # 上模座
    (Pilot_Punch_data)=Pilot_Punch.Pilot_Punch()# 引導沖
    time.sleep(2)
    # --------------------------------挖孔--------------------------------(Momo_Function_Hole)
    (Bolt_data, BoltQuantity, CB_data) = Plate_Locking_Hole.Plate_Locking_Hole()  # '螺栓
    (Pin_data, Pinhole_data, PinQuantity) = Plate_Pin_Hole.Plate_Pin_Hole()  # '合銷
    (Inner_Guiding_data, InnerGuidingQuantity) = Plate_Inner_Guiding_Post.Plate_Inner_Guiding_Post()  # '內導柱/套
    (SBT_data, SBT_CB_data, SBTQuantity) = Plate_SBT_Hole.Plate_SBT_Hole()  # '等高螺栓
    Plate_Pilot_Punch_Hole.Plate_Pilot_Punch_Hole()  # '引導沖孔
    Plate_Stripper_Punch.Plate_Stripper_Punch()  # '脫料釘
    # --------------------------------挖孔--------------------------------(Momo_Function_Hole)
    (Pin_data, SBT_data, outer_Guiding_data, stripper_pin_data) = Standard_Part.Standard_Part(Bolt_data, CB_data,
                                                                                              Pin_data,
                                                                                              Inner_Guiding_data,
                                                                                              SBT_data, SBT_CB_data)
    print('Pin_data')
    print(Pin_data)
    print("SBT_data:")
    print(SBT_data)
    print("outer_Guiding_data:")
    print(outer_Guiding_data)
    print("Bolt_data:")
    print(Bolt_data)
    print("CB_data:")
    print(CB_data)
    print("BoltQuantity:")
    print(BoltQuantity)
    print("PinQuantity:")
    print(PinQuantity)
    print("Inner_Guiding_data:")
    print(Inner_Guiding_data)
    print("InnerGuidingQuantity:")
    print(InnerGuidingQuantity)
    print("SBT_CB_data:")
    print(SBT_CB_data)
    print("SBTQuantity:")
    print(SBTQuantity)
    return Bolt_data, CB_data, BoltQuantity, Pin_data, PinQuantity, Inner_Guiding_data, InnerGuidingQuantity, SBT_data, SBT_CB_data, SBTQuantity, outer_Guiding_data

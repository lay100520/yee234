import win32com.client as win32
import csv
import openpyxl
import global_var as gvar
import defs

# with open(gvar.strip_parameters_file_root) as csvFile:
#     rows = csv.reader(csvFile)
#     strip_parameter_list = list(tuple(rows)[0])
#     gvar.strip_parameter_list = strip_parameter_list
catapp = win32.Dispatch('CATIA.Application')
documents1 = catapp.Documents
# gvar.PlateLineNumber =1
# gvar.StripDataList[1][1]=390.0
# gvar.StripDataList[37][1][1]=2
# gvar.StripDataList[38][1][2]=1
# gvar.StripDataList[38][1][3]=4
# gvar.StripDataList[38][1][4]=1
# gvar.StripDataList[38][1][6]=4
# gvar.StripDataList[38][1][7]=1
# gvar.StripDataList[79][1][0]=4
# gvar.StripDataList[79][1]=40
# gvar.StripDataList[80][1]=29.99999999999997
# gvar.StripDataList[81][1]=420.0
# gvar.StripDataList[82][1]=75.0
# gvar.StripDataList[83][1]=234.39999999999998
# gvar.StripDataList[84][1]=390.0
# gvar.StripDataList[85][1]=159.39999999999998
import Strip
Strip.StripSystem()
import ModuleMain
(Bolt_data, CB_data, BoltQuantity, Pin_data, PinQuantity, Inner_Guiding_data, InnerGuidingQuantity, SBT_data, SBT_CB_data, SBTQuantity, outer_Guiding_data) = ModuleMain.ModuleMain()
print(gvar.all_part_number)
print(gvar.all_part_name)
import drafting
drafting.drafting()
# ============↓組立↓============
import Locking_assemble
Locking_assemble.Lockingassemble(CB_data,BoltQuantity)  # 螺栓
import Pin_assemble
Pin_assemble.Pin_assemble(Pin_data,PinQuantity)  # 合銷
import Inner_Guiding_post_assemble
Inner_Guiding_post_assemble.Inner_Guiding_post_assemble(Inner_Guiding_data, InnerGuidingQuantity)  # 內導柱/套
import SBT_assemble
SBT_assemble.SBT_assemble(SBT_data)  # 等高螺栓
import Plate_out_Guide_posts
Plate_out_Guide_posts.Plate_out_Guide_posts(outer_Guiding_data)  # 外導柱/套
import out_Guide_posts_locking_assemble
out_Guide_posts_locking_assemble.out_Guide_posts_locking_assemble(outer_Guiding_data)  # 外導柱螺栓
import limiting_assembly
limiting_assembly.limiting_assembly()  # 限位柱
import allotype_punch_assemble
allotype_punch_assemble.allotype_punch_assemble()
import assemble_hide
assemble_hide.assemble_hide()  # 隱藏組力拘束
# ============↑組立↑============
productDocument1 = catapp.ActiveDocument
productDocument1.save()
import BOM
BOM.BOMMaking()
import os

output_file_root = str()
import_file_root = str()
Mold_status = str('閉模')
strip_parameter_list = []
file_path = os.path.dirname(os.path.realpath(__file__))
strip_parameters_file_root = str(file_path+'\\strip_parameter.csv')
# 儲存路徑 (output 零件)
save_path = str(file_path + '\\auto\\catia_output-GTCA022\\')
# 母檔輸入路徑 (input Data)
open_path = str(file_path + "\\auto\\catia_input-GTCA022\\")
# 模具規範路徑
die_rule_path = str(file_path + "\\auto\\die_rule\\")
# 2D出圖路徑
drafting_output_path = str(file_path + "\\auto\\drafting_output-GTCA022\\")
# 標準零件路徑
standard_path = str(file_path + "\\auto\\standard_Assembly\\")
# 製作一半的BOM表儲存路徑
onwork_BOM_open = str(file_path + "\\auto\\BOM表\\")
# BOM表儲存路徑
BOM_output_path = str(file_path + "\\auto\\BOM_output-GTCA022\\")
all_part_number = int()
all_part_name = [""]*99
SumKeywayCircleLine = float()
SumBootsPartCircleLine = float()
SumRivetHoleCircleLine = float()
SumCentralPocketCircleLine = float()
SumContourCircleCircleLine = float()
StripDataList = [[[0]*10for i in range(10)]for ii in range(99)]
PlateLineNumber = int()
Shear_Strength = int()
die_type = str()

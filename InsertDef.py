import global_var as gvar
import win32com.client as win32
import defs


def InsertInterferance(plate_line_parameter, interferance_pline_name, open_name, now_plate_line_number):
    g = now_plate_line_number
    total_op_number = int(gvar.strip_parameter_list[2])
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents
    partDocument1 = documents1.Open(gvar.open_path + "Data1.CATPart")
    partDocument2 = documents1.Open(gvar.open_path + open_name + ".CATPart")
    insert_interferance_count = [[[0] * 99 for I in range(99)] for I in range(99)]
    # ---------------------------------------------------------------------------------
    defs.window_change(partDocument1, partDocument2)
    part2 = partDocument1.part
    bodies1 = part2.Bodies
    body1 = bodies1.Item("Body.2")
    body1.Name = "Body.4"
    selection2 = partDocument1.Selection
    # ------------------------------------------------------------↓定義數據
    length = [99]  # 定義為長度抓出參數中數據值
    g = now_plate_line_number
    insert_X_max = [[0.0] * 50 for i in range(50)]
    insert_X_min = [[0.0] * 50 for i in range(50)]
    insert_Y_max = [[0.0] * 50 for i in range(50)]
    insert_Y_min = [[0.0] * 50 for i in range(50)]
    insert_interferance_2 = [[None] * 50 for i in range(50)]
    insert_interferance_max = [[""] * 50 for i in range(50)]
    insert_interferance_location_max = [[0.0] * 50 for i in range(50)]
    insert_interferance_totle = [[0.0] * 50 for i in range(50)]
    # ------------------------------------------------------------↑
    parameters2 = part2.Parameters
    length2 = parameters2.Item("strip_width")
    strip_width = length2.Value
    bodies2 = part2.Bodies
    body2 = bodies2.Item("Body.4")
    hybridShapeFactory2 = part2.HybridShapeFactory
    hybridShapes2 = body2.HybridShapes
    hybridShapePointCoord2 = hybridShapes2.Item("base_point")
    hybridShapePointCoord2.Y.Value = -strip_width / 2 - 100
    part2.Update()
    # ------------------------------------------------------------↓     記錄數據
    for i in range(1, g + 1):  # 第幾塊模板
        for tt in range(1, total_op_number + 1):  # 第幾OP
            op_number = 10 * tt
            for ii in range(1, int(plate_line_parameter[i][tt]) + 1):  # OP中第幾個元素
                selection2.Clear()
                selection2.Search(
                    "Name = plate_line_" + str(i) + "_op" + str(op_number) + interferance_pline_name + str(ii))
                if selection2.Count == 1:
                    part2.Parameters.Item("cut_line_formula_1").OptionalRelation.Modify(
                        "die\plate_line_" + str(i) + "_op" + str(op_number) + interferance_pline_name + str(ii))  # 草圖置換
                part2.Update()
                length[1] = part2.Parameters.Item("X_max_base")
                length[2] = part2.Parameters.Item("X_min_base")
                length[3] = part2.Parameters.Item("Y_max_base")
                length[4] = part2.Parameters.Item("Y_min_base")
                insert_X_max[tt][ii] = length[1].Value
                insert_X_min[tt][ii] = length[2].Value
                insert_Y_max[tt][ii] = length[3].Value
                insert_Y_min[tt][ii] = length[4].Value
    # ------------------------------------------------------------↑
    for i in range(1, 3):  # 第幾塊模板
        for tt in range(1, total_op_number + 1):  # 第幾OP
            op_number = 10 * tt
            for ii in range(1, int(plate_line_parameter[i][tt]) + 1):  # OP中第幾個元素
                # ------------------------------------------------------------------------------------------------↓  搜尋是否為干涉的入子
                eee = [None]
                fff = [None]
                for ee in range(1, total_op_number + 1):  # 組別
                    for ff in range(1, 11):  # 干涉數量
                        if insert_interferance_count[ee][ff][1] == tt and insert_interferance_count[ee][ff][2] == ii:
                            eee = ee
                            fff = ff
                if eee == None:
                    for ee in range(1, total_op_number + 1):
                        if insert_interferance_count[ee][1][1] == 0:
                            fff = 1
                            eee = ee
                if eee != None:
                    ee = eee
                    ff = fff
                    eee = [None]
                    fff = [None]
                # ------------------------------------------------------------------------------------------------↑
                insert_count = 0  # 初始化跟幾個入子干涉
                for aa in range(tt, tt + 2):  # 比較的自己OP站 和下 一OP站 的數值
                    if aa == total_op_number + 1:  # 最後一站不用比較 所以跳出
                        break
                    for bb in range(1, int(plate_line_parameter[i][aa]) + 1):  # 比較的OP中第幾個元素
                        insert_interferance_decide = 0
                        # ------------------------------------------------------------------------------------------------↓   X座標比較
                        if insert_X_max[tt][ii] < insert_X_max[aa][bb]:  # 比較情況1   A<B
                            if insert_X_max[tt][ii] > insert_X_min[aa][bb] or [
                                insert_X_min[aa][bb] - insert_X_max[tt][ii]] < 5:  # 比較情況2
                                insert_interferance_decide = insert_interferance_decide + 1
                        elif insert_X_max[tt][ii] > insert_X_max[aa][bb]:
                            if insert_X_max[aa][bb] > insert_X_min[tt][ii] or (
                                    insert_X_min[tt][ii] - insert_X_max[aa][bb]) < 5:  # A>B
                                insert_interferance_decide = insert_interferance_decide + 1
                        elif insert_X_max[tt][ii] == insert_X_max[aa][bb]:
                            insert_interferance_decide = insert_interferance_decide + 1
                        # ------------------------------------------------------------------------------------------------↑
                        # ------------------------------------------------------------------------------------------------↓   Y座標比較
                        if insert_Y_max[tt][ii] < insert_Y_max[aa][bb]:  # 比較情況1   A<B
                            if insert_Y_max[tt][ii] > insert_Y_min[aa][bb] or insert_Y_min[aa][bb] - insert_Y_max[tt][
                                ii] < 5:
                                insert_interferance_decide = insert_interferance_decide + 1
                        elif insert_Y_max[tt][ii] > insert_Y_max[aa][bb]:
                            if insert_Y_max[aa][bb] > insert_Y_min[tt][ii] or insert_Y_min[tt][ii] - insert_Y_max[aa][
                                bb] < 5:  # 比較情況1   A>B
                                insert_interferance_decide = insert_interferance_decide + 1
                        elif insert_Y_max[tt][ii] == insert_Y_max[aa][bb]:
                            insert_interferance_decide = insert_interferance_decide + 1
                        # ------------------------------------------------------------------------------------------------↑
                        if insert_Y_max[tt][ii] == insert_Y_max[aa][bb] and insert_X_max[tt][ii] == insert_X_max[aa][
                            bb]:
                            insert_interferance_decide = 0
                        PP1 = [None]
                        PPP1 = [None]
                        for PP in range(1, total_op_number + 1):
                            for PPP in range(1, 10):
                                if insert_interferance_count[PP][PPP][1] == aa and insert_interferance_count[PP][PPP][
                                    2] == bb:
                                    PP1 = PP
                                    PPP1 = PPP
                        if PP1 == None:
                            if insert_interferance_decide == 2:
                                insert_count = insert_count + 1
                                if insert_count == 1:
                                    insert_interferance_count[ee][ff][1] = tt
                                    insert_interferance_count[ee][ff][2] = ii
                                    insert_interferance_2[ee][ff] = "plate_line_" + str(i) + "_op" + str(
                                        op_number) + interferance_pline_name + str(ii)
                                    insert_interferance_count[ee][ff + insert_count][1] = aa
                                    insert_interferance_count[ee][ff + insert_count][2] = bb
                                else:
                                    insert_interferance_count[ee][ff + insert_count][1] = aa
                                    insert_interferance_count[ee][ff + insert_count][2] = bb
    for tt in range(1, total_op_number + 1):  # 組別
        insert_interferance_max[tt][1] = insert_X_max[int(insert_interferance_count[tt][1][1])][
            int(insert_interferance_count[tt][1][2])]  # 1=X_max
        insert_interferance_max[tt][3] = insert_X_min[int(insert_interferance_count[tt][1][1])][
            int(insert_interferance_count[tt][1][2])]  # 3=X_min
        insert_interferance_max[tt][5] = insert_Y_max[int(insert_interferance_count[tt][1][1])][
            int(insert_interferance_count[tt][1][2])]  # 5=Y_max
        insert_interferance_max[tt][7] = insert_Y_min[int(insert_interferance_count[tt][1][1])][
            int(insert_interferance_count[tt][1][2])]  # 7=Y_min
    for PP in range(1, total_op_number + 1):
        for PPP in range(1, 11):
            # ------------------------------------------------------------------------------------------------↓   計算總共有多少入子
            if insert_interferance_count[PP][PPP][1] != 0:
                insert_interferance_totle[PP] = PPP
            # ------------------------------------------------------------------------------------------------↑
            if insert_interferance_max[PP][1] <= insert_X_max[int(insert_interferance_count[PP][PPP][1])][
                insert_interferance_count[PP][PPP][2]]:
                insert_interferance_max[PP][1] = insert_X_max[int(insert_interferance_count[PP][PPP][1])][
                    insert_interferance_count[PP][PPP][2]]
                insert_interferance_location_max[PP][1] = insert_interferance_count[PP][PPP][1]
                insert_interferance_location_max[PP][2] = insert_interferance_count[PP][PPP][2]
            if insert_X_min[insert_interferance_count[PP][PPP][1]][int(insert_interferance_count[PP][PPP][1])] != 0:
                if insert_interferance_max[PP][3] >= insert_X_min[int(insert_interferance_count[PP][PPP][1])][
                    int(insert_interferance_count[PP][PPP][2])]:
                    insert_interferance_max[PP][3] = insert_X_min[int(insert_interferance_count[PP][PPP][1])][
                        int(insert_interferance_count[PP][PPP][2])]
                    insert_interferance_location_max[PP][3] = insert_interferance_count[PP][PPP][1]
                    insert_interferance_location_max[PP][4] = insert_interferance_count[PP][PPP][2]
            if insert_interferance_max[PP][5] <= insert_Y_max[int(insert_interferance_count[PP][PPP][1])][
                int(insert_interferance_count[PP][PPP][2])]:
                insert_interferance_max[PP][5] = insert_Y_max[int(insert_interferance_count[PP][PPP][1])][
                    int(insert_interferance_count[PP][PPP][2])]
                insert_interferance_location_max[PP][5] = insert_interferance_count[PP][PPP][1]
                insert_interferance_location_max[PP][6] = insert_interferance_count[PP][PPP][2]
            if insert_Y_min[int(insert_interferance_count[PP][PPP][2])][
                int(insert_interferance_count[PP][PPP][1])] != 0:
                if insert_interferance_max[PP][7] >= insert_Y_min[int(insert_interferance_count[PP][PPP][1])][
                    int(insert_interferance_count[PP][PPP][2])]:
                    insert_interferance_max[PP][7] = insert_Y_min[int(insert_interferance_count[PP][PPP][1])][
                        int(insert_interferance_count[PP][PPP][2])]
                    insert_interferance_location_max[PP][7] = insert_interferance_count[PP][PPP][1]
                    insert_interferance_location_max[PP][8] = insert_interferance_count[PP][PPP][2]
    partDocument2.Close()
    return insert_interferance_count

import numpy as np
import matplotlib.pyplot as plt
import xlrd
import xlwt


def open_excel():
    sensor_data = []
    try:
        global book
        book = xlrd.open_workbook("origin.xlsx")  # 文件名，把文件与py文件放在同一目录下
    except:
        print("open excel file failed!")
    try:
        global sheet
        sheet = book.sheets()[0]  # execl里面的worksheet1
    except:
        print("locate worksheet in excel failed!")
    for i in range(sheet.ncols):  # 第一行是标题名，对应表中的字段名所以应该从第二行开始
        col_data = sheet.col_values(i, 1)
        sensor_data.append(col_data)
    return sensor_data


def get_data(data, start, stop,):
    back_data = []
    for i in range(len(data)):
        back_data.append(data[i][start:stop])
    return back_data


def get_similar_value_index(data, value):
    minus_data = []
    for i in data:
        minus_data.append(abs(i - value))

    return minus_data.index(min(minus_data))


# 获取一个相应周期的特征
def get_curve_attributions(data):
    attributions = []
    attributions_init_value = []
    # attributions_min_value_index = []
    attributions_response_value = []
    # attributions_max_value_index = []
    attributions_time_response10_90 = []
    attributions_recover_value = []
    # attributions_time_recover10_90 = []
    attributions_response_sensitivity = []
    attributions_recover_sensitivity = []
    length = len(data[0])
    for i in range(1, len(data)):
        "取响应阶段初始值"
        init_value = min(data[i][:length//2])
        # attributions_min_value_index.append(data[i].index(min_value))
        attributions_init_value.append(init_value)
        "取响应阶段最大值"
        response_value = max(data[i][:length])
        # attributions_max_value_index.append(data[i].index(max_value))
        attributions_response_value.append(response_value)
        "求响应敏感度"
        attributions_response_sensitivity.append((response_value - init_value) / init_value)
        "取响应10-90%的时间"
        response_10percent_value = (9*init_value + response_value)/10   # 取上升10%的时间
        response_10percent_index = get_similar_value_index(data[i][data[i].index(init_value):data[i].
                                                           index(response_value)], response_10percent_value)

        response_90percent_value = (init_value + 9 * response_value) / 10   # 取上升90%的时间
        response_90percent_index = get_similar_value_index(data[i][data[i].index(init_value):data[i].
                                                           index(response_value)], response_90percent_value)
        attributions_time_response10_90.append(response_90percent_index - response_10percent_index)
        "取恢复阶段结束后的值"
        recover_value = min(data[i][length//2:length])
        attributions_recover_value.append(recover_value)
        "求恢复敏感度"
        attributions_recover_sensitivity.append((response_value - recover_value) / recover_value)
        """  "取恢复10-90%的时间（去恢复值索引时出问题了）"
        recover_10percent_value = (9*response_value + recover_value)/10

        recover_10percent_index = get_similar_value_index(data[i][data[i].index(response_value):data[i].
                                                          index(recover_value)], recover_10percent_value)

        recover_90percent_value = (response_value + 9 * recover_value) / 10

        recover_90percent_index = get_similar_value_index(data[i][data[i].index(response_value):data[i].
                                                          index(recover_value)], recover_90percent_value)

        attributions_time_recover10_90.append(recover_90percent_index - recover_10percent_index)"""

    attributions.append(attributions_init_value)
    attributions.append(attributions_response_value)
    attributions.append(attributions_recover_value)
    attributions.append(attributions_response_sensitivity)
    attributions.append(attributions_recover_sensitivity)
    attributions.append(attributions_time_response10_90)
    # attributions.append(attributions_time_recover10_90)

    return attributions


def data_write(file_path, datas):
    row = 1
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=False)  # 创建sheet
    sheet1.write(0, 0, '标签')
    sheet1.write_merge(0, 0, 1, 4, '初始值')
    sheet1.write_merge(0, 0, 5, 8, '响应灵敏度')
    sheet1.write_merge(0, 0, 9, 12, '恢复灵敏度')
    sheet1.write_merge(0, 0, 13, 16, '响应时间')
    # sheet1.write_merge(0, 0, 17, 20, '恢复时间')
    for i in range(len(datas)):
        sheet1.write(row, 0, '0,50')
        col = 1
        for j in range(4):
            sheet1.write(row, col, datas[i][0][j])
            col = col + 1
        for j in range(4):
            sheet1.write(row, col, datas[i][3][j])
            col = col + 1
        for j in range(4):
            sheet1.write(row, col, datas[i][4][j])
            col = col + 1
        for j in range(4):
            sheet1.write(row, col, datas[i][5][j])
            col = col + 1
        """for j in range(4):
            sheet1.write(row, col, datas[i][6][j])
            col = col + 1"""
        row = row + 1
    f.save(file_path)  # 保存文件


"""SO2 30ppm, SOF2 30ppm"""
curve_start_position_30_30 = [5420, 5720, 6020, 6320, 6620, 6920, 7220, 7520]  # 第一次循环数据不好，没加 ，应为9组数据
curve_stop_position_30_30 = [5820, 6120, 6420, 6720, 7020, 7320, 7620, 7920]

"""SO2 50ppm, SOF2 30ppm"""
curve_start_position_50_30 = [9020, 9320, 9620, 9920, 10220, 10520, 10820, 11120]
curve_stop_position_50_30 = [9420, 9720, 10020, 10320, 10620, 10920, 11220, 11520]

"""SO2 50ppm, SOF2 0ppm"""
curve_start_position_50_0 = [12320, 12620, 12920, 13220, 13520, 13820, 14120,
                             14420, 14720, 16220, 16520, 16820, 17120, 17420, 17720]
curve_stop_position_50_0 = [12720, 13020, 13320, 13620, 13920, 14220, 14520,
                            14820, 15120, 16620, 16920, 17220, 17520, 17820, 18120]

"""SO2 0ppm, SOF2 30ppm"""
curve_start_position_0_30 = [18920, 19220, 19520, 19820, 20120, 20420, 20720, 21020, 21320]
curve_stop_position_0_30 = [19320, 19620, 19920, 20220, 20520, 20820, 21120, 21420, 21720]

"""SO2 0ppm, SOF2 50ppm"""
curve_start_position_0_50 = [22520, 22820, 23120, 23420, 23720, 24020, 24320, 24620, 24920]
curve_stop_position_0_50 = [22920, 23220, 23520, 23820, 24120, 24420, 24720, 25020, 25320]

"""SO2 30ppm, SOF2 50ppm"""
curve_start_position_30_50 = [26120, 26420, 26720, 27020, 27320, 27620, 27920, 28220, 28520]
curve_stop_position_30_50 = [26520, 26820, 27120, 27420, 27720, 28020, 28320, 28620, 28920]

"""SO2 50ppm, SOF2 50ppm"""
curve_start_position_50_50 = [30020, 30320, 30620, 30920, 31220, 31520, 31820, 32120]
curve_stop_position_50_50 = [30420, 30720, 31020, 31320, 31620, 31920, 32220, 32520]

sensor_attribution = []
raw_data = open_excel()
"""deal_data = get_data(raw_data, 0, len(raw_data[1]))
plt.plot(deal_data[0], deal_data[1], 'ro')
plt.plot(deal_data[0], deal_data[2], 'ro')
plt.plot(deal_data[0], deal_data[3], 'ro')
plt.plot(deal_data[0], deal_data[4], 'ro')
plt.show()"""

"""print("请输入曲线的起始点和结束点的横坐标")
start = int(input("起始点坐标:"))""
stop = int(input("结束点坐标:"))"
one_period_data = get_data(raw_data, start, stop)
sensor_attribution.append(get_curve_attributions(one_period_data))"""

for i in range(len(curve_start_position_0_50)):
    one_period_data = get_data(raw_data, curve_start_position_0_50[i], curve_stop_position_0_50[i])
    sensor_attribution.append(get_curve_attributions(one_period_data))

for i in range(len(sensor_attribution)):
    print("第{}组曲线特征为:".format(i+1))
    for j in range(len(sensor_attribution[0][0])):
        print("\t传感器{}各特征为:".format(j+1))
        # for k in range(len(sensor_attribution[0])):
        print("\t初始值:{}".format(sensor_attribution[i][0][j]), end=' ')
        print("响应值:{}".format(sensor_attribution[i][1][j]), end=' ')
        print("恢复值:{}".format(sensor_attribution[i][2][j]), end=' ')
        print("响应灵敏度:{}".format(sensor_attribution[i][3][j]), end=' ')
        print("恢复灵敏度:{}".format(sensor_attribution[i][4][j]), end=' ')
        print("响应时间:{}".format(sensor_attribution[i][5][j]))
        # print("恢复时间:{}\n".format(sensor_attribution[i][6][j]))

data_write('0,50.xls', sensor_attribution)

import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import leastsq
import xlrd
import xlwt
from tkinter import _flatten


# 获取一个响应周期的特征
def get_curve_attributions(data):
    attributions = []
    attributions_time_response10_90 = []
    attributions_response_sensitivity = []
    attributions_recover_sensitivity = []
    length = len(data[0])
    for i in range(len(data)):
        "取响应阶段初始值"
        init_value = min(data[i][:length//2])
        init_value_index = data[i].index(init_value)
        print(init_value_index)
        "取响应阶段最大值"
        response_value = max(data[i][:length])
        response_value_index = data[i].index(response_value)
        print(response_value_index)
        "求响应敏感度"
        attributions_response_sensitivity.append(10 * (response_value - init_value) / init_value)
        "取响应10-90%的时间"
        response_10percent_value = (9*init_value + response_value)/10   # 取上升10%的时间
        response_10percent_index = get_similar_value_index(data[i][init_value_index:response_value_index],
                                                           response_10percent_value)

        response_90percent_value = (init_value + 9 * response_value) / 10   # 取上升90%的时间
        response_90percent_index = get_similar_value_index(data[i][init_value_index:response_value_index],
                                                           response_90percent_value)
        attributions_time_response10_90.append((response_90percent_index - response_10percent_index) / 60)
        "取恢复阶段结束后的值"
        recover_value = min(data[i][length//2:length])
        "求恢复敏感度"
        attributions_recover_sensitivity.append(10 * (response_value - recover_value) / recover_value)
    attributions.append(attributions_response_sensitivity)
    attributions.append(attributions_recover_sensitivity)
    attributions.append(attributions_time_response10_90)
    attributions = list(_flatten(attributions))
    return attributions


def get_similar_value_index(data, value):
    minus_data = []
    for i in data:
        minus_data.append(abs(i - value))

    return minus_data.index(min(minus_data))


def open_excel():
    sensor_data = []
    try:
        global book
        book = xlrd.open_workbook("连续测量第一次.xls")  # 文件名，把文件与py文件放在同一目录下
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


def function(paras, x, n):
    y = 0
    for i in range(n+1):
        temp = paras[n-i]*x**i
        y = y + temp
    return y


# 误差函数，即拟合曲线所求的值与实际值的差
def error(paras, x, y, n):
    return function(paras, x, n) - y


def fitting(xdata, ydata, n_th_start, n_th_end, satrt_position):
    plt.scatter(xdata + satrt_position, ydata, color="green",  linewidth=2)
    for i in range(n_th_start, n_th_end + 1):
        p0 = [1] * (i + 1)
        Para = leastsq(error, p0, args=(xdata, ydata, i), maxfev=50000)
        x = xdata + satrt_position
        y1 = function(Para[0], xdata, i)
        y1 = y1.tolist()
        # print(y1)
        # coefficient.append(Para[0].tolist())
        plt.scatter(x, y1, color=color_list[i - n_th_start], linewidth=2)
        return y1


def data_write(file_path, datas, str):
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=False)  # 创建sheet
    for i in range(len(datas)):
        sheet1.write(i, 0, str)
        for j in range(len(datas[0])):
            sheet1.write(i, j+1, datas[i][j])
    f.save(file_path)  # 保存文件

curve_start_position = [0, 300]  # 第一次循环数据不好，没加 ，应为9组数据
curve_stop_position = [400, 700]
all_attributions = []
color_list = ["red", "orange", "yellow", "blue", "purple", "brown", "chocolate"]
raw_data = open_excel()
for i in range(len(curve_start_position)):
    one_period_data = get_data(raw_data, curve_start_position[i], curve_stop_position[i])
    X = np.array(one_period_data[0]) - curve_start_position[i]
    y = []
    for j in range(1, 5):
        Y = np.array(one_period_data[j])/10000
        print(Y)
        y_ = fitting(X, Y, 10, 10, curve_start_position[i])
        y.append(y_)
    one_period_attributions = get_curve_attributions(y)
    all_attributions.append(one_period_attributions)

# plt.legend()  # 绘制图例
plt.show()
# data_write('50,0.xls', all_attributions, '50,0')


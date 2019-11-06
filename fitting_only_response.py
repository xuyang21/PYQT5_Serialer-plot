import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import leastsq
import xlrd
import xlwt
from tkinter import _flatten


# 获取一个响应周期的特征
def get_curve_attributions(data):
    attributions = []
    attributions_response_sensitivity = []
    for i in range(len(data)):
        "取响应阶段初始值"
        init_value = min(data[i][500:1000])
        print(init_value)
        "取响应阶段最大值"
        response_value = max(data[i][:3000])
        print(response_value)
        "求响应敏感度"
        attributions_response_sensitivity.append(10 * (response_value - init_value)/init_value)

    attributions.append(attributions_response_sensitivity)
    attributions = list(_flatten(attributions))
    print(attributions)
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
        book = xlrd.open_workbook("电科院实验结果完整版(10, 20, 30, 40, 50).xls")  # 文件名，把文件与py文件放在同一目录下
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
        xnew = np.arange(0, 400, 0.1)
        # x = xdata + satrt_position
        y1 = function(Para[0], xnew, i)
        xnew = xnew + satrt_position
        y1 = y1.tolist()
        # print(y1)
        # coefficient.append(Para[0].tolist())
        plt.scatter(xnew, y1, color=color_list[i - n_th_start], linewidth=2)
        return y1


def data_write(file_path, datas, str):
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=False)  # 创建sheet
    for i in range(len(datas)):
        sheet1.write(i, 0, str)
        for j in range(len(datas[0])):
            sheet1.write(i, j+1, datas[i][j])
    f.save(file_path)  # 保存文件

"""6.67%origin"""
# curve_start_position= [3050,3350,3650,3950,4250,4550,4850,5150,5450,5750,6050,6350,6650,6950,7250,7550]
# curve_stop_position = [3450,3750,4050,4350,4650,4950,5250,5550,5850,6150,6450,6750,7050,7350,7650.7950]


"""6.67%"""
# curve_start_position= [3150,3450,3750,4050,4350,4650,4950,5250,5550,5850,6150,6450]
# curve_stop_position = [3550,3850,4150,4450,4750,5050,5350,5650,5950,6250,6550,6850]

"""10%"""
# curve_start_position= [7350,7650,7950,8250,8550,8850,9150,9450,9750,10050,10350,10650]
# curve_stop_position = [7750,8050,8350,8650,8950,9250,9550,9850,10150,10450,10750,11050]

"""13.33%"""
# curve_start_position= [11550,11850,12150,12450,12750,13050,13350,13650,13950,14250,14550,14850]
# curve_stop_position = [11950,12250,12550,12850,13150,13450,13750,14050,14350,14650,14950,15250]

"""16.67"""
# curve_start_position=[15750,16050,16350,16650,16950,17250,17550,17850,18150,18450,18750,19050]
# curve_stop_position =[16150,16450,16750,17050,17350,17650,17950,18250,18550,18850,19150,19450]

"""20.00%"""
# curve_start_position=[19950,20250,20550,20850,21150,21450,21750,22050,22350,22650,22950,23250]
# curve_stop_position =[20350,20650,20950,21250,21550,21850,22150,22450,22750,23050,23350,23650]

"""23.33%"""
# curve_start_position=[24150,24450,24750,25050,25350,25650,25950,26250,26550,26850,27150,27450]
# curve_stop_position =[24550,24850,25150,25450,25750,26050,26350,26650,26950,27250,27550,27850]

"""26.67%"""
# curve_start_position=[28350,28650,28950,29250,29550,29850,30150,30450,30750,31050,31350,31650]
# curve_stop_position =[28750,29050,29350,29650,29950,30250,30550,30850,31150,31450,31750,32050]

"""30.00%"""
# curve_start_position=[32550,32850,33150,33450,33750,34050,34350,34650,34950,35250,35550,35850]
# curve_stop_position =[32950,33250,33550,33850,34150,34450,34750,35050,35350,35650,35950,36250]

"""33.33%"""
# curve_start_position=[36750,37050,37350,37650,37950,38250,38550,38850,39150,39450,39750,40050]
# curve_stop_position =[37150,37450,37750,38050,38350,38650,38950,39250,39550,39850,40150,40450]

"""36.67%"""
# curve_start_position= [40950,41250,41550,41850,42150,42450,42750,43050,43350,43650,43950,44250]
# curve_stop_position = [41350,41650,41950,42250,42550,42850,43150,43450,43750,44050,44350,44650]

"40.00%"
# curve_start_position= [45150,45450,45750,46050,46350,46650,46950,47250,47550,47850,48150,48450]
# curve_stop_position = [45550,45850,46150,46450,46750,47050,47350,47650,47950,48250,48550,48850]

"""重做的40.00%"""
# curve_start_position= [3150,3450,3750,4050,4350,4650]
# curve_stop_position = [3550,3850,4150,4450,4750,5050]

"43.33%"
# curve_start_position= [5550,5850,6150,6450,6750,7050]
# curve_stop_position = [5950,6250,6550,6850,7150,7450]

"46.67%"
# curve_start_position= [7950,8250,8550,8850,9150,9450]
# curve_stop_position = [8350,8650,8950,9250,9550,9850]

"53.33%"
# curve_start_position= [10350,10650,10950,11250,11550,11850]
# curve_stop_position = [10750,11050,11350,11650,11950,12250]

"60.00%"
# curve_start_position= [12750,13050,13350,13650,13950,14250]
# curve_stop_position = [13150,13450,13750,14050,14350,14650]

"66.67%"
# curve_start_position= [15150,15450,15750,16050,16350,16650]
# curve_stop_position = [15550,15850,16150,16450,16750,17050]


curve_start_position = [7900]
curve_stop_position = [8300]
all_attributions = []
color_list = ["red", "orange", "yellow", "blue", "purple", "brown", "chocolate"]
raw_data = open_excel()
for i in range(len(curve_start_position)):
    one_period_data = get_data(raw_data, curve_start_position[i], curve_stop_position[i])
    X = np.array(one_period_data[0]) - curve_start_position[i]
    y = []
    for j in range(1, 5):
        Y = np.array(one_period_data[j])/10000
        # print(Y)
        y_ = fitting(X, Y, 8, 8, curve_start_position[i])
        y.append(y_)
    one_period_attributions = get_curve_attributions(y)
    all_attributions.append(one_period_attributions)

# plt.legend()  # 绘制图例
plt.show()
# data_write('66.67%.xls', all_attributions, '66.67%')


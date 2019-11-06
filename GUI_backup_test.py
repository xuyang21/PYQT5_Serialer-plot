import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from PyQt5.QtCore import QTimer
from ui import Ui_Form
from PyQt5.QtGui import QIcon
import pyqtgraph as pg
import numpy as np
import xlwt, xlrd
import pickle
from scipy import interpolate
import json
from tkinter import _flatten
from scipy.optimize import leastsq


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
    for i in range(n_th_start, n_th_end + 1):
        p0 = [1] * (i + 1)
        Para = leastsq(error, p0, args=(xdata, ydata, i), maxfev=50000)
        # x = xdata + satrt_position
        y1 = function(Para[0], xdata, i)
        y1 = y1.tolist()
        # print(y1)
        return y1


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
        "取响应阶段最大值"
        response_value = max(data[i][:length])
        response_value_index = data[i].index(response_value)
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


def data_write(file_path, data):
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
    sheet1.write(0, 0, '时间')
    sheet1.write(0, 1, 'TGS2602')
    sheet1.write(0, 2, 'TGS2603')
    sheet1.write(0, 3, 'TGS2610')
    sheet1.write(0, 4, 'TGS2620')
    data = np.array(data)
    data = data.transpose()
    # 将数据写入第 i 行，第 j 列
    for i in range(data.shape[0]):
        for j in range(data.shape[1]):
            sheet1.write(i+1, j, data[i][j])
    f.save(file_path)  # 保存文件


class SerialSvmGui(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super().__init__()

        self.setupUi(self)
        self.setWindowIcon(QIcon('image.png'))
        self.init()
        self.received_data = []
        self.time_interval = 1000
        # 接收数据和发送数据数目置零
        self.showrecdatanum = 0
        self.lineEdit_showrecdatanum.setText(str(self.showrecdatanum))
        self.showtime = 0
        self.lineEdit_showtime.setText(str(self.showtime))

        self.flag = 1  # 暂停测量按钮
        self.datanum = 0
        self.time = []
        self.resistance_1 = []
        self.resistance_2 = []
        self.resistance_3 = []
        self.resistance_4 = []

        with open('erosion_record.json', 'r', encoding='utf-8') as file:
            try:
                self.erosion_accumulated = json.load(file)
            except:
                self.erosion_accumulated = []
        erosion_predict_total = np.array(self.erosion_accumulated)
        erosion_predict_total = (sum(erosion_predict_total) * 1000 // 1) / 1000
        # print(erosion_predict_total)
        self.lineEdit_show_erosion_total.setText(str(erosion_predict_total))
        # 显示剩余寿命
        lifetime_remain_percent = (1 - erosion_predict_total / 20000) * 100
        self.progressBar_lifetime_remain.setValue(lifetime_remain_percent)

        self.raw_data = open_excel()        # 绘图初始化
        pg.setConfigOption('background', 'w')
        pg.setConfigOption('foreground', 'k')
        self.pw = pg.PlotWidget()  # 实例化一个绘图部件
        self.pw.addLegend(size=(100, 50), offset=(0, 0))  # 设置图形的图例
        self.pw.showGrid(x=True, y=True, alpha=0.5)  # 设置图形网格的形式，我们设置显示横线和竖线，并且透明度惟0.5：
        self.pw.setLabel(axis='left', text=u'电阻值/Ω')
        self.pw.setLabel(axis='bottom', text=u'时间/s')
        self.pw.setXRange(min=0, max=100, padding=0)
        self.pw.setYRange(min=0, max=20000, padding=0)
        self.curve_1 = self.pw.plot(pen='k', name='TGS2602', symbol=0, symbolSize=3, symbolPen=(0, 0, 0),
                                    symbolBrush=(0, 0, 0))
        self.curve_2 = self.pw.plot(pen='r', name='TGS2603', symbol=0, symbolSize=3, symbolPen=(255, 0, 0),
                                    symbolBrush=(255, 0, 0))
        self.curve_3 = self.pw.plot(pen='b', name='TGS2610', symbol=0, symbolSize=3, symbolPen=(0, 0, 255),
                                    symbolBrush=(0, 0, 255))
        self.curve_4 = self.pw.plot(pen='g', name='TGS2620', symbol=0, symbolSize=3, symbolPen=(0, 255, 0),
                                    symbolBrush=(0, 255, 0))
        self.verticalLayout_graph.addWidget(self.pw)  # 添加绘图部件到网格布局层

    """定义信号与槽"""
    def init(self):
        # 开始测量
        self.start_measure_button.clicked.connect(self.start_measure)

        # 停止测量
        self.stop_measure_Button.clicked.connect(self.stop_measure)

        # 定时器设置
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.timeout)

        # 保存数据
        self.savedata_button.clicked.connect(self.save_data)

        # 开始识别
        self.svm_eval_button.clicked.connect(self.svm_eval)

        self.quit_Button.clicked.connect(self.quit)

    def start_measure(self):
        self.time_interval = int(self.lineEdit_time_interval.text())
        self.timer.start(self.time_interval)
        self.start_measure_button.setText("重新测量")
        # 清空之前测量的数据
        self.datanum = 0
        self.time.clear()
        self.resistance_1.clear()
        self.resistance_2.clear()
        self.resistance_3.clear()
        self.resistance_4.clear()

        # 接收数据和发送数据数目置零
        self.showrecdatanum = 0
        self.lineEdit_showrecdatanum.setText(str(self.showrecdatanum))
        self.showtime = 0
        self.lineEdit_showtime.setText(str(self.showtime))

    def stop_measure(self):
        if self.flag:
            self.timer.stop()
            self.stop_measure_Button.setText("继续测量")
        else:
            self.timer.start(self.time_interval)
            self.stop_measure_Button.setText("暂停测量")
        self.flag = abs(self.flag-1)

    def timeout(self):
        # self.ser.write(0x66)
        # self.received_data = self.ser.read(12)
        self.showrecdatanum = self.showrecdatanum + 12
        self.lineEdit_showrecdatanum.setText(str(self.showrecdatanum))
        self.datanum = self.datanum + 1
        self.showtime = self.datanum * self.time_interval/1000
        self.lineEdit_showtime.setText(str(self.showtime))

        self.time.append(self.raw_data[0][self.datanum])
        self.resistance_1.append(self.raw_data[1][self.datanum])
        self.resistance_2.append(self.raw_data[2][self.datanum])
        self.resistance_3.append(self.raw_data[3][self.datanum])
        self.resistance_4.append(self.raw_data[4][self.datanum])
        self.curve_1.setData(self.time, self.resistance_1)
        self.curve_2.setData(self.time, self.resistance_2)
        self.curve_3.setData(self.time, self.resistance_3)
        self.curve_4.setData(self.time, self.resistance_4)

    def save_data(self):
        data = [self.time, self.resistance_1, self.resistance_2, self.resistance_3, self.resistance_4]
        filename, ok = QFileDialog.getSaveFileName(self, "文件保存", "./", "All Files (*);;xls Files (*.xls)")
        if ok != '':
            data_write(filename, data)
            QMessageBox.information(self, "status", "文件已保存！")
        else:
            pass

    def svm_eval(self):
        list_SO2 = [30, 50, 50, 0, 0, 30, 50]
        list_SOF2 = [30, 30, 0, 30, 50, 50, 50]
        data = []
        start = int(self.lineEdit_startcoordinate.text())
        stop = int(self.lineEdit_endcoordinate.text())
        data.append(self.time)
        data.append(self.resistance_1)
        data.append(self.resistance_2)
        data.append(self.resistance_3)
        data.append(self.resistance_4)
        # 获取一周期内曲线的3种属性
        one_period_data = get_data(data, start, stop)
        X = np.array(one_period_data[0]) - start
        y = []
        for j in range(1, 5):
            Y = np.array(one_period_data[j]) / 10000
            y_ = fitting(X, Y, 10, 10, start)
            y.append(y_)
        one_period_attributions = get_curve_attributions(y)
        one_period_attributions = np.array(one_period_attributions)
        one_period_attributions = one_period_attributions.reshape(1, 12)
        # x = attributions.reshape(1, 12)
        # 载入训练好的SVM模型
        f2 = open('fitting.model', 'rb')
        s2 = f2.read()
        model = pickle.loads(s2)
        # 调用模型输出
        predicted = model.predict(one_period_attributions)
        print("预测标签为:{}". format(predicted))
        self.lineEdit_show_SO2F2.setText('0')
        self.lineEdit_show_H2O.setText('200')
        self.lineEdit_show_O2.setText('200')
        self.lineEdit_show_SO2.setText(str(list_SO2[predicted[0]]))
        self.lineEdit_show_SOF2.setText(str(list_SOF2[predicted[0]]))
        # SOF2浓度与烧蚀量的关系
        x = [1.19, 17.34, 108.93, 211.00, 338.76]
        y = [9.25, 21.60, 111.00, 408.00, 1096.00]
        # 进行插值
        f = interpolate.interp1d(x, y, kind="cubic")
        global erosion_predict
        # 预测此次烧蚀量
        erosion_predict = f(list_SOF2[predicted[0]])
        erosion_predict = (erosion_predict * 1000 // 1) / 1000
        print("此次开断烧蚀量为:{}".format(erosion_predict))
        self.lineEdit_show_erosion.setText(str(erosion_predict))
        # 计算到目前为止的烧蚀量
        self.erosion_accumulated.append(erosion_predict)
        erosion_predict_total = np.array(self.erosion_accumulated)
        erosion_predict_total = (sum(erosion_predict_total) * 1000 // 1) / 1000
        print("总烧蚀量为:{}".format(erosion_predict_total))
        self.lineEdit_show_erosion_total.setText(str(erosion_predict_total))
        # 显示剩余寿命
        lifetime_remain_percent = (1 - erosion_predict_total / 20000) * 100
        self.progressBar_lifetime_remain.setValue(lifetime_remain_percent)
        # 保存测得的烧蚀量
        with open('erosion_record.json', 'w', encoding='utf-8') as f:
            json.dump(self.erosion_accumulated, f, ensure_ascii=False)

    def quit(self):
        global app
        # app1 = QtWidgets.QApplication.instance()
        # app1.quit()
        app.quit()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = SerialSvmGui()
    window.setFixedSize(window.width(), window.height())
    # window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)   设置之后只有最小化按钮
    # window.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)
    window.show()
    sys.exit(app.exec_())

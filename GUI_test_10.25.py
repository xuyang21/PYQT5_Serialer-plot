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
        xnew = np.arange(0, 400, 0.1)
        y1 = function(Para[0], xnew, i)
        y1 = y1.tolist()
        return y1


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
        self.timer.stop()
        data = [self.time, self.resistance_1, self.resistance_2, self.resistance_3, self.resistance_4]
        try:
            start = int(self.lineEdit_startcoordinate.text())
            stop = int(self.lineEdit_endcoordinate.text())
            # 获取一周期内曲线的3种属性
        except:
            QMessageBox.critical(self, "Error", "请选择诊断范围")
            return None
        if (stop < start) or (len(self.time) < stop):
            QMessageBox.critical(self, "Error", "请选择正确的曲线范围")
            return None

        one_period_data = get_data(data, start, stop)
        X = np.array(one_period_data[0]) - start
        y = []
        for j in range(1, 5):
            Y = np.array(one_period_data[j]) / 10000
            y_ = fitting(X, Y, 8, 8, start)
            y.append(y_)
        try:
            one_period_attributions = get_curve_attributions(y)
        except:
            QMessageBox.critical(self, "Error", "无法获取特征")
            self.lineEdit_show_SO2F2.setText('0')
            self.lineEdit_show_H2O.setText('0')
            self.lineEdit_show_O2.setText('0')
            self.lineEdit_show_SO2.setText("0")
            self.lineEdit_show_SOF2.setText("0")
            self.lineEdit_show_erosion.setText("0")
            self.timer.start(self.time_interval)
            return None

        response = [0, 1.860235913, 3.166363845, 4.731732018, 6.191544929, 8.338593339, 10.20565894, 12.39826891,
                    14.50582942, 16.59688193, 21.09173263, 35.06726295, 42.46496963, 62.31995252, 98.70350851,
                    173.9876968, 200.0]
        ppm = [0, 4.9674825, 7.4475, 9.9275175, 12.4149825, 14.895, 17.3750175, 19.8624825, 22.3425, 24.8225175,
               27.3099825, 32.2700175, 34.7574825, 39.7175175, 44.685, 49.6524825, 50.00]
        # 进行插值
        f = interpolate.interp1d(response, ppm, kind="cubic")
        if one_period_attributions[1] > 200:
            predicted = 50.0000
        else:
            predicted = int(f(one_period_attributions[1]) * 10000)
            predicted = predicted / 10000
        print("预测浓度为:{}".format(predicted))

        so2_concentration = 0
        sof2_concentration = predicted
        so2f2_concentration = 0
        self.lineEdit_show_SO2F2.setText(str(so2f2_concentration))
        self.lineEdit_show_H2O.setText('0')
        self.lineEdit_show_O2.setText('0')
        self.lineEdit_show_SO2.setText(str(so2_concentration))
        self.lineEdit_show_SOF2.setText(str(sof2_concentration))
        # SOF2浓度与烧蚀量的关系
        if sof2_concentration <= 338.76:
            x = [0, 1.19, 17.34, 108.93, 211.00, 338.76]
            y = [0, 9.25, 21.60, 111.00, 408.00, 1096.00]
            # 进行插值
            f = interpolate.interp1d(x, y, kind="cubic")
            global erosion_predict
            # 预测此次烧蚀量
            erosion_predict = f(sof2_concentration)
        else:
            erosion_predict = 1096.00

        erosion_predict = (erosion_predict * 1000 // 1) / 1000

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

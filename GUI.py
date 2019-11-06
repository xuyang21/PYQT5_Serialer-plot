import sys
from PyQt5.QtSerialPort import QSerialPort, QSerialPortInfo
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from PyQt5.QtCore import QTimer
from ui import Ui_Form
from PyQt5.QtGui import QIcon
import pyqtgraph as pg
import numpy as np
import xlwt
import binascii
from scipy import interpolate
import json
from tkinter import _flatten
from scipy.optimize import leastsq
import time


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


def dec_to_binary_hex_char(dec):
    data = hex(dec)
    # print(data)
    txData = str(data)
    # print(txData)
    if len(txData) == 3:
        txData = list(txData)
        txData.insert(2, '0')
        txData = ''.join(txData)
    return txData[2:]


class SerialSvmGui(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super().__init__()

        self.setupUi(self)
        self.setWindowIcon(QIcon('image.png'))
        # self.stop_measure_Button.setEnabled(False)
        self.init()
        # 设置串口实例
        self.com = QSerialPort()
        self.port_check()

        self.time_interval = 1000
        # 接收数据和发送数据数目置零
        self.showrecdatanum = 0
        self.lineEdit_showrecdatanum.setText(str(self.showrecdatanum))
        self.showtime = 0
        self.lineEdit_showtime.setText(str(self.showtime))

        self.timer_complete_flag = 1
        self.flag = 1  # 暂停测量按钮
        self.time = []
        self.resistance_1 = []
        self.resistance_2 = []
        self.resistance_3 = []
        self.resistance_4 = []

        self.runtime = []
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

        txData = '66'
        self.command_upload_data = binascii.a2b_hex(txData)
        txData = '88'
        self.command_show_result = binascii.a2b_hex(txData)

        pg.setConfigOption('background', 'w')
        pg.setConfigOption('foreground', 'k')
        self.pw = pg.PlotWidget()  # 实例化一个绘图部件
        self.pw.addLegend(size=(100, 50), offset=(0, 0))  # 设置图形的图例
        self.pw.showGrid(x=True, y=True, alpha=0.5)  # 设置图形网格的形式，我们设置显示横线和竖线，并且透明度惟0.5：
        self.pw.setLabel(axis='left', text=u'电阻值/Ω')
        self.pw.setLabel(axis='bottom', text=u'时间/s')
        self.pw.setXRange(min=0, max=100, padding=0)
        self.pw.setYRange(min=0, max=20000, padding=0)
        self.curve_1 = self.pw.plot(pen='k', name='TGS2602', symbol='o', symbolSize=4, symbolPen=(0, 0, 0),
                                    symbolBrush=(0, 0, 0))
        self.curve_2 = self.pw.plot(pen='r', name='TGS2613', symbol='t', symbolSize=4, symbolPen=(255, 0, 0),
                                    symbolBrush=(255, 0, 0))
        self.curve_3 = self.pw.plot(pen='b', name='TGS2610', symbol='s', symbolSize=4, symbolPen=(0, 0, 255),
                                    symbolBrush=(0, 0, 255))
        self.curve_4 = self.pw.plot(pen='g', name='TGS2611', symbol='d', symbolSize=4, symbolPen=(0, 255, 0),
                                    symbolBrush=(0, 255, 0))
        self.verticalLayout_graph.addWidget(self.pw)  # 添加绘图部件到网格布局层

    """定义信号与槽"""
    def init(self):
        # 串口检测按钮
        self.checkserialport.clicked.connect(self.port_check)

        # 打开串口按钮
        self.open_button.clicked.connect(self.port_open)

        # 关闭串口按钮
        self.close_button.clicked.connect(self.port_close)

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

        # 退出应用
        self.quit_Button.clicked.connect(self.quit)

    # 串口检测
    def port_check(self):
        # 检测所有存在的串口，将信息存储在字典中
        self.serialportnum.clear()
        port_list = QSerialPortInfo.availablePorts()
        # print(self.port_list)
        if len(port_list) == 0:
            self.state_label.setText(" 无串口")
        else:
            for port in port_list:
                self.com.setPort(port)
                if self.com.open(QSerialPort.ReadWrite):
                    self.serialportnum.addItem('    ' + port.portName())
                    self.com.close()
            # self.serialportnum.removeItem(0)

    # 打开串口
    def port_open(self):
        if self.serialportnum.currentText().strip() == '无':
            QMessageBox.critical(self, "Port Error", "请选择串口！")
            return None
        self.com.setPortName(self.serialportnum.currentText().strip())
        self.com.setBaudRate(int(self.baudrate.currentText()))
        status = self.com.open(QSerialPort.ReadWrite)
        if status:
            QMessageBox.information(self, "Port status", "串口打开成功！")
            self.state_label.setText(self.com.portName() + " : OPEN")
        else:
            QMessageBox.critical(self, "Port Error", "串口打开失败！")
            return None

    # 关闭串口
    def port_close(self):
        try:
            self.timer.stop()
            self.com.close()
        except:
            pass
        self.state_label.setText(self.com.portName() + " : CLOSED")

    def start_measure(self):
        self.time_interval = int(self.lineEdit_time_interval.text())
        self.timer.start(self.time_interval)
        try:  # 第一次读不到数据，只能这里加一下
            self.com.write(self.command_upload_data)
        except:
            pass
        # try:
            # received_data = self.com.read(10)
            # print(received_data)
        # except:
        #     pass
        self.start_measure_button.setText("重新测量")
        # self.stop_measure_Button.setText("暂停测量")
        self.flag = 1  # 暂停测量按钮
        # 清空之前测量的数据
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

        self.stop_measure_Button.setEnabled(True)

    def stop_measure(self):
        self.time_interval = int(self.lineEdit_time_interval.text())
        self.timer.start(self.time_interval)

    def timeout(self):
        self.timer_complete_flag = 0
        start = time.time()
        # noinspection PyBroadException
        try:
            self.com.write(self.command_upload_data)
        except:
            QMessageBox.critical(self, "Port Error", "串口写入失败，请检查串口！")
            self.timer.stop()
            return None
        try:
            received_data = self.com.read(10)
            # print(received_data)

        except:
            QMessageBox.critical(self, "Port Error", "读数据出错，请检查连接！")
            self.timer.stop()
            return None
        received_data = binascii.b2a_hex(received_data).decode('ascii')
        data = []
        for i in range(0, len(received_data), 2):
            tem = received_data[i:i + 2]
            data.append(int(tem, 16))  # 将16进制字符串转成10进制整数
        print(data)
        self.showrecdatanum = self.showrecdatanum + 10
        self.showtime += self.time_interval / 1000
        self.time.append(self.showtime)
        try:
            temp1 = data[0] * 256 + data[1]
            temp2 = data[2] * 256 + data[3]
            temp3 = data[4] * 256 + data[5]
            temp4 = data[6] * 256 + data[7]
            temp5 = data[8] * 256 + data[9]
            resistance1 = 10000 * (temp1 - temp2) / temp2
            resistance2 = 10000 * (temp1 - temp3) / temp3
            resistance3 = 10000 * (temp1 - temp4) / temp4
            resistance4 = 10000 * (temp1 - temp5) / temp5
        except:
            QMessageBox.critical(self, "Port Error", "请检查串口")
            self.timer.stop()
            return None
        self.resistance_1.append(resistance1)
        # print(self.resistance_1)
        self.resistance_2.append(resistance2)
        self.resistance_3.append(resistance3)
        self.resistance_4.append(resistance4)

        self.lineEdit_showrecdatanum.setText(str(self.showrecdatanum))
        self.lineEdit_showtime.setText(str(self.showtime))

        self.curve_1.setData(self.time, self.resistance_1)
        self.curve_2.setData(self.time, self.resistance_2)
        self.curve_3.setData(self.time, self.resistance_3)
        self.curve_4.setData(self.time, self.resistance_4)
        end = time.time()
        # self.runtime.append(end-start)
        # print(max(self.runtime))
        print(end-start)
        self.timer_complete_flag = 1

    def save_data(self):
        self.timer.stop()
        data = [self.time, self.resistance_1, self.resistance_2, self.resistance_3, self.resistance_4]
        # filename, ok = QFileDialog.getSaveFileName(self, "文件保存", "./", "All Files (*);;xls Files (*.xls)")
        filename, ok = QFileDialog.getSaveFileName(self, "文件保存", "./", "xls Files (*.xls)")

        if ok != '':
            data_write(filename, data)
            QMessageBox.information(self, "status", "文件已保存！")
        else:
            pass

    def svm_eval(self):
        self.timer.stop()
        send_data = []
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

        if self.timer_complete_flag == 0:
            print("正在执行")
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
            self.lineEdit_show_H2O.setText('200')
            self.lineEdit_show_O2.setText('200')
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
            predicted = int(f(one_period_attributions[1])*10000)
            predicted = predicted/10000
        print("预测浓度为:{}".format(predicted))

        so2_concentration = 0
        byteH_so2 = (so2_concentration*10)//256
        send_data.append(dec_to_binary_hex_char(byteH_so2))
        byteL_so2 = (so2_concentration*10) % 256
        send_data.append(dec_to_binary_hex_char(byteL_so2))

        sof2_concentration = predicted
        print(sof2_concentration)
        byteH_sof2 = (so2_concentration * 10) // 256
        send_data.append(dec_to_binary_hex_char(byteH_sof2))
        byteL_sof2 = (so2_concentration * 10) % 256
        send_data.append(dec_to_binary_hex_char(byteL_sof2))

        so2f2_concentration = 0
        byteH_so2f2 = (so2f2_concentration * 10) // 256
        send_data.append(dec_to_binary_hex_char(byteH_so2f2))
        byteL_so2f2 = (so2f2_concentration * 10) % 256
        send_data.append(dec_to_binary_hex_char(byteL_so2f2))

        self.lineEdit_show_SO2F2.setText('0')
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
        # 液晶显示数字处理
        erosion_predict_integer = int(erosion_predict)
        byteH_erosion_predict_integer = erosion_predict_integer // 256
        send_data.append(dec_to_binary_hex_char(byteH_erosion_predict_integer))
        byteL_erosion_predict_integer = erosion_predict_integer % 256
        send_data.append(dec_to_binary_hex_char(byteL_erosion_predict_integer))
        erosion_predict_decimal = int((erosion_predict - erosion_predict_integer) * 100)
        byte_erosion_predict_decimal = erosion_predict_decimal
        send_data.append(dec_to_binary_hex_char(byte_erosion_predict_decimal))
        print("此次开断烧蚀量为:{}".format(erosion_predict))
        self.lineEdit_show_erosion.setText(str(erosion_predict))

        # 计算到目前为止的烧蚀量
        self.erosion_accumulated.append(erosion_predict)
        erosion_predict_total = np.array(self.erosion_accumulated)
        erosion_predict_total = (sum(erosion_predict_total) * 1000 // 1) / 1000
        print("总烧蚀量为:{}".format(erosion_predict_total))
        # 液晶显示数字处理
        erosion_predict_total_integer = int(erosion_predict_total)
        byteH_erosion_predict_total_integer = erosion_predict_total_integer // 256
        send_data.append(dec_to_binary_hex_char(byteH_erosion_predict_total_integer))
        byteL_erosion_predict_total_integer = erosion_predict_total_integer % 256
        send_data.append(dec_to_binary_hex_char(byteL_erosion_predict_total_integer))
        erosion_predict_total_decimal = int((erosion_predict_total - erosion_predict_total_integer) * 100)
        byte_erosion_predict_total_decimal = erosion_predict_total_decimal
        send_data.append(dec_to_binary_hex_char(byte_erosion_predict_total_decimal))

        self.lineEdit_show_erosion_total.setText(str(erosion_predict_total))
        # 显示剩余寿命
        lifetime_remain_percent = (1 - erosion_predict_total / 20000) * 100
        self.progressBar_lifetime_remain.setValue(lifetime_remain_percent)
        # 液晶显示数字处理
        lifetime_integer = int(lifetime_remain_percent)
        byte_lifetime_integer = lifetime_integer
        send_data.append(dec_to_binary_hex_char(byte_lifetime_integer))
        lifetime_decimal = int((lifetime_remain_percent - lifetime_integer) * 100)
        byte_lifetime_decimal = lifetime_decimal
        send_data.append(dec_to_binary_hex_char(byte_lifetime_decimal))

        # 保存测得的烧蚀量
        with open('erosion_record.json', 'w', encoding='utf-8') as f:
            json.dump(self.erosion_accumulated, f, ensure_ascii=False)

        """向下位机发送数据"""
        try:
            self.com.write(self.command_show_result)
            hexData = ''.join(send_data)
            print(hexData)
            t = 0
            for i in range(10000):
                t = t + 1
            hexData = binascii.a2b_hex(hexData)
            print(hexData)
            self.com.write(hexData)
            for i in range(10):
                t = t + 1
        except:
            QMessageBox.critical(self, "Port Error", "与下位机通信失败，请检查串口！")
            self.timer.stop()
            return None
        # self.timer.start(self.time_interval)

    def quit(self):
        self.com.close()
        global app
        # app1 = QtWidgets.QApplication.instance()
        # app1.quit()
        app.quit()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = SerialSvmGui()
    # 禁止串口进行拉伸操作
    window.setFixedSize(window.width(), window.height())
    # window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)   设置之后只有最小化按钮
    # window.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)
    window.show()
    sys.exit(app.exec_())

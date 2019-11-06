# -*- coding:utf-8 -*-
"""
@author:Lisa
@file:svm_Iris.py
@func:Use SVM to achieve Iris flower classification
@time:2018/5/30 0030上午 9:58
"""
from sklearn import svm
import numpy as np
from sklearn.model_selection import train_test_split
import xlrd
import pickle


attributions = []
label = []
# 1.读取数据集
book = xlrd.open_workbook("阵列响应值.xlsx")  # 文件名，把文件与py文件放在同一目录下
sheet = book.sheets()[0]  # execl里面的worksheet1

col_0 = sheet.col_values(0, 0)
for i in range(1, sheet.ncols):
    attribution = sheet.col_values(i, 0)
    attributions.append(attribution)

for i in col_0:
    # temp = gas_concentration[i]
    label.append(i/5)

# 划分标签
label = np.array(label)
y = label.reshape(label.size, 1)
# 划分数据
attributions = attributions[:4]
attributions = np.array(attributions)
x = attributions.transpose()

# print(x)
# print(y)
# 划分训练样本与测试样本
train_data, test_data, train_label, test_label = train_test_split(x, y, random_state=1, train_size=0.7, test_size=0.3)

# print(train_data)
# sklearn.model_selection.
# 训练s vm分类器
classifier = svm.SVC(C=2, kernel='rbf', gamma=10, decision_function_shape='ovr')  # ovr:一对多策略
classifier.fit(train_data, train_label.ravel())  # ravel函数在降维时默认是行序优先
print("训练集：", classifier.score(train_data, train_label))
print("测试集：", classifier.score(test_data, test_label))

# 查看决策函数
# print('train_decision_function:\n', classifier.decision_function(train_data))  # (90,3)
# print('predict_result:\n', classifier.predict(train_data))
train_label = np.array(train_label)
train_label = train_label.transpose()
print(train_label)

s = pickle.dumps(classifier)
f = open('fitting.model', "wb+")
f.write(s)
f.close()
print("Done\n")



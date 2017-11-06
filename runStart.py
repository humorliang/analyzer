# -*- coding:utf-8 -*-
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog
from PyQt5 import QtCore
from PyQt5 import QtGui
from analysis1 import Ui_Form_1
from analysis2 import Ui_Form_2
from main import Ui_Form
from function import myFuntion
import sys


class myClass(QWidget, Ui_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.toolButton.clicked.connect(self.btn_open)
        self.toolButton_2.clicked.connect(self.btn_open1)
        self.toolButton_3.clicked.connect(self.btn_open2)

    def btn_open(self):
        _translate = QtCore.QCoreApplication.translate
        fileName1, filetype = QFileDialog.getOpenFileName(self, '选取文件', 'C:/', 'All Files (*);;(*.xlsx)')  # 文件扩展名过滤
        self.label_4.setText(_translate("Form", fileName1))

    def btn_open1(self):
        _translate = QtCore.QCoreApplication.translate
        fileName1, filetype = QFileDialog.getOpenFileName(self, '选取文件', 'C:/', 'All Files (*);;(*.xlsx)')  # 文件扩展名过滤

        self.label_5.setText(_translate("Form", fileName1))

    def btn_open2(self):
        _translate = QtCore.QCoreApplication.translate
        fileName1, filetype = QFileDialog.getOpenFileName(self, '选取文件', 'C:/', 'All Files (*);;(*.xlsx)')  # 文件扩展名过滤
        # print(fileName1, filetype)
        self.label_6.setText(_translate("Form", fileName1))

    def analyPccCI(self):
        # 获取pcc文件路径
        PccPos = self.label_6.text()
        # PccPos = "G:/homeWork/file/PCC_G50.xlsx"
        # 获取pcc文件的ID
        PccId = self.spinBox.value() - 1
        # PccId = 0

        # 获取ci文件路径
        CiPos = self.label_4.text()
        # CiPos = "G:/homeWork/file/CI.xlsx"
        # 获取ci文件的ID
        CiId = self.spinBox_2.value() - 1
        # CiId = 0
        # 获取ci文件的条件列
        CiOpt = self.spinBox_4.value() - 1
        # CiOpt = 1

        # 功能函数应用
        myFun = myFuntion()

        pccTotal = myFun.numSpaceFun(PccPos, PccId)  # PCC 的符合总数

        Ci_x, Ci_y = myFun.getFileFun(CiPos, CiId, CiOpt)  # ci处理后的list
        CiDict = myFun.optItemFun(Ci_x, Ci_y)  # 结果为字典
        ciTotal = list(CiDict.keys())  # ci的符合项

        PccCiRes = myFun.compFun(pccTotal, ciTotal)

        return len(pccTotal), len(ciTotal), PccCiRes[3],len(PccCiRes[0])

    def analyTarCI(self):
        # 获取CI文件路径
        CiPos = self.label_4.text()
        # 获取CI文件的ID
        CiId = self.spinBox_2.value() - 1
        # 获取条件列
        CiOpt = self.spinBox_4.value() - 1

        # 获取target文件的路径
        tarPos = self.label_5.text()
        # 获取target文件的ID
        tarId = self.spinBox_3.value() - 1
        # 获取target文件的条件
        tarOpt = self.spinBox_5.value() - 1

        myFun = myFuntion()
        Ci_x, Ci_y = myFun.getFileFun(CiPos, CiId, CiOpt)  # ci处理后的list
        CiDict = myFun.optItemFun(Ci_x, Ci_y)  # 结果为字典
        ciTotal = list(CiDict.keys())  # ci的符合项

        Tar_x, Tar_y = myFun.getFileFun(tarPos, tarId, tarOpt)
        TarDict = myFun.optCutFun(Tar_x, Tar_y)
        TarFitTotal, TarDf = myFun.comTarFun(TarDict)  # tar中的released项

        comRes = myFun.compFun(ciTotal, TarFitTotal)
        comRes2 = myFun.compFun(ciTotal, TarDf)

        rota = myFun.rotaFun(len(comRes[0]), len(ciTotal))

        return len(ciTotal), len(TarDict.keys()), len(comRes[0]), comRes[3], rota, len(comRes[0]), len(comRes2[0])

    def detailResFun(self):
        # 获取PCC文件路径
        PccPos = self.label_6.text()
        # 获取Pcc文件的ID
        PccId = self.spinBox.value() - 1

        # 获取CI文件路径
        CiPos = self.label_4.text()
        # 获取CI文件的ID
        CiId = self.spinBox_2.value() - 1
        # 获取条件列
        CiOpt = self.spinBox_4.value() - 1

        # 获取target文件的路径
        tarPos = self.label_5.text()
        # 获取target文件的ID
        tarId = self.spinBox_3.value() - 1
        # 获取target文件的条件
        tarOpt = self.spinBox_5.value() - 1

        # 功能类的构建
        myFun = myFuntion()

        pccTotal = myFun.numSpaceFun(PccPos, PccId)  # PCC 的符合总数

        Ci_x, Ci_y = myFun.getFileFun(CiPos, CiId, CiOpt)  # ci处理后的list
        CiDict = myFun.optItemFun(Ci_x, Ci_y)  # 结果为字典
        ciTotal = list(CiDict.keys())  # ci的符合项

        Tar_x, Tar_y = myFun.getFileFun(tarPos, tarId, tarOpt)
        TarDict = myFun.optCutFun(Tar_x, Tar_y)
        TarFitTotal, TarDf = myFun.comTarFun(TarDict)  # tar中的released项

        comRes = myFun.compFun(ciTotal, TarFitTotal)  # ci与tar比较
        comRes2 = myFun.compFun(ciTotal, TarDf)  # ci与tar比较Darft

        comResP_C = myFun.compFun(pccTotal, ciTotal)  # pcc与ci比较
        rota = myFun.rotaFun(len(comRes[0]), len(ciTotal))

        # 最终结果列表 少的 多的
        Res_List = [pccTotal, ciTotal, TarFitTotal, comResP_C[1], comResP_C[4], comRes[1], comRes[4],
                    comRes[0], comRes2[0],comResP_C[0],comRes[0]]
        myFun.writeFun(Res_List)


# 二级页面1 pcc
class secondWindows(QWidget, Ui_Form_1):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def show_w(self):
        res = w.analyPccCI()
        _translate = QtCore.QCoreApplication.translate
        self.label_6.setText(_translate("Form", str(res[0])))
        self.label_7.setText(_translate("Form", str(res[1])))
        self.label_9.setText(_translate("Form", str(res[2])))
        self.label_10.setText(_translate("Form", str(res[3])))
        self.show()

    def close_w(self):
        self.close()


# 二级页面2 ci
class secondWindows2(QWidget, Ui_Form_2):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

    def show_w(self):
        res = list(w.analyTarCI())
        _translate = QtCore.QCoreApplication.translate
        self.label_6.setText(_translate("Form", str(res[0])))
        self.label_7.setText(_translate("Form", str(res[1])))
        self.label_12.setText(_translate("Form", str(res[2])))
        self.label_9.setText(_translate("Form", str(res[3])))
        self.label_8.setText(_translate("Form", str(res[4])))
        self.label_15.setText(_translate("Form", str(res[5])))
        self.label_16.setText(_translate("Form", str(res[6])))
        self.show()

        print("22")

    def close_w(self):
        self.close()

        # 具体项目


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = myClass()
    w1 = secondWindows()
    w2 = secondWindows2()
    w.show()
    w.pushButton_2.clicked.connect(w1.show_w)
    w.pushButton_3.clicked.connect(w.detailResFun)
    w.pushButton.clicked.connect(w2.show_w)
    sys.exit(app.exec_())

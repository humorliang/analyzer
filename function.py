# -*- coding:utf-8 -*-
import re
import xlrd as xr
import xlwt as xw


class myFuntion():
    # 赛选出以数字开头和结尾的项，且不包含1217.2
    def numSpaceFun(self, xlsPos, colNum):
        # 打开文件进行过滤
        # 打开文件
        data = xr.open_workbook(xlsPos)
        # 打开表
        table = data.sheet_by_index(0)
        # 获取列值
        colValue = table.col_values(colNum)
        resList = []
        for x in range(len(colValue)):
            # 以数字开头和结尾的数
            if re.search(r'^[0-9].*[0-9]$', str(colValue[x])):
                if str(colValue[x]) != "1217.2":
                    resList.append(colValue[x])
        return resList

    # 条件筛选函数CI文件中的符合项
    def optItemFun(self, listBase, listTarget):
        # 字典集合
        optDict = {}
        for a in range(len(listBase)):
            if re.search(r'^[0-9].*[0-9]$', str(listBase[a])):
                if str(listBase[a]) != "1217.2" and str(listTarget[a]) != "T":
                    optDict[listBase[a]] = listTarget[a]
        return optDict

    # 获取文件列值得函数
    def getFileFun(self, xlsPos, colNum, colNum2):
        # 打开文件进行过滤
        # 打开文件
        data = xr.open_workbook(xlsPos)
        # 打开表
        table = data.sheet_by_index(0)
        # 获取列值
        colValue = table.col_values(colNum)
        colValue2 = table.col_values(colNum2)
        return list(colValue), list(colValue2)

    # 字符串切割函数
    def optCutFun(self, listBase, listOpt):
        optDict = {}
        for a in range(len(listBase)):
            x = listBase[a]
            if re.search(r'^[0-9]', str(x)):
                x = x[0:7]
                if re.search(r'^[0-9].*[0-9]$', str(x)):
                    if x.find('.') != -1:
                        if x != "1217.2":
                            optDict[x] = listOpt[a]
                else:
                    x = x[0:6]
                    if x.find('.') != -1:
                        if x != "1217.2":
                            optDict[x] = listOpt[a]
        return optDict

    # 数据比较函数
    def compFun(self, criterionList, comList):

        # 在基础里面的并且也在目标里的
        res = [a for a in criterionList if a in comList]

        # 基础里没做的
        resBase = [b for b in criterionList if b not in comList]

        # 多余的
        moreRes = [c for c in comList if c not in criterionList]
        resList = [res, resBase, len(res), len(resBase), moreRes]

        return resList

    # 目标数据比较函数 返回released项
    def comTarFun(self, tarDict):
        # released项
        tarList = []
        tarList2 = []
        for key, value in tarDict.items():
            if value == "released":
                tarList.append(key)
            elif value == "draft":
                tarList2.append(key)
        return tarList, tarList2

    # 百分比函数
    def rotaFun(self, tarNum, baseNum):
        a = ("%.2f%%" % (tarNum * 100 / baseNum))
        return a

    # 输出xls函数
    def writeFun(self, R_list):

        data = xw.Workbook()  # 创建工作簿
        sheet = data.add_sheet(u'result', cell_overwrite_ok=True)  # 创建表
        titleList = ["PCC", "CI", "PCC-CI", "CI-PCC", "PCC-CI-Fit", "CI", "Objects", "CI-Objects", "Objects-CI","Ob-CI-fit", "Released",
                     "Draft"]
        # 标题
        for m in range(12):
            sheet.write(0, m, titleList[m])
            sheet.col(int(m)).width = 256 * 15
        # 写内容
        # pcc
        for i in range(len(R_list[0])):
            sheet.write(i + 1, 0, R_list[0][i])
        # ci
        for i in range(len(R_list[1])):
            sheet.write(i + 1, 1, R_list[1][i])
        # pcc-ci
        for i in range(len(R_list[3])):
            sheet.write(i + 1, 2, R_list[3][i])
        # ci-pcc
        for i in range(len(R_list[4])):
            sheet.write(i + 1, 3, R_list[4][i])
        # pcc-ci-fit
        for i in range(len(R_list[9])):
            sheet.write(i + 1, 4, R_list[9][i])
        # ci
        for i in range(len(R_list[1])):
            sheet.write(i + 1, 5, R_list[1][i])
        # ob
        for i in range(len(R_list[2])):
            sheet.write(i + 1, 6, R_list[2][i])
        # ci-ob
        for i in range(len(R_list[5])):
            sheet.write(i + 1, 7, R_list[5][i])
        # ob-ci
        for i in range(len(R_list[6])):
            sheet.write(i + 1, 8, R_list[6][i])
        # ob-ci-fit
        for i in range(len(R_list[10])):
            sheet.write(i + 1, 9, R_list[10][i])
        # re
        for i in range(len(R_list[7])):
            sheet.write(i + 1, 10, R_list[7][i])
        # df
        for i in range(len(R_list[8])):
            sheet.write(i + 1, 11, R_list[8][i])

        data.save('detailResult.xls')


if __name__ == '__main__':
    xlsPos = "G:/homeWork/file/CI.xlsx"
    xlsPosB = "G:/homeWork/file/books.xlsx"
    xlsPosPcc = "G:/homeWork/file/PCC_G50.xlsx"
    myFun = myFuntion()
    x, y = myFun.getFileFun(xlsPos, 0, 1)
    m, n = myFun.getFileFun(xlsPosB, 0, 1)
    z = myFun.numSpaceFun(xlsPosPcc, 0)
    a = myFun.optItemFun(x, y)  # 字典CI数据处理项
    b = myFun.optCutFun(m, n)  # 字典tar数据处理项
    # c = myFun.compFun(list(a.keys()), list(b.keys()))
    d, e = myFun.comTarFun(b)  # released项
    h = myFun.rotaFun(1, 3)
    print(h)
    print(a)
    print(b)
    print(d)
    print(e)
    # print(c[2],len(c[0]),len(c[1]))

#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time    : 2023/4/13 18:05
# @Author  : YongKuo
# @Email   : YongKuoD@gmail.com
# @File    : RPA8.py
# @Project : rpa
# @Software: PyCharm


import os
import requests
import json
import pandas as pd


# excel 表格规范字段，和表格表头一一对应，顺序不可更改
excelKey = ['invoiceCode', 'invoiceNo', 'buyerName', 'buyerTaxNo', 'buyerAddressPhone', 'buyerBankAccount',
                    'buyerPhone', 'buyerEmail',
                    'sellerName', 'sellerTaxNo', 'sellerAddressPhone', 'sellerBankAccount', 'taxationMethod',
                    'deductibleAmount', 'invoiceType',
                    'invoiceType', 'payee', 'checker', 'drawer', 'remarks', 'deviceType', 'invoiceListMark', 'serialNo',
                    'goodsLineNo', 'invoiceLineNature',
                    'preferentialMarkFlag', 'goodsCode', 'goodsName', 'goodsTaxRate', 'goodsSpecification',
                    'goodsUnit', 'goodsQuantity', 'includTaxgoodsPrice', 'excludTaxgoodsPrice', 'priceTaxMark',
                    'includTaxgoodsTotalPrice',
                    'excludTaxgoodsTotalPrice', 'goodsTotalTax', 'invoiceTotalPriceTax', 'invoiceTotalPrice',
                    'invoiceTotalTax', 'invoiceTotalPriceTax', 'invoiceTerminalCode', 'invoiceDate',
                    'invoiceStatus', 'invoiceInvalidDate', 'orderNo', 'deliveryNo', 'sourceMark', 'invoiceCheckMark',
                    "machineCode", 'agencyCode', 'agencyName',
                    'agencyTaxNo', "playStatus", 'invoiceStatus', "className", "studentName"]

# 获取 excel 表格规范
def get_clomns():
    levels = [['发票代码','发票号码','购货单位名称','购货纳税人识别号','购货单位地址电话','购货单位银行账号','购方客户电话','购方客户邮箱','销货单位名称',
    '销货单位纳税识别号','销货单位地址电话','销货单位银行帐号','征收方式','差额征收扣除额','发票类型','开票类型','收款人','审核人','开票人','备注',
    '设备类型','清单标志','发票请求流水号','发票明细','合计金额（含税）','合计金额（不含税）','合计税额','价税合计','开票终端标识','开票日期',
    '发票状态','作废日期','机器编号','业务发票请求流水号','快递单号','来源标识','验签状态','机构代码','机构名称','机构税号','支付状态',
    '一体机开票状态','班级名称（备注）','教务系统学员姓名'],
    ['','发票行行号','发票行性质','优惠政策','商品编码','商品名称' ,'税率','规格型号','单位','数量','单价（含稅）','单价（不含稅）','含税标志','金额（含税）','金额（不含税）','税额']]

    codes=[[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,23,23,23,23,23,23,23,23,23,23,23,23,23,23,
                           24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43],
                          [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]]

    mulClomns = pd.MultiIndex(levels=levels,
                   codes=codes)

    return mulClomns



class RPA(object):
    # 初始化
    def __init__(self):
        self.get_config()
        self.data = self.get_data()
    def get_config(self):
        # print(os.getcwd())
        configFile = os.path.join(os.getcwd(),"配置文件.xlsx")
        dataframe = pd.read_excel(configFile, header=None)
        config = list(dataframe.loc[:, 1])
        self.targetDir = config[0]
        if not os.path.exists(self.targetDir):
            os.makedirs(self.targetDir)
        self.stime = config[1]
        self.etime = config[2]



    # 获取接口数据
    def get_data(self,):

        url = "http://60.205.245.225:31315/cms/web/invoiceAll?startTime=%s&endTime=%s"%(self.stime, self.etime)
        response = requests.get(url=url)
        content = json.loads(response.text)
        # print(content)
        return content


    def processing_issstData(self):
        '''
            isSST 数据格式
                | "outputInvoiceQuery" : outputkeys
                     | "invoiceDetailsList" : detailsListKeys

                | "invoiceInfo" : key2

                | 未含有的字段 : nokeys

        :return: DataFrame  issst的数据
        '''

        isSST = self.data["isSST"]
        # outputInvoiceQuery 字段 key 值
        outputkeys = ['invoiceCode','invoiceNo','buyerName','buyerTaxNo','buyerAddressPhone','buyerBankAccount','buyerPhone','buyerEmail',
                'sellerName','sellerTaxNo','sellerAddressPhone','sellerBankAccount','taxationMethod','deductibleAmount',
                'invoiceType','invoiceType','payee','checker','drawer','remarks','deviceType','invoiceListMark','serialNo',
                'invoiceTotalPriceTax', 'invoiceTotalPrice', 'invoiceTotalTax', 'invoiceTotalPriceTax', 'invoiceTerminalCode', 'invoiceDate',
                'invoiceInvalidDate', 'orderNo', 'sourceMark', 'invoiceCheckMark','invoiceStatus']

        # "invoiceDetailsList" 字段 key 值
        detailsListKeys = ['goodsLineNo','invoiceLineNature','preferentialMarkFlag','goodsCode','goodsName','goodsTaxRate','goodsSpecification',
                'goodsUnit','goodsQuantity','includTaxgoodsPrice','excludTaxgoodsPrice','priceTaxMark','includTaxgoodsTotalPrice',
                'excludTaxgoodsTotalPrice','goodsTotalTax']

        # invoiceInfo 字段 key 值
        invoiceInfoKeys = ['invoiceStatus',"machineCode","playStatus","className","studentName"]
        # 未包含 key 值
        noKey= ['deliveryNo','agencyCode','agencyName','agencyTaxNo']
        data = []
        for sst in isSST:
            datadict = {}
            outputInvoice = sst["outputInvoiceQuery"]
            detailsList = sst["outputInvoiceQuery"]["invoiceDetailsList"][0]
            invoiceInfo = sst["invoiceInfo"][0]
            for ko in outputkeys:
                datadict[ko] = outputInvoice[ko]
            datadict['invoiceStatus1'] = datadict['invoiceStatus']

            for kd in detailsListKeys:
                datadict[kd] = detailsList[kd]
            for ki in invoiceInfoKeys:
                datadict[ki] = invoiceInfo[ki]

            data.append(datadict)

        # 构建 dataframe
        dataFrame = pd.DataFrame(data)

        # 未包含的字段设为 ''
        for kn in noKey:
            dataFrame[kn] = ''


        return dataFrame


    def processing_notsstData(self):
        '''
            notSST 数据格式
                |
                | "outputInvoiceQuery" : outputkeys
                     | "invoiceDetailsList" : detailsListKeys

                | "invoiceInfo" : key2

                | 未含有的字段 : nokeys

        :return: DataFrame  notsst的数据
        '''

        notSST = self.data["notSST"]
        # 单挑数据直接暴露的 key 值
        baseKey = ['serialNo', 'invoiceCode', 'invoiceNo', 'deviceType', 'buyerTaxNo', 'buyerName', 'buyerAddressPhone',
                'buyerBankAccount',
                'sellerTaxNo', 'sellerName', 'sellerAddressPhone', 'sellerBankAccount', 'invoiceTotalPrice',
                'invoiceTotalTax', 'invoiceTotalPriceTax',
                'invoiceListMark', 'invoiceType', 'taxationMethod', 'invoiceDate', 'deductibleAmount', 'remarks',
                'drawer', 'checker', 'payee',
                'buyerEmail', 'buyerPhone', 'invoiceInvalidDate', 'invoiceStatus', 'invoiceCheckMark',
                'invoiceTerminalCode', 'sourceMark', 'orderNo']

        # invoiceDetailsList 字段 key 值
        detailsListKeys = ['goodsLineNo', 'invoiceLineNature', 'preferentialMarkFlag', 'goodsCode', 'goodsName', 'goodsTaxRate',
                'goodsSpecification',
                'goodsUnit', 'goodsQuantity', 'includTaxgoodsPrice', 'excludTaxgoodsPrice', 'priceTaxMark',
                'includTaxgoodsTotalPrice',
                'excludTaxgoodsTotalPrice', 'goodsTotalTax']
        # 未包含的 key 值
        noKeys = ['machineCode', 'playStatus', 'className', 'studentName','agencyTaxNo','agencyCode','agencyName','deliveryNo']

        data = []
        for one in notSST:
            dataDict = {}
            for kb in baseKey:
                dataDict[kb] = one[kb]
            if one['invoiceDetailsList']:
                DetailsList = one['invoiceDetailsList'][0]
                for kd in detailsListKeys:
                    dataDict[kd] = DetailsList[kd]
            else:
                for kd in detailsListKeys:
                    dataDict[kd] = ''
            data.append(dataDict)

        # 构建 dataframe
        dataFrame = pd.DataFrame(data)
        if self.dataframe_isempty(dataFrame):
            return
        dataFrame = dataFrame[dataFrame['invoiceStatus'] == "00"]
        dataFrame['invoiceStatus'] = "开具成功"
        for kn in noKeys:
            dataFrame[kn] = ''
        # 根据 excel 表格字段顺序输出
        dataFrame = dataFrame[excelKey]
        # 更改 dataframe clomns
        dataFrame.columns = get_clomns()

        return dataFrame



    def dataframe_isempty(self,dataframe):
        if isinstance(dataframe, pd.DataFrame):
            return dataframe.empty
        else:
            return True

    # 结合数据输出对应的 excel 表格
    def create(self):

        # 一体机已开票数据
        issstData = self.processing_issstData()
        invoicedData = None
        uninvoicedData = None
        if not self.dataframe_isempty(issstData):
            # A  一体机里面已开票数据 支付状态已支付 以班级名称命名
            # 支付状态 是    发票状态 已开 00

            issstData = issstData[issstData['invoiceStatus1'] == "00"]
            issstData = issstData[issstData['invoiceStatus'] == "已开"]
            issstData.loc[issstData['playStatus'] == '是', 'playStatus'] = "已支付"
            issstData.loc[issstData['playStatus'] == '支付', 'playStatus'] = "已支付"
            # issstData[issstData['playStatus'] == "是"]["playStatus"] = "已支付"
            invoicedData = issstData[issstData['playStatus'] == "已支付"]
            invoicedData = invoicedData[excelKey]
            invoicedData.columns = get_clomns()
            invoicedDir = os.path.join(self.targetDir,"A")
            if not os.path.exists(invoicedDir):
                os.makedirs(invoicedDir)
            className = invoicedData["班级名称（备注）"]
            for name in list(className):

                idataClass = invoicedData[invoicedData["班级名称（备注）"] == name]


                fileName = os.path.join(invoicedDir,name+".xlsx")
                idataClass.index = list(range(1, idataClass.shape[0] + 1))
                idataClass.to_excel(fileName,index_label="序号",)
            # B 一体机里面已开票数据 支付状态未支付 以班级名称命名  支付状态字段 playStatus
            uninvoicedData =  issstData[issstData['playStatus'] != "已支付"]

            uninvoicedData = uninvoicedData[excelKey]
            uninvoicedData.columns = get_clomns()
            uninvoicedDir= os.path.join(self.targetDir,"B")
            if not os.path.exists(uninvoicedDir):
                os.makedirs(uninvoicedDir)
            className = uninvoicedData["班级名称（备注）"]
            for name in list(className):
                udataClass = uninvoicedData[uninvoicedData["班级名称（备注）"] == name]
                fileName = os.path.join(uninvoicedDir,name+".xlsx")
                udataClass.index = list(range(1, udataClass.shape[0] + 1))
                udataClass.to_excel(fileName)
        # print(className)
        # C 非一体机的数据，   表名（特殊开票明细））

        notsstData = self.processing_notsstData()



        if  not self.dataframe_isempty(notsstData):
            notsstDir = os.path.join(self.targetDir,"C")
            if not os.path.exists(notsstDir):
                os.makedirs(notsstDir)
            notsstFile = os.path.join(self.targetDir,"C","特殊开票明细.xlsx")
            notsstData.index = list(range(1, notsstData.shape[0] + 1))
            notsstData.to_excel(notsstFile,index_label="序号")

        # 总表

        allFile = os.path.join(self.targetDir, "总表.xlsx")
        allData = None
        count = 0
        for data in [invoicedData , uninvoicedData , notsstData]:
            if ((not self.dataframe_isempty(data) )and (count == 0)):
                allData = data
            if ((not self.dataframe_isempty(data) )and count >0):
                allData = pd.concat([allData, data], ignore_index=True)
        # print(allData)
        if not self.dataframe_isempty(allData):
            allData.index = list(range(1, allData.shape[0] + 1))
            allData.to_excel(allFile, index_label="序号")





if __name__ == "__main__":
    a = RPA()
    a.create()


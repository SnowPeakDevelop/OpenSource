# -*- coding:utf-8 -*-
 
import xdrlib, sys
import xlrd
import os
import json
import codecs

xpath="../excel"
cpath="outputJson/"
xtype="xlsx"
allTable = []
file_path=[]


#取得列表中所有的type文件
def collect_xls(list_collect,type1):
    for each_element in list_collect:
        if isinstance(each_element,list):
            collect_xls(each_element,type1)
        elif each_element.endswith(type1):
              allTable.insert(0,each_element)
    return allTable
#读取所有文件夹中的xls文件
def read_xls(path,type2):
    #遍历路径文件夹
    for file in os.walk(path):
        name = []
        for each_list in file[2]:
            file_path=file[0]+"/"+each_list
            #os.walk()函数返回三个参数：路径，子文件夹，路径下的文件，利用字符串拼接file[0]和file[2]得到文件的路径
            name.insert(0,file_path)
        all_xls = collect_xls(name, type2)
    #遍历所有type文件路径并读取数据
    for evey_name in all_xls:
        xls_data = open_excel(evey_name)
         # 存储Excel中的所有数据
        excel_data = {}
        for each_sheet in xls_data.sheets():
            excel_sheet_data = run(each_sheet)
             # 存传入表的所有数据
            excel_data[each_sheet.name] = excel_sheet_data
        # 将excel_data数据转化成json字符串，indent代表缩进为4, ensure_ascii = false 代表汉字直接可以显示出来
        data = json.dumps(excel_data, ensure_ascii=False ,indent=4)
        base_name = os.path.basename(evey_name)[0:-5] # 取出路径中的文件名字，删除后缀
        # 拼接生成文件路径, 写入一个自定义的路径中
        file_name = os.path.join(cpath,base_name)
        # 将json字符串写入ts文件
        # with codecs.open(file_name + '.ts', 'w', 'utf-8') as fir:
        #     fir.write("export default class " + base_name  + data)
        #     print(file_name + ".ts to successful")
        # 将json字符串写入json文件
        with codecs.open(file_name + '.json', 'w', 'utf-8') as fir:
            fir.write(data)
            print(file_name + ".json to successful")


#打开excel文件
def open_excel(file):
    data = xlrd.open_workbook(file)
    return data
 
def run(sheet):
    sheet_data = {}
    # 第一列：id
    ids = sheet.col_values(0)
    # 第一行：键
    keys = sheet.row_values(0)
    # 行长度
    row_len = sheet.row_len(0)
 
    # 遍历行
    for row_num in range(1, sheet.nrows):
        _id = int(ids[row_num])
        sheet_data[_id] = {"Id": _id}
        # 遍历列
        for col_num in range(1, row_len):
            key = keys[col_num]
            value = sheet.row_values(row_num)[col_num]
            ctype = sheet.cell(row_num,col_num).ctype
            if ctype == 2 and value % 1 == 0.0:
                value = int(value)
            # 存row_num行col_num列对应的值
            sheet_data[_id][key] = value
    return sheet_data
 
def main():
    read_xls(xpath, xtype)
if __name__=="__main__":
    main()
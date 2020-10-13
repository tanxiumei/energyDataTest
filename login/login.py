import json
import hashlib
# from copy import copy
from xlutils.copy import copy

import requests
from datetime import datetime
import xlwt
import xlrd
from xlrd import open_workbook

Nowtime = datetime.now().strftime("%Y%m%d%H%M%S")

#定义一个全局变量，在写入Excel中使用
write_row_num = 1


# 登录接口
def loginCode():
    url = "http://10.100.101.198:8088/ebx-rook/oauth/authverify2.as"
    params = {
        "response_type": "code",
        "client_id": "O000000063",
        "redirect_uri": "http://10.100.101.198:8088/ebx-rook/demo.jsp",
        "uname": "prod",
        "passwd": "abc123++"
    }
    search_header = {
        "Content-Type": "application/json"
    }
    try:
        result = requests.post(url, data=json.dumps(params), headers=search_header)

        lg_token = result.json()

        # 获取登录的code值
        access_code = lg_token["code"]
        # print("code的值：" + access_code)

        # 获取md5加密的sign值
        # param = appKey + grant_type + redirect_uri + code + appSecret;
        appKey = "O000000063"
        grant_type = "authorization_code"
        redirect_uri = "http://10.100.101.198:8088/ebx-rook/demo.jsp"
        code = access_code
        appSecret = "590752705B63B2DADD84050303C09ECF"
        md5_param = appKey + grant_type + redirect_uri + code + appSecret

        h1 = hashlib.md5()

        # Tips
        # 此处必须声明encode
        h1.update(md5_param.encode(encoding='utf-8'))

        # print('MD5加密前为 ：' + md5_param)
        # print('MD5加密后为 ：' + h1.hexdigest())
        sign = h1.hexdigest()
        return code, sign


    except IndentationError as e:
        print("nothing111")


def access_login(code, sign, macName):
    url = "http://10.100.101.198:8088/ebx-rook/oauth/token.as"
    params = {
        "client_id": "O000000063",
        "client_secret": sign,
        "grant_type": "authorization_code",
        "redirect_uri": "http://10.100.101.198:8088/ebx-rook/demo.jsp",
        "code": code
    }
    search_header = {
        "Content-Type": "application/json"
    }
    try:
        result = requests.post(url, data=json.dumps(params), headers=search_header)

        lg_token = result.json()

        # 获取登录的code值  accessToken
        access_token = lg_token["data"]["accessToken"]

        print("access_token的值11：" + access_token)

        # 获取md5加密的sign值
        # param = appKey + grant_type + redirect_uri + code + appSecret;
        access_token = access_token
        client_id = "O000000063"
        mac = macName
        method = "GET_BOX_DAY_POWER"
        projectCode = "P00000000001"
        timestamp = Nowtime
        appSecret = "590752705B63B2DADD84050303C09ECF"
        year = "2020"
        month = "01"
        day = "16"
        appSecret = "590752705B63B2DADD84050303C09ECF"
        md5_param2 = access_token + client_id + day + mac + method + month + projectCode + timestamp + year + appSecret
        h1 = hashlib.md5()
        # 此处必须声明encode
        h1.update(md5_param2.encode(encoding='utf-8'))
        sign = h1.hexdigest()
        print("sign2的值：" + sign)

        # 获取实时状态数据使用的sign值：
        get_now_method = "GET_BOX_CHANNELS_REALTIME"
        # param = access_token + client_id + mac + method + projectCode + timestamp + appSecret
        get_now_param = access_token + client_id + mac + get_now_method + projectCode + timestamp + appSecret
        h2 = hashlib.md5()
        # 此处必须声明encode
        h2.update(get_now_param.encode(encoding='utf-8'))
        get_now_energy_sign = h2.hexdigest()
        print("get_now_energy_sign的值：" + get_now_energy_sign)

        return access_token, sign, get_now_energy_sign


    except IndentationError as e:
        print("nothing2222")


# 将列表写进Excel
# file_name:自定义文件名
# data_list:列表
# sheet_name:工作表名称（有默认值）
def write_to_excel(file_name, data_list, listname, row=0, col=0, sheet_name="5floor"):
    book = xlwt.Workbook()
    sheet = book.add_sheet(sheet_name)
    rowNum = row
    templist = listname
    for data_unit in data_list:
        colNum = col
        for data in templist:
            sheet.write(rowNum, colNum, data_unit[data])
            colNum += 1
        rowNum += 1
        book.save(file_name)
def write_to_excel_naxin(file_name, data_list, listname, row=0, col=0, sheet_name="5floor"):
    book = xlwt.Workbook()
    sheet = book.add_sheet(sheet_name)
    rowNum = row
    templist = listname
    for data_unit in data_list:
        colNum = col
        for data in templist:
            sheet.write(rowNum, colNum, data_unit[data])
            colNum += 1
        rowNum += 1
        book.save(file_name)

def write_now_to_excel(file_name, data_list, listname, row=0, col=1, sheet_name="5floor"):
    book = xlwt.Workbook()
    sheet = book.add_sheet(sheet_name)
    rowNum = row
    name_templist = listname
    print(name_templist)
    i = 0
    for data_unit in data_list:
        colNum = col
        templist = data_unit
        sheet.write(rowNum, colNum, name_templist[i])
        i += 1
        colNum+=1
        for data in templist:
            sheet.write(rowNum, colNum, data)
            colNum += 1

        rowNum += 1

        book.save(file_name)

def write_to_oldexcel(file_name, data_list, listname, row=0, col=7, sheet_name="5floor"):
    oldbook = xlrd.open_workbook(file_name)
    sheet = oldbook.sheet_by_index(0)
    global write_row_num
    rowNum = write_row_num
    # colNum = sheet.ncols

    oldWb = xlrd.open_workbook(file_name);  # 先打开已存在的表
    newWb = copy(oldWb)  # 复制
    newWs = newWb.get_sheet(0);  # 取sheet表

    templist = listname
    print(templist)
    print("纳新科技！")
    print(len(data_list))
    print(data_list)
    for data_unit in data_list:
        colNum = col
        for data in templist:
            newWs.write(rowNum, colNum, data_unit[data])
            colNum += 1
        rowNum += 1
        col = 7
        newWb.save(file_name)

    write_row_num = len(data_list) + write_row_num
    # newWs.write(2, 4, "pass");  # 写入 2行4列写入pass
    # newWb.save(file_name);  # 保存至result路径


def get_day_energy(access_token, sign, mac):
    url = "http://10.100.101.198:8088/ebx-rook/invoke/router.as"
    params = {
        "client_id": "O000000063",
        "method": "GET_BOX_DAY_POWER",
        "access_token": access_token,
        "timestamp": Nowtime,
        "sign": sign,
        "projectCode": "P00000000001",
        "mac": mac,
        "year": "2020",
        "month": "01",
        "day": "16"
    }
    search_header = {
        "Content-Type": "application/json"
    }
    try:
        result = requests.post(url, data=json.dumps(params), headers=search_header)

        result_data = result.json()
        print("曼顿result_data的值")
        print(type(result_data))
        print(result_data)
        # 获取登录的code值  accessToken
        data_addr = result_data["data"]
        print("曼顿data_addr的值")
        print(type(data_addr))
        print(data_addr)
        # write_to_excel("data.xls",data_addr)
        mandun_list = {"addr", "electricity"}
        write_to_excel("data.xls", data_addr, mandun_list, 1, 1, "5floor")
    except IndentationError as e:
        print("nothing3333")


def get_now_energy(access_token, sign, mac):
    # for mac in mac_list:
    test_mac = mac
    print("test_mac的值：" + test_mac)
    url = "http://10.100.101.198:8088/ebx-rook/invoke/router.as"
    params = {
        "client_id": "O000000063",
        "method": "GET_BOX_CHANNELS_REALTIME",
        "access_token": access_token,
        "timestamp": Nowtime,
        "sign": sign,
        "projectCode": "P00000000001",
        "mac": test_mac
    }
    search_header = {
        "Content-Type": "application/json"
    }
    try:
        print("test!")
        result = requests.post(url, data=json.dumps(params), headers=search_header)
        result_data = result.json()
        print(result_data)
        now_data = result_data["data"]
        print("456")
        print(type(result_data))
        # 获取登录的code值  accessToken
        #        data_addr = result_data["data"]
        # print(type(data_addr))
        # write_to_excel("data.xls",data_addr)
        now_data_list = ["mac", "addr", "aW"]
        write_to_oldexcel("now_energy_data.xls", now_data, now_data_list)
    except NameError as e:
        print("nothing555555")


def get_day_energy_naxin():
    url = "http://10.100.101.199:9797/api/dashboard/electric_consume_day"
    params = {
        "start": "2020-01-07T00:00:00",
        "end": "2020-01-08T00:00:00",
        "floor": 5
    }
    search_header = {
        "Content-Type": "application/json"
    }
    try:
        result = requests.get(url, data=json.dumps(params), headers=search_header)

        result_data = result.json()
        print("纳新result_data")
        print(type(result_data))
        print(result_data)
        # 获取登录的code值  accessToken
        data_list = result_data[0]
        ele_data = data_list["2020-01-07"]["detail"]
        print("纳新ele_data的值")
        print(type(ele_data))
        print(ele_data)
        temp_list = []
        for data in ele_data:
            temp_list.append(ele_data[data])

        print("ele_data的值")
        print(type(temp_list))
        print(temp_list)
        naxin_list = ["light", "ac", "socket"]
        # naxin_list = {"electricity",}
        # write_to_excel(file_name, data_list, listname, row=0, col=0, sheet_name="5floor"):
        write_to_excel_naxin("data_naxin.xls", temp_list, naxin_list, 1, 1, "5floor")
    except IndentationError as e:
        print("nothing4444")

def get_now_energy_naxin():
    url = "http://10.100.101.199:9797/api/dashboard/energyshow"
    params = {
        "floor": 3
    }
    search_header = {
        "Content-Type": "application/json"
    }
    try:
        result = requests.get(url, data=json.dumps(params), headers=search_header)

        result_data = result.json()
        print(type(result_data))
        print(result_data)
        # 获取登录的code值  accessToken
        temp_list = []
        name_list = []
        for data in result_data:
            temp_list.append(result_data[data])
            name_list.append(data)
        print(type(temp_list))
        print(temp_list)
        write_now_to_excel("data_naxin_now.xls", temp_list, name_list, 1, 1, "5floor")
    except IndentationError as e:
        print("nothing666666")

if __name__ == '__main__':

    mac_list_5floor = ["187ED5321B70", "187ED532272C", "187ED5320D90", "187ED53216A8", "187ED53219CC"]
    #mac_list_3floor = ["187ED531DCF0","187ED5320880","187ED53224A4","187ED531F6F4","187ED53211C0"]
    #mac_list_3floor = ["187ED531D860"]
    for macTemp in mac_list_5floor:
        listdata = loginCode()
        code = listdata[0]
        sign = listdata[1]
        print(code)
        print(sign)
        mac = macTemp
        loginListData = access_login(code, sign, mac)
        accessToken = loginListData[0]
        accessSign = loginListData[1]
        get_now_energy_accessSign = loginListData[2]
        #print("accessToken的值:" + accessToken)
        #print("accessSign的值:" + accessSign)
        #print("get_now_energy_accessSign的值:" + get_now_energy_accessSign)
        #获取曼顿某日功率的函数
        get_day_energy(accessToken, accessSign, mac)
        #获取纳新某日功率的函数
        #get_day_energy_naxin()


        #获取曼顿接口的实时功率函数
        get_now_energy(accessToken, get_now_energy_accessSign, mac)
#获取纳新接口的实时功能函数
    get_now_energy_naxin()
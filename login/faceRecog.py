import json
from copy import copy

import requests
import xlrd
import xlwt


def loginAccess():
    url = "http://10.100.101.199:9797/api/login"
    params = {
	"username":"hxh",
	"password":"123456"
}
    search_header = {
        "Content-Type": "application/json"
    }
    try:
        result = requests.post(url, data=json.dumps(params), headers=search_header)
        lg_token = result.json()
        # 获取登录的code值
        access_code = lg_token["access_token"]
        #print("access_code的值：" + access_code)
        refresh_token = lg_token["refresh_token"]
        #print("refresh_token的值：" + refresh_token)
        return access_code
    except IndentationError as e:
        print("nothing111")

def write_now_to_excel(file_name, data_list, listname, row=2, col=1, sheet_name="5floor"):
    book = xlwt.Workbook()
    sheet = book.add_sheet(sheet_name)
    rowNum = 2
    colNum = 1
    for data_unit in data_list:

        templist = data_unit
        for k in templist:
            sheet.write(rowNum,colNum,templist[k])
            colNum += 1
        colNum = 1
        rowNum += 1

        book.save(file_name)

def faceRecog(accessToken):
    print("获取人脸识别数据"+accessToken)
    url = "http://10.100.101.199:9797/api/dashboard/faceRecog"
    token = "Bearer " + accessToken
    params = {
        "start":"2020-01-15T00:00:00",
        "end":"2020-01-15T18:25:11"
    }
    search_header = {
        "Authorization":token
    }
    try:
        print(123123)

        print(token)
        result = requests.get(url, params=params, headers=search_header)
        msg = result.json()
        write_now_to_excel("data_naxin_now.xls", msg,1, 1, "5floor")
        return msg
    except IndentationError as e:
        print("nothing111")

if __name__ == '__main__':
    accessToken = loginAccess()
    print(accessToken)
    faceRecog(accessToken)


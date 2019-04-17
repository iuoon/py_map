import requests
import time

if __name__ == "__main__":

    for line in open("card.txt","r",encoding="utf-8"): #设置文件对象并读取每一行文件
        arr=line.split("#")
        print(arr[0])
        print(arr[1])
        header = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
                  'Accept-Encoding': 'gzip, deflate',
                  'Accept-Language': 'zh-CN,zh;q=0.9',
                  'Cache-Control': 'max-age=0',
                  'Connection': 'keep-alive',
                  'Host': 't.dianping.com',
                  'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36',
                  'Cookie': arr[2].replace("\n","")
                  }
        paramData={
            "pkId":"27694",
            "cardNo":arr[1],
            "entityName":"Administrator"
        }
        ret = requests.post('http://222.82.237.75:8888/E-office/sit/getEmployeeByIdCard', data = paramData, timeout=30, headers=header)
        data = ret.json()
        if data['status']:
            print(data['data'][0][5])
            employee=data['data'][0][5]
            print(employee['idCardPhoto'])
            print(str(employee['value']))
            pa = {
                "pkId":"27694",
                "idCard":arr[1],
                "idCardPhoto":str(employee['idCardPhoto']),
                "idCardName":arr[0],
                "employeeId":str(employee['value']),
                "currentStep":"xcglryInfo",
                "confirmStatus":"Y",
                "sync":"false"
            }
            obj = requests.post('http://222.82.237.75:8888/E-office/sit/employeeIdConfirm', data = pa, timeout=30, headers=header)
            print(obj.text)
            time.sleep(10)

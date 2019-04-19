import requests
import time

if __name__ == "__main__":

    for line in open("card.txt","r",encoding="utf-8"): #设置文件对象并读取每一行文件
        arr=line.split("#")
        print(arr[0])
        print(arr[1])
        host=arr[3].replace("\n","")
        header = { 'Accept': 'application/json, text/javascript, */*; q=0.01',
                   'Referer': host+'E-office/sit/employeeReadCard',
                   'x-requested-with': 'XMLHttpRequest',
                  'Accept-Encoding': 'gzip, deflate',
                  'Accept-Language': 'zh-CN',
                  'Cache-Control': 'max-age=0',
                  'Connection': 'keep-alive',
                  'Accept-Encoding': 'gzip, deflate',
                  'Host': '222.82.237.75:8070',
                  'Pragma': 'no-cache',
                  'Content-Type': 'application/x-www-form-urlencoded',
                  'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; InfoPath.3; .NET4.0C)',
                  'Cookie': arr[2].replace("\n","")
                  }
        paramData={
            "licenseNum":'91650100397711723R',
            "cardNo":str(arr[1]),
            "employeeName":str(arr[0]),
            "entityName":'Administrator'
        }
        ret='{"status":False}'
        while True:
            try:
                ret = requests.post(host+'E-office/sit/getEmployeeByIdCard', data = paramData, timeout=30, headers=header)
                print(ret.text)
                break
            except requests.exceptions.ConnectionError:
                print('[ERROR]ConnectionError -- will retry connect')
                time.sleep(3)
        data = ret.json()
        if data['status']:
            print(data['data'][0][4])
            employee=data['data'][0][4]
            print(employee['idCardPhoto'])
            print(str(employee['value']))
            pa = {
                "licenseNum":"91650100397711723R",
                "idCard":arr[1],
                "idCardPhoto":str(employee['idCardPhoto']),
                "idCardName":arr[0],
                "employeeId":str(employee['value']),
                "currentStep":"xcglryInfo",
                "confirmStatus":"Y",
                "sync":"false"
            }
            while True:
                try:
                    obj = requests.post(host+'E-office/sit/employeeIdConfirm', data = pa, timeout=30, headers=header)
                    print(obj.text)
                    break
                except requests.exceptions.ConnectionError:
                    print('[ERROR]ConnectionError -- will retry connect')
                    time.sleep(3)
            time.sleep(10)

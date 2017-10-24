import random

import requests
import xlrd
import xlwt          #测试环境导入合同模版表里批量造数据，帐号（55555000000）

name = 'sdsdsds'

print(name)


book=xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet=book.add_sheet('test',cell_overwrite_ok=True)

book.save('dhs5.xlsx')



#获取客户名称
def auto():
    url='https://test01.weibangong.me/api/customer/web/filter/my'
    headers={
        'Content-Tyoe':'application/json;charset=UTF-8',
        'Accept': 'application/json, text/plain, */*',
        'Authorization':'Bearer eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiI3NzU5NTkxNDk0MzMyODk3Iiwic3ViIjoiNzc1OTU4IiwiYWlkIjo3OTgwNjIsInRpZCI6Nzk4MDYxLCJpYXQiOjB9.TxbZYnU_11lN17MuSGki3bJ2FCih6lUdq9bNIFyMvoQ',
        'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36'
    }
    data={
        "category": "OWNED",
        "orderField": "updatedAt",
        "orderDirection": "DESC",
        "offset": 0,
        "limit": 20,
        "scope": [],
        "item": ""
    }
    resp1=requests.post(url,json=data,headers=headers)
    print(resp1.status_code)
    total=resp1.json()['total']
    print(total)
    if total<20:
        random1=random.randint(0,total-1)     #客户数量少于20个时
    else:
        random1=random.randint(0,19)          #客户数量多于20个时

    print(random1)
    customer=resp1.json()['items'][random1]['name']
    customerId=resp1.json()['items'][random1]['id']
    print(customer)                         #获取客户名称和客户id


    #获取客户下第一个机会的名称

    url='https://test02.weibangong.me/api/opportunity/feed/customer/'+str(customerId)+'?limit=999&offset=0'
    print(url)
    headers={
        'Content-Tyoe':'application/json;charset=UTF-8',
        'Authorization':'Bearer eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiI3NzU5NTkxNDk0MzMyODk3Iiwic3ViIjoiNzc1OTU4IiwiYWlkIjo3OTgwNjIsInRpZCI6Nzk4MDYxLCJpYXQiOjB9.TxbZYnU_11lN17MuSGki3bJ2FCih6lUdq9bNIFyMvoQ',
        'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'
    }
    resp=requests.get(url,headers=headers)

    print(resp.json())
    items=resp.json()['items']

    if items==[]:
        opportunity=''         #客户下没有机会时不填写机会
    else:
        opportunity=resp.json()['items'][0]['name']
    #opportunity=opportunityName['items'][0]['name']
    print('机会是'+opportunity)

    #获取我方签约人
    url='https://test01.admin.weibangong.me/api/security/employee/manager_v2'
    headers={
        'Content-Tyoe':'application/json;charset=UTF-8',
        'Authorization':'Bearer eyJhbGciOiJIUzI1NiJ9.eyJqdGkiOiI3NzU5NTkxNDk0MzMyODk3Iiwic3ViIjoiNzc1OTU4IiwiYWlkIjo3OTgwNjIsInRpZCI6Nzk4MDYxLCJpYXQiOjB9.TxbZYnU_11lN17MuSGki3bJ2FCih6lUdq9bNIFyMvoQ',
        'User-Agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'
    }
    data={
        "viewRoles": [
            1,
            12
        ],
        "pageVo": {
            "offset": 0,
            "limit": 20,
            "query": ""
        },
        "status": []
    }
    resp2=requests.post(url,json=data,headers=headers,verify=False)
    print(resp2.status_code)
    random3=random.randint(0,19)
    datePerson=resp2.json()['items'][random3]['fullname']+'#'+resp2.json()['items'][random3]['mobile']
    print(datePerson)


    contractTitle='导入合同标题'+str(random.randint(0,1000))
    discount=random.randint(1,10)
    contractNum='合同编号'+str(random.randint(0,1000))
    contractTotal=random.randint(1000,100000)
    return contractTitle,contractTotal,discount,contractNum,customer,opportunity,datePerson



#读取合同导入模版
data=xlrd.open_workbook(u'合同导入模板 (4).xlsx')
table=data.sheet_by_name(u'合同')
title=table.row_values(0)

#选择自定义模版内容还是自动生成数据
print('合同模版数据自定义还是使用自动生成数据？自定义请输入1，自动生成请输入2')
custom=input()
if '2' in custom:

    print(title)
    lenght = len(title)
    print(lenght)
    for i in range(1, 15, 1):
        contractTitle,contractTotal,discount,contractNum,customer,opportunity,datePerson=auto()
        for l in range(0, lenght, 1):
            print(title[l])
            if '折扣' in title[l]:
                sheet.write(i, l, discount)
                print(1)
            elif '金额' in title[l]:
                sheet.write(i, l, contractTotal)
            elif '数字' in title[l]:
                sheet.write(i, l, '123')
            elif '开始日期' in title[l]:
                sheet.write(i, l, '2016-01-01')
            elif '结束日期' in title[l]:
                sheet.write(i, l, '2016-01-01')
                print(2)
            elif '单选' in title[l]:
                sheet.write(i, l, '合同单选一')
            elif '多选' in title[l]:
                sheet.write(i, l, '合同多选一')
            elif '我方签约人' in title[l]:
                sheet.write(i, l, datePerson)
            elif '客户' in title[l]:
                sheet.write(i, l, customer)
            elif '机会' in title[l]:
                sheet.write(i, l, opportunity)
            elif '付款方式' in title[l]:
                sheet.write(i, l, '银行转帐')
            else:
                sheet.write(i, l, contractTitle)


else:
    # 依次让用户输入合同模版中的参数
    print('请输入合同标题')
    contractTitle = input()
    print('请输入合同折扣')
    discount = input()
    print('请输入合同编号')
    contractNum=input()
    print('请输入客户名称')
    customer=input()
    print('请输入机会名称')
    opportunity=input()
    print('请输入合同金额')
    contractTotal=input()
    print('请输入签约日期')
    contractDate=input()
    print('请输入付款方式：')
    payTpye=input()
    print('请输入我方签约人(格式为姓名#帐号)')
    datePerson=input()
    print('请输入开始日期')
    startDate=input()
    print('请输入结束日期')
    endDate=input()


    #输入造数据数量
    print('请输入您想要导入合同的条数')
    num=input()
    Num=int(num)+1


    #读取模版所有字段
    print(title)
    lenght=len(title)
    print(lenght)


    for l in range(0,lenght,1):
        for i in range(1,Num,1):
            print(title[l])
            if '折扣' in title[l]:
                sheet.write(i,l,discount)
                print(1)
            elif '金额' in title[l]:
                sheet.write(i,l,contractTotal)
            elif '数字' in title[l]:
                sheet.write(i,l,'123')
            elif '开始日期' in title[l]:
                sheet.write(i,l,startDate)
            elif '结束日期' in title[l]:
                sheet.write(i,l,endDate)
                print(2)
            elif '单选' in title[l]:
                sheet.write(i,l,'合同单选一')
            elif '多选' in title[l]:
                sheet.write(i,l,'合同多选一')
            elif '我方签约人' in title[l]:
                sheet.write(i,l,'侯方域5#55500000001')
            elif '客户' in title[l]:
                sheet.write(i,l,customer)
            elif '机会' in title[l]:
                sheet.write(i,l,opportunity)
            elif '付款方式' in title[l]:
                sheet.write(i,l,payTpye)
            else:
                sheet.write(i,l,contractTitle)

book.save(u'合同导入模板 (4).xlsx')
print('给自己一个爱的抱抱')








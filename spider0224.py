#coding:utf-8

from selenium import webdriver
import time
import xlwt
import sys
import re

def search(driver,i,worksheet,mytxt):
    if(i==1):
        driver.get("http://m.54114.cn/hangzhou/")
        #mytxt = u"杭州海康威视数字技术股份有限公司"
        driver.find_element_by_xpath("/html/body/div[2]/div/form/div/div/input[2]").send_keys(mytxt)
        driver.find_element_by_xpath("//*[@id=\"qixc\"]").click()
    
    else:
        driver.get("http://m.54114.cn/hangzhou/")
        driver.find_element_by_xpath("/html/body/div[2]/div/form/div/div/input[2]").send_keys(mytxt)
        driver.find_element_by_xpath("//*[@id=\"qixc\"]").click()
    
    try:
        incname=driver.find_element_by_xpath("/html/body/div[3]/div[3]/ul/li[1]/a").text#.replace('\n','').replace('<font color="red">' ,'').replace('</font>','')
    except:
        worksheet.write(i,1,label=mytxt)
        worksheet.write(i,2,label='null')
        worksheet.write(i,3,label='null')
        worksheet.write(i,4,label='null')
        worksheet.write(i,5,label='null')
        worksheet.write(i,6,label='null')
        print(mytxt)
        print('未检索到该公司的信息\n')
        return
    
    if(incname==mytxt):
        print(incname)
        worksheet.write(i,1,label=incname)
    else:
        print('not this incname!')
    try:
        realweb = driver.find_element_by_xpath("/html/body/div[3]/div[3]/ul/li/a").get_attribute("href")
        driver.get(realweb)
    except:
        print("null")

    
    
    try:
        phone = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[2]/span/font/a").text
        m = re.findall(r'\(?0\d{2,3}[)-]?\d{7,8}',phone)
        if m:
            print(m[0])
            worksheet.write(i,2,label=m[0])
    except:
        print("phone null")
        
    
    try:
        mail = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[3]/span").text
        if(mail.find(u'邮箱')!=-1 ):
            mails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", mail)
            print(mails[0])
            worksheet.write(i,3,label=mails[0])
    except:
        print("mail null")

    try:
        url_or_addr = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[4]/span').text
        
        if(url_or_addr.find(u'网址：')!=-1):
            url = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[4]/span/a').text
            print(url)
            worksheet.write(i,4,label=url)
        if(url_or_addr.find(u'地址：')!=-1 ):
            print(url_or_addr.replace('地址：',''))
            worksheet.write(i,5,label=url_or_addr.replace('地址：',''))
    except:
        print("url null")


    try:
        addr_or_sales = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[5]/span").text
        if(addr_or_sales.find(u'地址：')!=-1 ):
            print(addr_or_sales.replace('地址：',''))
            worksheet.write(i,5,label=addr_or_sales.replace('地址：',''))
        if(addr_or_sales.find(u'经营范围：')!=-1 ):
            print(addr_or_sales.replace('经营范围：','').replace('...',''))
            worksheet.write(i,6,label=addr_or_sales.replace('经营范围：','').replace('...',''))
    except:
        print("addr null")

    try:
        sales = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[6]/span").text
        if(sales.find(u'经营范围：')!=-1 ):
            print(sales.replace('经营范围：','').replace('...',''))
            worksheet.write(i,6,label=sales.replace('经营范围：','').replace('...',''))
    except:
        print("sales null")
    
    print(" ")
    #driver.close()


if __name__=='__main__':
    reload(sys) 
    sys.setdefaultencoding('utf-8')
    f = open("inc.txt")
    line = f.readline()
    i = 0

    workbook = xlwt.Workbook(encoding="utf-8")
    worksheet = workbook.add_sheet("my worksheet")
    worksheet.write(0,0,label=u"序号")
    worksheet.write(0,1,label=u"公司名称")
    worksheet.write(0,2,label=u"电话")
    worksheet.write(0,3,label=u"邮箱")
    worksheet.write(0,4,label=u"网址")
    worksheet.write(0,5,label=u"地址")
    worksheet.write(0,6,label=u"经营范围")

    driver = webdriver.Chrome()
    while line:
        i=i+1
        print(i)
        worksheet.write(i,0,label=i)
        search(driver,i,worksheet,line.replace('\n','').replace(')','）').replace('(','（').decode('utf-8'))
        line = f.readline()
    f.close()
    driver.close()
    workbook.save('excel.xls')

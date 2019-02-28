#coding:utf-8
#指定公司全称的列表文件，该程序每次从中读取一行，去指定网页检索该公司的信息，然后存储为excel表格

from selenium import webdriver
import time
import xlwt
import sys
import re

def search(driver,i,worksheet,mytxt):
    driver.get("http://m.54114.cn/hangzhou/")#打开网页
    #mytxt = u"杭州海康威视数字技术股份有限公司"
    driver.find_element_by_xpath("/html/body/div[2]/div/form/div/div/input[2]").send_keys(mytxt)#向输入框注入待搜索字符串
    driver.find_element_by_xpath("//*[@id=\"qixc\"]").click()#单击搜索按钮

    try:
        incname=driver.find_element_by_xpath("/html/body/div[3]/div[3]/ul/li[1]/a").text #获取搜索到的第一个公司名称
    except Exception,e:#如果报错，说明网页中没有搜索结果，则在表格中该公司的一行全部填写null，然后退出
        print('1'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
        worksheet.write(i,1,label=mytxt)
        worksheet.write(i,2,label='null')
        worksheet.write(i,3,label='null')
        worksheet.write(i,4,label='null')
        worksheet.write(i,5,label='null')
        worksheet.write(i,6,label='null')
        print(mytxt)
        print('未检索到该公司的信息\n')
        return
    
    if(incname==mytxt):#如果搜索到的公司名称与输入的相等，则
        print(incname)
        worksheet.write(i,1,label=incname)# 将信息输入表格
    else:#如果搜索到的公司名称与输入的不等，则说明该网站目前未收录该公司的信息，表格里写入null后退出
        print(mytxt)
        worksheet.write(i,1,label=mytxt)
        worksheet.write(i,2,label='null')
        worksheet.write(i,3,label='null')
        worksheet.write(i,4,label='null')
        worksheet.write(i,5,label='null')
        worksheet.write(i,6,label='null')
        print('未检索到该公司的信息\n')
        return

    try:#获取详情页的地址，并单击进入该页面
        realweb = driver.find_element_by_xpath("/html/body/div[3]/div[3]/ul/li/a").get_attribute("href")
        driver.get(realweb)
    except Exception,e:#如果获取详情页的网址有误，则直接退出
        print('2'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
        return 

    
    #正则表达式提取电话号码，电话号码有多种形式，因此下面用了4种表达式，目前尚未遇到例外
    try:
        phone = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[2]/span/font/a").text
        o = re.findall(r'\d{3,4}[-]?\d{3}[-]?\d{4}',phone)#400-123-4567 或 400-1234567
        m = re.findall(r'\(?0\d{2,3}[) -]?\d{7,8}',phone)#座机
        l = re.findall(r'(\d{8,9})',phone)#座机纯8位 或纯9位的号码
        n = re.findall(r'(86)?(1\d{10})',phone)#手机
        
        #优先向表格写入座机，因为座机具备一定的信息，大公司的号码网上页能搜得到
        have_phone = 0#标记号码是否已经输出到表格中
        if (m):
            have_phone = 1
            print(m[0])
            worksheet.write(i,2,label=m[0])
        if (o and (have_phone == 0) ):
            have_phone = 1
            print(o[0])
            worksheet.write(i,2,label=o[0])
        if (l and (have_phone == 0) ):
            have_phone = 1
            print(l[0])
            worksheet.write(i,2,label=l[0])
        if (n and (have_phone == 0) ):
            have_phone = 1
            print(n[0])
            worksheet.write(i,2,label=n[0])
    except Exception,e:#“电话：暂无联系方式” 这种形势虽然有
        worksheet.write(i,2,label='null')
        print('3'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)

    have_mail = 0#标记各个属性是否已经成功提取
    have_url  = 0
    have_addr = 0
    have_sales= 0


    #邮箱、网址这2项可能会有1-2项缺失，且仅根据网页标签无法区分，只能每次获取都进行三次匹配
    #分情况讨论：
    #1 假如邮箱和网址都缺失，只有两次获取成功，分别是地址、经营范围
    #2 假如仅邮箱缺失，3次分别获取到的是网址、地址、经营范围
    #3 假如仅网址缺失，3次分别获取到的是邮箱、地址、经营范围

    try:#第1次获取
        mail_or_url = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[3]/span").text
        if(mail_or_url.find(u'邮箱')!=-1 ):
            have_mail=1
            mails = re.findall(r"[a-zA-Z0-9\.\-+_]+@[a-zA-Z0-9\.\-+_]+[\.]?[a-zA-Z]+", mail_or_url)#有些邮箱不规范，因此这里的'.'是可选项
            print(mails[0])
            worksheet.write(i,3,label=mails[0])
        else:
            pass
        if(mail_or_url.find(u'网址')!=-1):
            if(have_mail==0):
                worksheet.write(i,3,label='null')#mail is null
            have_url  = 1
            url = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[3]/span/a[1]').text
            print(url)
            worksheet.write(i,4,label=url)
        else:
            pass
        if(mail_or_url.find(u'地址：')!=-1 ):
            if(have_mail==0):
                worksheet.write(i,3,label='null')#mail is null
            if(have_url==0):
                worksheet.write(i,4,label='null')#url is null
            have_addr = 1
            print(mail_or_url.replace('地址：',''))
            worksheet.write(i,5,label=mail_or_url.replace('地址：',''))
        else:
            pass
    except IndexError:
            worksheet.write(j+1,3,label=(mail_or_url.replace('邮箱：','').replace('。','.')))#有些邮箱不规范,把点写作了句号
    except Exception,e:
        print('4'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)

    try:#第2次获取
        url_or_addr = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[4]/span').text
        if(have_url  == 0):
            if(url_or_addr.find(u'网址：')!=-1):     #/html/body/div[3]/div[4]/ul/li[4]/span/a[1]
                have_url  = 1
                url = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[4]/span/a').text
                print(url)
                worksheet.write(i,4,label=url)
        else:
            pass
        if(have_addr  == 0):
            if(url_or_addr.find(u'地址：')!=-1 ):
                if(have_url==0):
                    worksheet.write(i,4,label='null')#url is null
                have_addr = 1
                print(url_or_addr.replace('地址：',''))
                worksheet.write(i,5,label=url_or_addr.replace('地址：',''))
        else:
            pass 
        if(url_or_addr.find(u'经营范围：')!=-1 ):
            have_sales= 1
            print(url_or_addr.replace('经营范围：','').replace('...',''))
            worksheet.write(i,6,label=url_or_addr.replace('经营范围：','').replace('...',''))
        else:
            pass
        
    except Exception,e:
        print('5'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)


    try:#第3次获取
        if(have_sales== 1):#如果前面已经出现过经营范围了，后面就没必要判断了。因为经营范围是最后一个项目
            pass
        else:
            addr_or_sales = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[5]/span").text
           
            if(addr_or_sales.find(u'地址：')!=-1 ):
                if(have_url==0):
                    worksheet.write(i,4,label='null')#url is null
                print(addr_or_sales.replace('地址：',''))
                worksheet.write(i,5,label=addr_or_sales.replace('地址：',''))
            if(addr_or_sales.find(u'经营范围：')!=-1 ):
                have_sales= 1
                print(addr_or_sales.replace('经营范围：','').replace('...',''))
                worksheet.write(i,6,label=addr_or_sales.replace('经营范围：','').replace('...',''))
    except Exception,e:
        print('6'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)

    try:
        if(have_sales== 1):#如果前面已经出现过经营范围了，后面就没必要判断了。因为经营范围是最后一个项目
            pass
        else:
            sales = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[6]/span").text
            if(sales.find(u'经营范围：')!=-1 ):
                print(sales.replace('经营范围：','').replace('...',''))
                worksheet.write(i,6,label=sales.replace('经营范围：','').replace('...',''))
    except Exception,e:
        print('7'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
    print(" ")


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
        #由于该网站的检索结果中的小括号都使用全角编码，所以检索之前将可能存在的英文半角括号替换为全角括号，否则检索结果不唯一
        search(driver,i,worksheet,line.replace('\n','').replace(')','）').replace('(','（').decode('utf-8'))
        
        line = f.readline()
    f.close()
    driver.close()
    workbook.save('excel.xls')

#待添加的功能 timeout 的情况下，要workbook.save一下，把已经爬到的数据写到文件里

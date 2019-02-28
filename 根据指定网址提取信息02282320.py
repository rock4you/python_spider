#coding:utf-8
#按照指定网址获取信息

from selenium import webdriver
import time
import xlwt
import sys
import re

def search(driver,city,isdebug):

    workbook = xlwt.Workbook(encoding="utf-8")
    worksheet = workbook.add_sheet("my worksheet")
    worksheet.write(0,0,label=u"序号")
    worksheet.write(0,1,label=u"公司名称")
    worksheet.write(0,2,label=u"电话")
    worksheet.write(0,3,label=u"邮箱")
    worksheet.write(0,4,label=u"网址")
    worksheet.write(0,5,label=u"地址")
    worksheet.write(0,6,label=u"经营范围")


    url_list = []# 存放所有页面中所有公司的链接
    with open(str('incURL_'+city+'.txt'),'r') as f:
        for line in f:
            url_list.append(line)

    
    
    #对每个公司的链接，进入该网址，获取信息
    for j in range(len(url_list)):
        url = url_list[j]
        try:
            driver.get(url)
        except Exception,e:#“电话：暂无联系方式” 这种形势虽然有
            workbook.save(city+'.xls')
            print('3'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            return

        #
        company = driver.find_element_by_xpath("/html/body/div[3]/div[1]/strong").text
        if(isdebug==1):
            print(j)
            print(company)
        worksheet.write(j+1,0,label=str(j))
        worksheet.write(j+1,1,label=company)
        

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
                if(isdebug==1):
                    print(m[0])
                worksheet.write(j+1,2,label=m[0])
            if (o and (have_phone == 0) ):
                have_phone = 1
                if(isdebug==1):
                    print(o[0])
                worksheet.write(j+1,2,label=o[0])
            if (l and (have_phone == 0) ):
                have_phone = 1
                if(isdebug==1):
                    print(l[0])
                worksheet.write(j+1,2,label=l[0])
            if (n and (have_phone == 0) ):
                have_phone = 1
                if(isdebug==1):
                    print(n[0])
                worksheet.write(j+1,2,label=n[0])
        except Exception,e:#“电话：暂无联系方式” 
            worksheet.write(j+1,2,label='null')
            print('4'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'.xls')
            return

        
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
                
                mails = re.findall(r"[a-zA-Z0-9\.\-+_]+@[a-zA-Z0-9\.\-+_]+[\.]?[a-zA-Z]+", mail_or_url)
                
                if(isdebug==1):
                    print(mails[0])
                
                worksheet.write(j+1,3,label=mails[0])
                
            else:
                pass
            if(mail_or_url.find(u'网址')!=-1):
                if(have_mail==0):
                    worksheet.write(j+1,3,label='null')#mail is null
                have_url  = 1
                url = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[3]/span/a[1]').text
                if(isdebug==1):
                    print(url)
                worksheet.write(j+1,4,label=url)
            else:
                pass
            if(mail_or_url.find(u'地址：')!=-1 ):
                if(have_mail==0):
                    worksheet.write(j+1,3,label='null')#mail is null
                if(have_url==0):
                    worksheet.write(j+1,4,label='null')#url is null
                have_addr = 1
                if(isdebug==1):
                    print(mail_or_url.replace('地址：',''))
                worksheet.write(j+1,5,label=mail_or_url.replace('地址：',''))
            else:
                pass
        except IndexError:
            worksheet.write(j+1,3,label=(mail_or_url.replace('邮箱：','').replace('。','.')))
        except Exception,e:
            print('5'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'.xls')
            return

        try:#第2次获取
            url_or_addr = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[4]/span').text
            if(have_url  == 0):
                if(url_or_addr.find(u'网址：')!=-1):     #/html/body/div[3]/div[4]/ul/li[4]/span/a[1]
                    have_url  = 1
                    url = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[4]/span/a').text
                    if(isdebug==1):
                        print(url)
                    worksheet.write(j+1,4,label=url)
            else:
                pass
            if(have_addr  == 0):
                if(url_or_addr.find(u'地址：')!=-1 ):
                    if(have_url==0):
                        worksheet.write(j+1,4,label='null')#url is null
                    have_addr = 1
                    if(isdebug==1):
                        print(url_or_addr.replace('地址：',''))
                    worksheet.write(j,5,label=url_or_addr.replace('地址：',''))
            else:
                pass 
            if(url_or_addr.find(u'经营范围：')!=-1 ):
                have_sales= 1
                if(isdebug==1):
                    print(url_or_addr.replace('经营范围：','').replace('...',''))
                worksheet.write(j+1,6,label=url_or_addr.replace('经营范围：','').replace('...',''))
            else:
                pass
            
        except Exception,e:
            print('6'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'.xls')
            return


        try:#第3次获取
            if(have_sales== 1):#如果前面已经出现过经营范围了，后面就没必要判断了。因为经营范围是最后一个项目
                pass
            else:
                addr_or_sales = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[5]/span").text
            
                if(addr_or_sales.find(u'地址：')!=-1 ):
                    if(have_url==0):
                        worksheet.write(j+1,4,label='null')#url is null
                    if(isdebug==1):
                        print(addr_or_sales.replace('地址：',''))
                    worksheet.write(j+1,5,label=addr_or_sales.replace('地址：',''))
                if(addr_or_sales.find(u'经营范围：')!=-1 ):
                    have_sales= 1
                    if(isdebug==1):
                        print(addr_or_sales.replace('经营范围：','').replace('...',''))
                    worksheet.write(j+1,6,label=addr_or_sales.replace('经营范围：','').replace('...',''))
        except Exception,e:
            print('7'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'.xls')
            return

        try:
            if(have_sales== 1):#如果前面已经出现过经营范围了，后面就没必要判断了。因为经营范围是最后一个项目
                pass
            else:
                sales = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[6]/span").text
                if(sales.find(u'经营范围：')!=-1 ):
                    if(isdebug==1):
                        print(sales.replace('经营范围：','').replace('...',''))
                    worksheet.write(j+1,6,label=sales.replace('经营范围：','').replace('...',''))
        except Exception,e:
            print('8'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'.xls')
            return
        if(isdebug==1):        
            print(" ")
    workbook.save(city+'.xls')
    print(city+' is done.')
    

if __name__=='__main__':
    reload(sys) 
    sys.setdefaultencoding('utf-8')
    driver = webdriver.Chrome()
    isdebug = 1

    city =['hangzhou']#,'shenzhen','hangzhou']
    for i in range(len(city)):
        search(driver,city[i],isdebug)
    
    time.sleep(5)
    driver.close()
    
#上海 bug网址 http://m.54114.cn/hangye90/8f3678a2d1.html

#写一个函数，读文件，文件里每一行都是一个公司信息页面。如果中断，还可以通过这个方式继续进行下去。

#TimeoutException() 超时类型的错误可以考虑加一个断点重启的功能


#excel添加 城市 字段，如 xx市


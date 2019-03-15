#coding:utf-8
#获取指定城市所有类型的公司信息，每种类型的公司保存为一个excel表格 
#默认不显示浏览器图形界面

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
import time
import xlwt
import sys
import re

def search(driver,city,isdebug,web_order):
    try:#浏览器打开网页
        driver.get("http://m.54114.cn/"+city+'/'+web_order +'/')#-------------------网址修改之一
    except Exception,e:#如果网页打开有误，则直接退出
        print('0'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
        return 
    s=(driver.find_element_by_xpath("/html/body/div[4]").text)#.encode('utf-8')  #获取公司总数目
    try:
        s=s.split(' ')[11]
        s=s.replace(u'下','').replace(u'一','').replace(u'页','').replace('>','')
    except IndexError:#有些行业可能只有一个页面、甚至只有一个公司
        pass
    try:
        page_num=int(s)
    except:#有些行业甚至一个公司都没有
        return

    workbook = xlwt.Workbook(encoding="utf-8")
    worksheet = workbook.add_sheet("my worksheet")#,cell_overwrite_ok=True)#解决重写报错
    worksheet.write(0,0,label=u"序号")
    worksheet.write(0,1,label=u"公司名称")
    worksheet.write(0,2,label=u"电话")
    worksheet.write(0,3,label=u"邮箱")
    worksheet.write(0,4,label=u"网址")
    worksheet.write(0,5,label=u"地址")
    worksheet.write(0,6,label=u"经营范围")

    url_num=page_num*20
    #url_num = (int)(driver.find_element_by_xpath("/html/body/div[4]/span[1]").text.replace(u'共','').replace(u'纪录',''))

    page_url = []  #获取每个页面的链接，存放到page_url
    for i in range(1,page_num+1):
        page_url.append("http://m.54114.cn/"+city+"/"+web_order+"_p"+str(i)+"/")#-------------------网址修改之二
    
    url_list = []# 存放所有页面中所有公司的链接

    for i in range(page_num):#对每个页面，尝试获取每个公司的链接
        current_page = page_url[i]
        driver.get(current_page)
        try:
            for j in range(1,21):
                xpath = "/html/body/div[3]/div[3]/ul/li["+str(j)+"]/a"
                url_list.append(driver.find_element_by_xpath(xpath).get_attribute("href"))
        except NoSuchElementException:
            pass
        except Exception,e:
            print('1'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print(current_page)
            print('order:'+str(i+1))
            return
    
    if isdebug==1:#如果是调试模式，则把所有公司信息页面的网址都写入到文件中
        fp = open(str("incURL_"+city+".txt"),"w+")
        try:
            for i in range(1000):
                fp.write(str(url_list[i])+'\n')
        except IndexError:#如果某页面列出的公司数量没有20个,忽略即可
            pass
        except Exception,e:
            print('2'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(i+1))
            return
            #print(len(url_list ))
        fp.close()
    
    for j in range(url_num):#对每个公司的链接，进入该网址，获取信息
        try:
            url = url_list[j]
        except IndexError:#默认的是每页有20个公司。如果没有这么多，则直接退出for循环
            break
        try:
            driver.get(url)
        #//如果打开有误，说明实际的网页数量并没有url_num这么多，退出即可
        except Exception,e:#“电话：暂无联系方式” 这种形势虽然有
            workbook.save(city+'_'+web_order+'.xls')
            print('3'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            return
        company = driver.find_element_by_xpath("/html/body/div[3]/div[1]/strong").text
        if(isdebug==1):
            print(j+1)
            print(company)
        worksheet.write(j+1,0,label=str(j+1))
        worksheet.write(j+1,1,label=company)
        
        #正则表达式提取电话号码，电话号码有多种形式，因此下面用了4种表达式，例外 电话：(0571);
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
            elif (o and (have_phone == 0) ):
                have_phone = 1
                if(isdebug==1):
                    print(o[0])
                worksheet.write(j+1,2,label=o[0])
            elif (l and (have_phone == 0) ):
                have_phone = 1
                if(isdebug==1):
                    print(l[0])
                worksheet.write(j+1,2,label=l[0])
            elif (n and (have_phone == 0) ):
                have_phone = 1
                if(isdebug==1):
                    print(n[0])
                worksheet.write(j+1,2,label=n[0])
            else:#处理例外情况 电话：(0571);
                worksheet.write(j+1,2,label='null')
        except NoSuchElementException:#“电话：暂无联系方式”
            worksheet.write(j+1,2,label='null')
        except Exception,e: 
            print('4'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'_'+web_order+'.xls')
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
            if(mail_or_url.find(u'邮箱')==0 ):
                have_mail=1
                mails = re.findall(r"[a-zA-Z0-9\.\-+_]+@?[a-zA-Z0-9\.\-+_]+[\.]?[a-zA-Z]+", mail_or_url)#邮箱不规范，@后没写点，因此这里的.设置为可选项
                if(isdebug==1):
                    print(mails[0])
                worksheet.write(j+1,3,label=mails[0])
            else:
                pass
            if(mail_or_url.find(u'网址')==0):
                if(have_mail==0):
                    worksheet.write(j+1,3,label='null')#mail is null
                have_url  = 1
                url = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[3]/span/a[1]').text
                if(isdebug==1):
                    print(url)
                worksheet.write(j+1,4,label=url)
            else:
                pass
            if(mail_or_url.find(u'地址：')==0 ):
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
            worksheet.write(j+1,3,label=(mail_or_url.replace('邮箱：','').replace('。','.')))#有些邮箱不规范,把点写作了句号
        except Exception,e:
            print('5'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'_'+web_order+'.xls')
            return

        try:#第2次获取
            url_or_addr = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[4]/span').text
            if(have_url  == 0):
                if(url_or_addr.find(u'网址：')==0):     #/html/body/div[3]/div[4]/ul/li[4]/span/a[1]
                    have_url  = 1
                    inc_url = driver.find_element_by_xpath('/html/body/div[3]/div[4]/ul/li[4]/span/a').text
                    if(isdebug==1):
                        print(inc_url)
                    worksheet.write(j+1,4,label=inc_url)
            else:
                pass
            if(have_addr == 0):
                if(url_or_addr.find(u'网址：')==-1 and url_or_addr.find(u'地址：')==0 ):#有些情况下，网址的那一行里有“下载地址”。 首字符匹配到“地址”并且没有出现“网址”才能算作地址
                    if(have_url==0):
                        worksheet.write(j+1,4,label='null')#url is null
                    have_addr = 1
                    if(isdebug==1):
                        print(url_or_addr.replace('地址：',''))
                    worksheet.write(j+1,5,label=url_or_addr.replace('地址：',''))
            else:
                pass 
            if(url_or_addr.find(u'经营范围：')==0 ):
                have_sales= 1
                if(have_addr==0):
                    worksheet.write(j+1,5,label='null')
                if(isdebug==1):
                    print(url_or_addr.replace('经营范围：','').replace('...',''))
                worksheet.write(j+1,6,label=url_or_addr.replace('经营范围：','').replace('...',''))
            else:
                pass 
        except Exception,e:
            print('6'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'_'+web_order+'.xls')
            return

        try:#第3次获取
            if(have_sales== 1):#如果前面已经出现过经营范围了，后面就没必要判断了。因为经营范围是最后一个项目
                pass
            else:
                addr_or_sales = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[5]/span").text
            
                if(addr_or_sales.find(u'地址：')==0 ):#有的在经营范围里出现了 停车地址...  and have_addr==0
                    have_addr=1
                    if(have_url==0):
                        worksheet.write(j+1,4,label='null')#url is null
                    if(isdebug==1):
                        print(addr_or_sales.replace('地址：',''))
                    worksheet.write(j+1,5,label=addr_or_sales.replace('地址：',''))
                if(addr_or_sales.find(u'经营范围：')==0 ):
                    have_sales= 1
                    if(have_addr==0):
                        worksheet.write(j+1,5,label='null')
                    if(isdebug==1):
                        print(addr_or_sales.replace('经营范围：','').replace('...',''))
                    worksheet.write(j+1,6,label=addr_or_sales.replace('经营范围：','').replace('...',''))
        except Exception,e:
            print('7'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'_'+web_order+'.xls')
            return

        try:
            if(have_sales== 1):#如果前面已经出现过经营范围了，后面就没必要判断了。因为经营范围是最后一个项目
                pass
            else:
                sales = driver.find_element_by_xpath("/html/body/div[3]/div[4]/ul/li[6]/span").text
                if(sales.find(u'经营范围：')==0 ):
                    #这里曾经重写报错Attempt to overwrite cell: sheetname=u'my worksheet' rowx=1 colx=5 Exception
                    #报错的原因:前面写入地址后，没有更新have_addr的值，导致这里重复写入地址null
                    if(have_addr==0):
                        worksheet.write(j+1,5,label='null')
                    if(isdebug==1):
                        print(sales.replace('经营范围：','').replace('...',''))
                    worksheet.write(j+1,6,label=sales.replace('经营范围：','').replace('...',''))
        except Exception,e:
            print('8'+str(Exception)+' '+str(e)+' '+repr(e)+' '+e.message)
            print('order:'+str(j))
            print(url)
            workbook.save(city+'_'+web_order+'.xls')
            return
        if(isdebug==1):        
            print(" ")
    workbook.save(city+'_'+web_order+'.xls')
    print(city+' '+web_order+' is done.')

if __name__=='__main__':
    reload(sys) 
    sys.setdefaultencoding('utf-8')

    isdebug = 0 #如果是1就在终端打印信息，如果是0就不打印
    
    if(isdebug==0):#默认不显示浏览器图形界面
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        driver = webdriver.Chrome(chrome_options=chrome_options)
    else:#如果是调试模式，则显示浏览器图形界面
        driver = webdriver.Chrome()

    city =[ 'beijing','shanghai','guangzhou','shenzhen','hangzhou']
    
    time_start = time.time()
    for i in range(len(city)):
        for j in range(1,21):
            web_order ='hangye'
            web_order =web_order+str(j)
            search(driver,city[i],isdebug,web_order)
    #time.sleep(5) #信息爬取完毕之后，网页显示5秒再关闭
    driver.close()
    driver.quit()
    time_end = time.time()
    print('time cost:',time_end-time_start,'s')
    
#后期考虑，TimeoutException() 超时类型的错误可以通过读取网址列表的形式，增加一个断点重启的功能
#后期考虑，在excel添加一栏 城市 字段，如 xx市。需要正则提取xx省xx市、xx市两种情况

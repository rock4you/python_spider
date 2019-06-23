#python 2.*
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time
import xlwt
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

#伪装成浏览器，防止被识破
option = webdriver.ChromeOptions()
option.add_argument('--user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.146 Safari/537.36"')
driver = webdriver.Chrome(chrome_options=option)

#打开登录页面
driver.get('https://www.qichacha.com/user_login')
#单击用户名密码登录的标签
tag = driver.find_element_by_xpath('//*[@id="normalLogin"]')
tag.click()
#将用户名、密码注入
driver.find_element_by_id('nameNormal').send_keys('username')
driver.find_element_by_id('pwdNormal').send_keys('password')
time.sleep(10)#休眠，人工完成验证步骤，等待程序单击“登录”
#单击登录按钮
btn = driver.find_element_by_xpath('//*[@id="user_login_normal"]/button')
btn.click()

inc_list = ['阿里巴巴','腾讯','今日头条','滴滴','美团']
inc_len = len(inc_list)

for i in range(inc_len):
    txt = inc_list[i]
    time.sleep(1)
    
    if (i==0):
        #向搜索框注入文字
        txt=txt.decode('utf-8')
        driver.find_element_by_id('searchkey').send_keys(txt)
        #单击搜索按钮
        srh_btn = driver.find_element_by_xpath('//*[@id="V3_Search_bt"]')
        srh_btn.click()
    else:
        #向搜索框注入下一个公司地址
        txt=txt.decode('utf-8')
        driver.find_element_by_id('headerKey').send_keys(txt)
        #搜索按钮 
        srh_btn = driver.find_element_by_xpath('/html/body/header/div/form/div/div/span/button')
        srh_btn.click()

    #获取首个企业文本
    print(i+1)
    inc_full = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/a').text                                
    print(inc_full)
    money = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/p[1]/span[1]').text
    print(money)
    date = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/p[1]/span[2]').text
    print(date)
    mail_phone = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/p[2]').text
    print(mail_phone)
    addr = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/p[3]').text
    print(addr)
    try:
        stock_or_others = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/p[4]').text
        print(stock_or_others)
    except:
        pass

    #获取网页地址，进入
    inner = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/a').get_attribute("href")
    driver.get(inner)

    #单击进入后 官网 通过href属性获得：
    inc_web = driver.find_element_by_xpath('//*[@id="company-top"]/div[2]/div[2]/div[3]/div[1]/span[3]/a').get_attribute("href")
    print("官网："+inc_web)
    print(' ')

driver.close()

#bug list:
#UnicodeDecodeError: 'utf8' codec can't decode byte 0xe9 in position 0: unexpected end of data
#原因：向搜索栏注入中文字符串时，必须先采用如下方式转换成utf-8编码
#解决：send_keys("阿里巴巴".decode('utf-8'))

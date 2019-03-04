#在企查查输入企业简称，获取企业全称
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import time
import xlwt
import sys

reload(sys)
sys.setdefaultencoding('utf-8')
driver = webdriver.Chrome()

#打开登录页面
driver.get('https://www.qichacha.com/user_login')

#单击用户名密码登录的标签
tag = driver.find_element_by_xpath('//*[@id="normalLogin"]')
tag.click()

#将用户名、密码注入
driver.find_element_by_id('nameNormal').send_keys('18569502048')
driver.find_element_by_id('pwdNormal').send_keys('041812')

time.sleep(10)#休眠，人工完成验证步骤，等待程序单击“登录”

#单击登录按钮
btn = driver.find_element_by_xpath('//*[@id="user_login_normal"]/button')
btn.click()
time.sleep(1)

#向搜索框注入文字
txt='阿里巴巴'.decode('utf-8')
driver.find_element_by_id('searchkey').send_keys(txt)

#单击搜索按钮
srh_btn = driver.find_element_by_xpath('//*[@id="V3_Search_bt"]')
srh_btn.click()

#获取首个企业文本
inc = driver.find_element_by_xpath('//*[@id="search-result"]/tr[1]/td[3]/a').text                                
print(inc)
time.sleep(5)
driver.close()

#bug list:
#UnicodeDecodeError: 'utf8' codec can't decode byte 0xe9 in position 0: unexpected end of data
#原因：向搜索栏注入中文字符串时，必须先采用如下方式转换成utf-8编码
#解决：send_keys("阿里巴巴".decode('utf-8'))

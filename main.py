#!_*_coding:utf8_*_
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
from selenium.webdriver.common.action_chains  import ActionChains
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import time
import json
import xlwt
import datetime
from xls_input import xls_input
now = datetime.datetime.now()
year = now.strftime('%Y')
month1 = now.strftime('%m')
month = int(month1)-1
day = now.strftime('%d')

workbook =  xlwt.Workbook(encoding='utf-8')
booksheet  =  workbook.add_sheet('Sheet1')
booksheet.write(0,0,'编号')
booksheet.write(0,1,'流程编号')
booksheet.write(0,2,'IDC编号')
booksheet.write(0,3,'机房')
booksheet.write(0,4,'用户')
booksheet.write(0,5,'人数')
booksheet.write(0,6,'入时间')
booksheet.write(0,7,'出时间    ')
booksheet.write(0,8,'时间')
booksheet.write(0,9,'事由类型及说明')
booksheet.write(0,10,'录入时间')
booksheet.write(0,11,'其他')
info = {
}

browser = webdriver.Firefox() # Get local session of firefox
browser.get("http://bms.stocom.net") # Load page
firsthandle=browser.current_window_handle
elem = browser.find_element_by_id("p_username")
elem.send_keys("runser2")
elem = browser.find_element_by_id("p_password")
elem.send_keys("Stocom20151001")
time.sleep(5)
elem = browser.find_element_by_class_name("loginbutton")
elem.click()
time.sleep(6)

elem = browser.find_element_by_link_text("我的工作")
ActionChains(browser).move_to_element(elem).perform()
elem.click()
elem = browser.find_element_by_link_text("已办事宜-数据中心")
elem.click()
allhandles=browser.window_handles
for handle in allhandles:
  if handle != firsthandle:
    secondhandle = handle
    browser.switch_to_window(secondhandle)
   # print 'now new windows'
browser.maximize_window()
time.sleep(3)
elem = browser.find_element_by_id("inqu_status-0-begTime")
elem.click()
#lookfor the date


elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/div[1]/div/select[1]")
elem.click()
elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/div[1]/div/select[1]/option[@value=%s]"% year)
elem.click()
elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/div[1]/div/select[2]")
elem.click()
elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/div[1]/div/select[2]/option[@value=%s]"%month)
elem.click()
date_count=1
while True:
    try:
        elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/table/tbody/tr[1]/td[%s]/a"%date_count)
        break
    except:
        date_count += 1
#print '---->',date_count
date_location = [1,date_count]
#print date_location
short_of = int(day)-1
if date_location[1]+short_of>7:
    date_location[0]=(date_location[1]+short_of)/7+1
    date_location[1]=(date_location[1]+short_of)%7
else:
    date_location[1]=date_location[1]+short_of
#print date_location

elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/table/tbody/tr[%s]/td[%s]/a"%(date_location[0],date_location[1]))
elem.click()


elem = browser.find_element_by_id("inqu_status-0-endTime")
elem.click()
elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/div[1]/div/select[1]")
elem.click()
elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/div[1]/div/select[1]/option[@value=%s]"% year)
elem.click()
elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/div[1]/div/select[2]")
elem.click()
elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/div[1]/div/select[2]/option[@value=%s]"% month)
elem.click()
elem = browser.find_element_by_xpath("//*[@id='ui-datepicker-div']/table/tbody/tr[%s]/td[%s]/a"%(date_location[0],date_location[1]+1))
elem.click()
elem = browser.find_element_by_id("inqu_status-0-processTitle")
elem.send_keys(u"数据中心临时人员申请流程")
elem = browser.find_element_by_xpath("//*[@id='button_query']")
elem.click()

m = 0
while True:
    try:
        info[m] = {"编号": None,
                   "流程号": None,
                   "IDC编号": None,
                   "机房": None,
                   "用户": None,
                   "人数": None,
                   "入时间": None,
                   "出时间": None,
                   "时间": None,
                   "事由类型及说明": None,
                   "录入时间": None,
                   "其他": None}
        elem = browser.find_element_by_xpath("//tr[@rawrowindex=%s]/td[@id='ef_grid_resultDataTdprocessSequenceId']/div/a"%m)
        info[m]["流程号"]=elem.text
        elem = browser.find_element_by_xpath("//tr[@rawrowindex=%s]/td[@id='ef_grid_resultDataTdmanage']/div/a[1]/span/span"%m)
        ActionChains(browser).move_to_element(elem).perform()
        elem.click()
        print m
        allhandles=browser.window_handles
        for handle in allhandles:
          if handle != firsthandle and handle != secondhandle:
            thirdhandle = handle
            browser.switch_to_window(thirdhandle)
        browser.maximize_window()
        browser.implicitly_wait(5)
        elem = browser.find_element_by_xpath("//div[@id='formDiv']/table/tbody/tr[2]/td[2]/div")
        info[m]["编号"]=elem.text
        elem = browser.find_element_by_xpath("//*[@id='formDiv']/table/tbody/tr[3]/td[2]/div")
        info[m]["用户"]=elem.text
        elem = browser.find_element_by_xpath("//*[@id='formDiv']/table/tbody/tr[4]/td[2]/div")
        info[m]["申请人"]=elem.text
        elem=browser.find_element_by_xpath("//*[@id='formDiv']/table/tbody/tr[4]/td[4]/div")
        info[m]["联系手机"]=elem.text
        elem = browser.find_element_by_xpath("//*[@id='formDiv']/table/tbody/tr[5]/td[2]/div")
        info[m]["事由类型及说明"]=elem.text
        elem = browser.find_element_by_xpath("//*[@id='formDiv']/table/tbody/tr[7]/td[2]/div")
        info[m]["机房"]=elem.text
        elem = browser.find_element_by_xpath("//*[@id='formDiv']/table/tbody/tr[7]/td[4]/div")
        info[m]["机房号"]=elem.text
        elem = browser.find_element_by_xpath("//*[@id='formDiv']/table/tbody/tr[8]/td[2]/div")
        info[m]["入时间"]=elem.text
        start_time = elem.text
        elem = browser.find_element_by_xpath("//*[@id='formDiv']/table/tbody/tr[8]/td[4]/div")
        info[m]["出时间"]=elem.text
        end_time = elem.text
        elem = browser.find_element_by_xpath("//*[@id='formDiv']/table/tbody/tr[10]/td[2]/div")
        info[m]["IDC编号"]=elem.text
        info[m]["录入时间"]=datetime.datetime.now().strftime('%Y-%m-%d')
        info[m]["时间"]=str(start_time)+'--'+str(end_time)
        n = 0
        while True:
          try:
             n+= 1
             elem = browser.find_element_by_xpath("//*[@id='ef_grid_visitor__data_table']/tbody/tr[%s]"% n)
          except:
            n-=1
            break
        info[m]["人数"]=n
        browser.close()
        browser.switch_to_window(secondhandle)
        m+=1
    except:
      break
print json.dumps(info,encoding="utf-8",ensure_ascii=False)


for i in range(m):
    for o in range(i+1,m):
      if info[i]["流程号"] == info[o]["流程号"]:
        info.pop(i)
        break
      else:
        pass
#print json.dumps(info,encoding="utf-8",ensure_ascii=False)

print json.dumps(info,encoding="utf-8",ensure_ascii=False)
count = info.keys()
print count[-1]
allcount = len(info.keys())
xls_input(count,info)
#browser.close()
#browser.switch_to_window(firsthandle)
#browser.close()

# -*- coding:utf-8 -*-
__author__ = 'liuyun'
import urllib
import urllib2
import re
import string
import xlrd
import random,time
import xlwt
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

# page = 1
# url = 'www.baidu.com'
user_agent = 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'
headers = { 'User-Agent' : user_agent }

def baidu_search(keyword):
    p= {'q': keyword}
    # request=urllib2.Request("http://www.baidu.com/s?"+urllib.urlencode(p),headers=headers)
    request=urllib2.Request("http://cn.bing.com/search?"+urllib.urlencode(p),headers=headers)
    res= urllib2.urlopen(request)
    return res

def getnum(regex,text):
    res = re.findall(regex, text)
    return res
new_hw = ['虚拟现实','智能驾驶','无人机','可回收卫星','机器人厨师']
data = xlrd.open_workbook('intelli_hardware.xlsx')
table = data.sheets()[1]
nrows = table.nrows
col1 = table.col_values(1)
search_res = []
search_res2 = []
workbook = xlsxwriter.Workbook('zm10.xlsx')  #创建一个excel文件
worksheet = workbook.add_worksheet()        #创建一个工作表对象

j = 0
for v in col1:
    worksheet.write(j,0,v)
    j = j +1

k = 1
for nh in new_hw:
    worksheet.write(0,k,nh)
    k = k +1


icol = 1
for nh in new_hw:
    search_res = []
    iraw = 1
    for v in col1[1:]:
        if v:
            key_search =  "\"" + v.encode('utf-8') + " "+"\"" + "\"" + nh + "\""
            try:
                keyword = key_search
                res = baidu_search(keyword)
                response = res.read()
                # print response
                resNum = getnum('(<span class="sb_count">)(\w.*?)(?=条结果)', response)
                # resNum = getnum("(结果约)(\w.*?)(?=个)", response)
                if resNum:
                    num = string.atoi(resNum[0][1].replace(',',''))
                else:
                    num = 0
                search_res.append(num)
                print key_search + ':' + str(num)
                time.sleep(random.uniform(0,1))
            except urllib2.URLError, e:
                if hasattr(e,"code"):
                    print e.code
                if hasattr(e,"reason"):
                    print e.reason
            worksheet.write(iraw,icol,num)
            iraw = iraw + 1
    search_res2.append(search_res)
    icol = icol + 1

flag = 1
workbook.close()
print 'over'
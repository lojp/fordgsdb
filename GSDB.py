import requests
import json
import re
from bs4 import BeautifulSoup
import pprint
import urllib.parse as urlparse
from urllib.parse import urlencode
import urllib.request
import pandas as pd
import xlrd
from xlutils.copy import copy

payload = {'cdsid': 'jluo27', 'b64Pwd': 'eW91bmcxMXo=','WslIP':'19.244.68.41','fastRegister':'No'}
login_url = 'https://www.wsl.ford.com/login.cgi'
home_url = 'https://web.gsdb2.ford.com/GSDBHomepageWeb/home.do'
urlpre = 'https://web.gsdb2.ford.com/GSDBHomepageWeb/siteInformationSearchPost.do?method=httpGet&site='
data = xlrd.open_workbook('C:/Users/jluo27/Desktop/GSDB.xlsx')
wb = copy(data)
sht = data.sheets()[0]  
nrows = sht.nrows
s = requests.Session()  # 可以在多次访问中保留cookie
s.post(login_url, data=payload)  # POST帐号和密码，设置headers

for i in range(nrows)[1:]:
	values = sht.cell(i,0).value
	url = urlpre + values.lower()
	s.post(home_url)
	r = s.get(url)  # 已经是登录状态了
	dat = r.text
	soup = BeautifulSoup(dat,'html.parser')
	table = soup.find_all(class_='search_results_table')
	for row in table:
		tablecells = row.findAll('td')
		suppcode = str(tablecells[1])
		suppname = str(tablecells[7])
		sitename = str(tablecells[23])
		sitecty = str(tablecells[55])
		sh = wb.get_sheet(0)
		sh.write(i,1,re.sub(r'\s+', ' ',re.sub(r'(\<.*?\>)', '',suppcode)).strip())
		sh.write(i,2,re.sub(r'\s+', ' ',re.sub(r'(\<.*?\>)', '',suppname)).strip())
		sh.write(i,3,re.sub(r'(\<.*?\>)', '',sitename).strip())
		sh.write(i,4,re.sub(r'\s+', ' ',re.sub(r'(\<.*?\>)', '',sitecty)).strip())
		wb.save('C:/Users/jluo27/Desktop/xhxh.xls')
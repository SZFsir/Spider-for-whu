# -*- coding: utf-8 -*-
# 爬取武汉大学教务系统
#python2.7
#2018.3.11
#通过本地保存cookies,实现绕过验证码

import requests
import md5
import re
import cPickle
import xlrd
from xlutils.copy import copy
import xlwt
from bs4 import BeautifulSoup
import time
from urllib import urlencode
import traceback


url_login='http://210.42.121.241/servlet/Login'
url_img = 'http://210.42.121.241/servlet/GenImg'
url_stu = 'http://210.42.121.241/stu/stu_course_parent.jsp'
url_cou = 'http://210.42.121.241/servlet/Svlt_QueryStuLsn'
url_stu_info = 'http://210.42.121.241/stu/student_information.jsp'
user_agent = 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT'
headers = {'User-Agent':user_agent,'Referer':url_login}



#产生验证码到指定位置
def get_cap(stream):
	with open('vcode.jpg','wb') as f:
 		f.write(stream.content)
 	return str(raw_input("输入验证码:"))

#密码md5加密
def pwd_md5(pwd):
 	str1 = str(pwd)
 	md = md5.new()
 	md.update(str1.encode(encoding = 'utf-8'))
 	return md.hexdigest()

def login(username,password,cookies,headers,vcode):


	#观察到武汉大学教务系统登录post 提交的数据只有username、password、captcha.所以直接提交数据即可
	pwd_md = pwd_md5(password)
	post_data = {'id':username,'pwd':pwd_md,'xdvfb':vcode}
	r = requests.post(url_login,data=post_data,cookies = cookies,headers=headers)
	pattern = re.compile(r'inputWraper')
	#if pattern.findall(r.text):
		#print 'wrong!!!!!!!!!!!!'
		#return None

	#print r.text
	return r.text

#爬取课表
def crawl_schedule(index_html,cookies,headers):
	#发现课表存在与一个构造url中，其中参数csrftoken可以从主页提取，构造正则提取
	try:
		pattern = re.compile(r'csrftoken=(.*?)\',')
		csrftoken = pattern.findall(index_html)[0]
		print csrftoken
		#构造url
		get_url='?csrftoken=%s'%csrftoken+r'&action=normalLsn&year=2017&term=%CF%C2&state='
		url_course = url_cou + get_url
		#得到课表html
		stu = requests.get(url_course,cookies = cookies,headers=headers)
		print stu.text
	except:
		print 'nothing'

#加载cookies数据
def load_session():
	try:
		with open('session.txt','rb') as f:
			headers = cPickle.load(f)
			cookies = cPickle.load(f)
			vcode = cPickle.load(f)
	except Exception,e:

		session,vcode = get_session()
		return session.headers,session.cookies,vcode

	r = requests.get(url_stu_info,cookies = cookies,headers = headers)

	if r.status_code != 200:
		session = get_session()
		return session.headers,session.cookies
	print "load_session successful!!"
	return headers,cookies,vcode


#第一次运行人工识别验证码并输入有效数据，获得cookies并保存
def get_session():
	print "验证码过期，请重新登录：\n"
	session = requests.session()
	#先产生验证码
	img_stream = session.get(url_img,stream = True,headers = headers)
	vcode = get_cap(img_stream)
	#观察到武汉大学教务系统登录post 提交的数据只有username、password、captcha.所以直接提交数据即可
	#第一次为获取cookies
	stu_id = raw_input("输入学生号: ")
	pwd_md = pwd_md5(int(raw_input("输入密码: ")))
	post_data = {'id':stu_id,'pwd':pwd_md,'xdvfb':vcode}
	r = session.post(url_login,data=post_data,headers=headers)
	if r.status_code == 200:
		with open('session.txt','wb') as f:
			cPickle.dump(session.headers,f)
			cPickle.dump(session.cookies.get_dict(),f)
			cPickle.dump(vcode,f)
		return session,vcode
	return "get_session occurs problem"

#从excel中读取学生账号密码信息
def get_stu_login_data():
	workbook = xlrd.open_workbook(r'info.xlsx')
	all_sheet_list = workbook.sheet_names()
	return workbook

#修改excel方法
def init_store_stu_info():
	workbook = xlrd.open_workbook(r'info1.xlsx')
	all_sheet_list = workbook.sheet_names()
	workbooknew = copy(workbook)
	for i in range(0,5):
		sheet =  workbooknew.get_sheet(i)
		sheet.write(0,10,u'密码是否修改')
		sheet.write(0,11,u'籍贯')
	return workbooknew

def creat_excel_to_store():
	f = xlwt.Workbook()
	sheet_score = f.add_sheet(u'all',cell_overwrite_ok=True)
	return f,sheet_score




#爬取
def crawl():
	#获取session
	headers,cookies,vcode = load_session()
	#获取学生登录信息数据表
	workbook = get_stu_login_data()
	#
	workbook_store = init_store_stu_info()

	file,sheet_score = creat_excel_to_store()

	for sheet_index in range(0,5):
		sheet = workbook.sheet_by_index(sheet_index)
		sheet_store = workbook_store.get_sheet(sheet_index)

		for i in range(1,sheet.nrows):
			name = sheet.cell_value(i,4)
			stu_id = int(sheet.cell_value(i,3))
			identitycode = sheet.cell_value(i,8)
			password = int(identitycode[6:14])
			print stu_id
			#print password
			#登录教务系统,获取主页html
			index_html = login(stu_id,password,cookies,headers,vcode)


			#here is the part of crawl the info from the system
			# try:
				
			# 	if index_html == None:
			# 		raise "login have problem"

			# 	#个人信息get就能得到
			# 	stu_info_html = requests.get(url_stu_info,cookies = cookies,headers=headers)
			# 	#print stu_info_html.text

			# 	soup = BeautifulSoup(stu_info_html.text,'html.parser',from_encoding = 'utf-8')


			# 	tds = soup.find_all('td')				
			# 	hometown = tds[5].string
			# 	school = tds[6].string
			# 	major = tds[7].string	

			# 	if index_html == None:
			# 		raise "login have problem"


			# 	sheet_store.write(i,11,hometown)
			# 	sheet_store.write(i,6,school)
			# 	sheet_store.write(i,7,major)
			# 	sheet_store.write(i,10,u'暂时没有修改密码')
			# 	print hometown

			# except Exception,e:
			# 	print e
			# 	sheet_store.write(i,10,u'已经修改了密码')
			# #爬取课表
			#crawl_schedule(index_html,cookies,headers)

			
			score_html = crawl_score(index_html,cookies,headers)
			try:
				soup = BeautifulSoup(score_html,'html.parser',from_encoding = 'utf-8')
				trs = soup.find_all('tr')
				sheet_score.write(3*i-2,0,name)
				sheet_score.write(3*i-2,1,str(stu_id))
				m = 0
				for tr in trs:
					try:
						tds = tr.find_all('td')
						sheet_score.write(3*i-1,m,tds[1].string)
						sheet_score.write(3*i,m,tds[9].string)
						#print tds[1],tds[9]
						m = m+1
					except:
						print 'nothing'
						#traceback.print_exc()
			except Exception,e:
				print e
				print 'get score wrong'
				#traceback.print_exc()

	file.save('score.xlsx')



			
			


	workbook_store.save('info1.xlsx')


def crawl_score(index_html,cookies,headers):
	#发现课表存在与一个构造url中，其中参数csrftoken可以从主页提取，构造正则提取
	try:
		pattern = re.compile(r'csrftoken=(.*?)\',')
		csrftoken = pattern.findall(index_html)[0]
		#print csrftoken
		#构造url
		url_z = 'http://210.42.121.241/servlet/Svlt_QueryStuScore'
		get_url='?csrftoken=%s'%csrftoken+r'&year=0&term=&learnType=&scoreFlag=0&t='
		
		url_t = str(time.strftime('%a%%20%b%%20%d%%20%Y%%20%H:%M:%S')) + '%20GMT+0800%20(CST)'
		url_score = url_z + get_url +url_t 
		#print url_score

		#得到课表html
		stu = requests.get(url_score,cookies = cookies,headers=headers)
		#print stu.text
		return stu.text
	except Exception,e:
		print e
		print 'nothing'
		#traceback.print_exc()






if __name__ == '__main__':
	crawl()




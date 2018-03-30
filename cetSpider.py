# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
import re
import xlrd
import xlwt
import sys
import traceback
import codecs
from proxy_test import getip
import random


def crawl():
	reload(sys)
	sys.setdefaultencoding('utf8')

	bad_ip = []
	ip_list,ip_len = getip()





	url = 'http://202.114.74.136/cet/Default.aspx'
	user_agent = 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT'
	headers = {'User-Agent':user_agent,'Referer':url}
	proxies = {
	"http":"http://121.58.17.52:80",
	"https":"http://121.58.17.52:80"
	}

	# to get the vcode
	session = requests.session()
	page1 = session.get('http://202.114.74.136/cet/createImg.aspx',stream = True,headers = headers)

	#print(page1.content)
	# pattern = re.compile(r'GIF(.*?)<DOCTYPE')
	# stream = pattern.findall(page1.content)


	with open('vcode4.gif','wb') as f:
	 	f.write(page1.content)
	 
	print  page1.content
	vcode = str(raw_input())
	print vcode

	html_content = session.get(url,headers = headers)
	soup = BeautifulSoup(html_content.content,'html.parser')

	print html_content.content
	# to provide all of needed parameter to crawl in post
	info1 = soup.find(id = '__EVENTVALIDATION').get('value')
	info2 = soup.find(id = '__VIEWSTATE').get('value')
	info3 = soup.find(id = '__VIEWSTATEGENERATOR').get('value')


	#get student info through excel file

	f = xlwt.Workbook()


	#to store data to excel
	#inital excel file
	sheet_store1 = f.add_sheet(u'cs',cell_overwrite_ok=True)
	sheet_store2 = f.add_sheet(u'xa',cell_overwrite_ok=True)
	sheet_store3 = f.add_sheet(u'wl',cell_overwrite_ok=True)
	sheet_store4 = f.add_sheet(u'zg',cell_overwrite_ok=True)
	sheet_store5 = f.add_sheet(u'wa',cell_overwrite_ok=True)
	sheet_store = [sheet_store1,sheet_store2,sheet_store3,sheet_store4,sheet_store5]
	for sh in sheet_store:	
		sh.write(0,0,u"examTime")
		sh.write(0,1,u"degree")
		sh.write(0,2,u'score')
		sh.write(0,3,u"stuId")
		sh.write(0,4,u"name")
		sh.write(0,5,u"gender")
		sh.write(0,6,u"school")
		sh.write(0,7,u"major")
		sh.write(0,8,u"idNum")
		sh.write(0,9,u"examNum")

	#open the original info source
	workbook = xlrd.open_workbook(r'naili.xlsx',)
	all_sheet_list = workbook.sheet_names()


	try:
		for sheet_index in range(0,2):
			sheet = workbook.sheet_by_index(sheet_index)
			k = 1
			# f = codecs.open('stu.txt','a',encoding = 'utf-8')
			for i in range(1,sheet.nrows):

				#get ip for proxy
				if len(ip_list)<2:
					ip_list,ip_len = getip()

				name = sheet.cell_value(i,1)
				s_id = sheet.cell_value(i,2)
				print "这是对第%s个人信息的爬取"%i
				print name,int(s_id)
				# use the same session to crawl the data
				try:

					#to get good ip to crawl
					try:
						ip_list.remove(bad_ip)
					except Exception,e:
						print "no bad_ip"
					print ip_list


					#get random ip through the ip pool
					ip_len = len(ip_list) -1
					rand = random.randint(0,ip_len)
					ip = ip_list[rand][0]
					ip_port = ip_list[rand][1]
					proxies = {'http':'http://%s:%s'%(ip,ip_port),'https':'http://%s:%s'%(ip,ip_port)}

					#to get data
					post_data = {'__EVENTVALIDATION':info1,'__VIEWSTATE':info2,'__VIEWSTATEGENERATOR':info3,'Button1':'查询','TextBox1':s_id,'TextBox2':'','TextBox3':name,'TextBox4':vcode}
					r = session.post(url,data=post_data,headers=headers,proxies = proxies,timeout =4)
					#print r.text
			
					#print r.text
					#exam if it really take the exam
					pattern = re.compile(r'考试(.*?)请检查输入是否有误')
					print pattern.findall(r.text)
					try:
						if pattern.findall(r.text):
							continue
							print "%s于2017.12.16未参加考试"%name
							
						else:
							print "%s于2017.12.16参加了考试"%neam
					except Exception,e:
						print '  '


					if r.status_code != 200:
						print "验证码可能失效"
						bad_ip = [ip,ip_port]
						i =i-1
						continue

					soup = BeautifulSoup(r.text,'html.parser',from_encoding = 'utf-8')
					table = soup.find(id = 'DetailsView1')
					trs = table.find_all('tr')
					if (r.status_code != 200) and (trs ==None):
						print r.text

					#to get the photo
					if trs != None:
						try:
							photo = session.get('http://202.114.74.136/cet/getphoto.aspx',stream = True,headers = headers,proxies = proxies,timeout =4)
							with open('pic2/'+s_id+name+'.jpg','wb') as photo_f:
					 			photo_f.write(photo.content)
					 	except:
					 		print 'do not have photo'


					#to store data
					j = 0
					for tr in trs:
						tds = tr.find_all('td')

						try:
							sheet_store[sheet_index].write(k,j,tds[1].string)
							j=j+1
							
							print tds[1].string
						except Exception,e:
							print 'pic'


					k = k+1

				except requests.exceptions.ConnectTimeout:
					NETWORK_STATUS = False
					bad_ip = [ip,ip_port]
				except requests.exceptions.Timeout:
					REQUEST_TIMEOUT = True
					bad_ip = [ip,ip_port]
				except Exception,e:
					print name+'failed'
					print e
					traceback.print_exc()
	except Exception,e:
		f.save('info3.xlsx')



	f.save('info3.xlsx')


if __name__ == '__main__':
	crawl()





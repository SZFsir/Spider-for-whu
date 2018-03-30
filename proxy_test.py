import requests
import json
import sys


reload(sys)
sys.setdefaultencoding('utf8')

def getip():
	num = 100
	a = []

	ip_html = requests.get('http://127.0.0.1:8000/?type=0&count=%d'%num)
	ips = json.loads(ip_html.text)

	user_agent = 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'
	headers = {'User-Agent':user_agent}

	try:
		for i in range(0,num):
			ip = ips[i][0]
			ip_port = ips[i][1]
			print str(ip)+':'+str(ip_port)
			proxies  = {'http':'http://%s:%s'%(ip,ip_port),'https':'http://%s:%s'%(ip,ip_port)}
			url = 'http://www.whatismyip.com.tw/'
			try:
				r = requests.get(url,headers = headers,proxies = proxies,timeout= 2)
			except requests.exceptions.ConnectTimeout:
				NETWORK_STATUS = False
			except requests.exceptions.Timeout:
				REQUEST_TIMEOUT = True
			except Exception,e:
				print e
			print r
			#print r.content
			if r.status_code == 200:
				a.append([ip,ip_port])
	except Exception,e:
		print "error"

	print a

	return a,len(a)


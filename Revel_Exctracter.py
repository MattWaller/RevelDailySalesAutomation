try:

	import time
	from datetime import datetime, timedelta
	from bs4 import BeautifulSoup
	import Credentials
	from Driver import driver

	now = datetime.now()
	timezz = str(now.hour) + "." + str(now.minute)
	timef = float(timezz)


	year = str(now.year)
	month = str(now.month)
	today = str(now.day)
	yesterday = datetime.strftime(datetime.now() - timedelta(1), '%d')
	lmonth = datetime.strftime(datetime.now() - timedelta(1), '%m')
	lyear = datetime.strftime(datetime.now() - timedelta(1), '%Y')

	#if day / month less than 10
	if (int(today)<10):
		today = str('0'+ today)
	if (int(month)<10):
		month = str('0'+ month)


	timeV = "+03%3A00%3A00"
	to = '&range_to='
	us = '%2f'
	ds = '-'

	#defining links
	download = 'DOMAIN/reports/operations/data.csv?employee=&online_app=&online_app_type=&online_app_platform=&show_unpaid=1&show_irregular=1&range_from='+ lmonth + us + yesterday + us + lyear +  timeV + to + month + us + today + us + year + timeV
	downloadIng = 'DOMAIN/inventoryx/api/inventory-by-date/Ingredient/?export=1&offset=0&limit=50&product_class=&vendor=&display_based_on_primary_stock=&search_query=&range_from=' + lmonth +  us + yesterday + us + lyear +  timeV + to + month + us + today + us + year + timeV + '&date_from=' + lyear + ds + lmonth + ds + yesterday + '&date_to=' + year + ds + month + ds + today

	username = Credentials.login['consumer_username']
	password = Credentials.login['consumer_secret']



	driver.get(download)  


	# go to download file link

	time.sleep(3)

	# verify that login page is live.
	forgot = driver.find_element_by_xpath('//*[contains(text(),"Forgot")]').text
	try:
		if forgot == 'Forgot your password?':

			driver.find_element_by_xpath('//*[@id="id_username"]').send_keys(username)
			time.sleep(1)

			driver.find_element_by_xpath('//*[@id="id_password"]').send_keys(password)
			time.sleep(1)
			driver.find_element_by_xpath('//*[@id="form-login"]/fieldset/div[3]/input').click()


			time.sleep(5)
			print("file Grabbed!")
			

			driver.get(downloadIng)
			time.sleep(5)
			print("file Grabbed!")

			driver.quit()
			
	except Exception as e:
		raise e 




	# Convert & move Ingredients Report into Google Drive
	from datetime import datetime, timedelta
	import os.path
	import cloudconvert
	import requests
	import xlrd
	import csv
	import shutil

	SuffixThree = datetime.strftime(datetime.now() - timedelta(1), '%m_%d_%Y')

	ingredientsExist = os.path.isfile(r'G:\My Drive\RevelAccountingAutomation\RawData\IngredientsReport.csv')
	print(ingredientsExist)
	try:
		if ingredientsExist:
			os.remove(r"G:\My Drive\RevelAccountingAutomation\RawData\IngredientsReport.csv")
	except Exception as e:
		raise e
  
  #UTILIZE API FROM CLOUDCONVERT
	api = cloudconvert.Api('KEY HERE')
	process = api.convert({
	    'inputformat': 'xlsx',
	    'outputformat': 'csv',
	    'input': 'upload',
	    'file': open(r'C:\Users\Administrator\Downloads\Inventory_Summary_1_DOMAIN_' + SuffixThree + '_' + SuffixThree + '.xlsx', 'rb')
	})
	process.wait() # wait until conversion finished
	process.download(r"G:\My Drive\RevelAccountingAutomation\RawData\IngredientsReport.csv") # download output file

	rawDataExist = os.path.isfile(r'C:\Users\Administrator\Downloads\Inventory_Summary_1_DOMAIN_' + SuffixThree + '_' + SuffixThree + '.xlsx')
	try:
		if rawDataExist:
			os.remove(r'C:\Users\Administrator\Downloads\Inventory_Summary_1_DOMAIN_' + SuffixThree + '_' + SuffixThree + '.xlsx')
	except Exception as e:
		raise e

	# Relocate Operation Sales into Google Drive




	exists = os.path.isfile(r'G:\My Drive\RevelAccountingAutomation\RawData\OperationsReport.csv')
	try:
		if  exists:
		 	os.remove(r"G:\My Drive\RevelAccountingAutomation\RawData\OperationsReport.csv")
	except Exception as e:
		raise e

	SuffixOne = datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d_')
	SuffixTwo = datetime.strftime(datetime.now(), '%Y-%m-%d')

	FileName = 'C:\\Users\\Administrator\\Downloads\\'+'Operations_Report_DOMAIN_' + SuffixOne + SuffixTwo +'.csv'

	shutil.move(FileName, r"G:\My Drive\RevelAccountingAutomation\RawData\OperationsReport.csv")


	exit()

except Exception as e:
	import sys
	import datetime
	from Error_email import Error_email
	scriptname = sys.argv[0]
	timestamp = datetime.datetime.now()
	errorMsg = repr(e)
	Error_email(scriptname,timestamp,errorMsg)

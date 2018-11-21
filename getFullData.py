import sys
import os
import datetime
import requests
import xlsxwriter
import openpyxl as XL
import xml.etree.ElementTree as ET
import config

baseurl = 'https://api.ebay.com/ws/api.dll'
	
baseparams = {
	'Content-Type' : 'text/xml',
	'X-EBAY-API-COMPATIBILITY-LEVEL' : '1081',
	'X-EBAY-API-CALL-NAME' : 'GetSellerList',
	'X-EBAY-API-SITEID' : '0'
	}
	
authtoken = os.environ['key']
userid = os.environ['userid']

today = datetime.datetime.now()
future = today + datetime.timedelta(days=120)

pre = '{urn:ebay:apis:eBLBaseComponents}'

def getxml(page_number):
	return """
<?xml version="1.0" encoding="utf-8"?>
<GetSellerListRequest xmlns="urn:ebay:apis:eBLBaseComponents">
	<RequesterCredentials>
    <eBayAuthToken>{}</eBayAuthToken>
  </RequesterCredentials>
  <EndTimeFrom>{}</EndTimeFrom>
  <EndTimeTo>{}</EndTimeTo>
  <Pagination>
    <EntriesPerPage>200</EntriesPerPage>
    <PageNumber>{}</PageNumber>
  </Pagination>
  <UserID>{}</UserID>
  <DetailLevel>ReturnAll</DetailLevel>
  <OutputSelector>ItemID</OutputSelector>
  <OutputSelector>Title</OutputSelector>
  <OutputSelector>PaginationResult</OutputSelector>
  <OutputSelector>SellingStatus</OutputSelector>
  <OutputSelector>PrivateListing</OutputSelector>
</GetSellerListRequest>""".format(authtoken, today, future, str(page_number), userid)

listings = []

curPage = numPages = 1

while (curPage <= numPages):
	r = requests.post(baseurl, data=getxml(curPage), headers=baseparams)
	root = ET.fromstring(r.content)
	
	numPages = int(root.find(pre + 'PaginationResult').find(pre + 'TotalNumberOfPages').text)

	itemArr = root.find(pre + 'ItemArray')

	for eachItem in itemArr:
		itemid = eachItem.find(pre + 'ItemID').text
		title = eachItem.find(pre + 'Title').text
		price = eachItem.find(pre + 'SellingStatus').find(pre + 'CurrentPrice').text
		status = eachItem.find(pre + 'SellingStatus').find(pre + 'ListingStatus').text
		isPrivate = eachItem.find(pre + 'PrivateListing').text
		
		toadd = {
			'itemid' : itemid,
			'title' : title,
			'price' : price,
			'status' : status,
			'isPrivate' : isPrivate
			}
			
		listings.append(toadd)
		
	curPage = curPage + 1
	
wb = xlsxwriter.Workbook(f'{userid}.xlsx')
ws = wb.add_worksheet()

ws.write(0,0,'itemid')
ws.write(0,1,'title')
ws.write(0,2,'price')
ws.write(0,3,'status')
ws.write(0,4,'isPrivate')

row = 1

for eachListing in listings:
	ws.write(row,0,eachListing['itemid'])
	ws.write(row,1,eachListing['title'])
	ws.write(row,2,eachListing['price'])
	ws.write(row,3,eachListing['status'])
	ws.write(row,4,eachListing['isPrivate'])
	row = row + 1

wb.close()
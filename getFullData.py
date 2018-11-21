import os
import datetime
import requests
import boto3
import openpyxl as XL
import xml.etree.ElementTree as ET

baseurl = 'https://api.ebay.com/ws/api.dll'
	
baseparams = {
	'Content-Type' : 'text/xml',
	'X-EBAY-API-COMPATIBILITY-LEVEL' : '1081',
	'X-EBAY-API-CALL-NAME' : 'GetSellerList',
	'X-EBAY-API-SITEID' : '0'
	}
	
authtoken = os.environ['key']
userid = os.environ['userid']

bucket = boto3.resource('s3').Bucket('ebayreports')

today = datetime.datetime.now()
future = today + datetime.timedelta(days=120)

pre = '{urn:ebay:apis:eBLBaseComponents}'
fields = [
	'itemid',
	'title',
	'price',
	'status',
	'isPrivate'
	]

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

def main(event, context):

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
		
	wb = XL.Workbook()
	ws = wb.active
	
	ws.append(fields)
	
	row = 1
	
	for eachListing in listings:
		ws.append([eachListing[field] for field in fields])
		row = row + 1
	
	wb.save('/tmp/file.xlsx')
	
	bucket.Object(f"{userid} - Full Listing Details - {today}.xlsx").put(Body=open("/tmp/file.xlsx", 'rb'))
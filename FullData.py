import os
import datetime
import requests
import boto3
from threading import Thread
from threading import RLock
import openpyxl as XL
import xml.etree.ElementTree as ET

bucket = boto3.resource('s3').Bucket('ebayreports')
key = os.environ['key']
userid = os.environ['userid']

url = 'https://api.ebay.com/ws/api.dll'

getAllItemIdsParams = {
	'Content-Type' : 'text/xml',
	'X-EBAY-API-COMPATIBILITY-LEVEL' : '1081',
	'X-EBAY-API-CALL-NAME' : 'GetSellerList',
	'X-EBAY-API-SITEID' : '0'
	}
	
getAllItemsParams = {
	'Content-Type' : 'text/xml',
	'X-EBAY-API-COMPATIBILITY-LEVEL' : '1081',
	'X-EBAY-API-CALL-NAME' : 'GetItem',
	'X-EBAY-API-SITEID' : '0'
	}

today = datetime.datetime.now()
future = today + datetime.timedelta(days=120)

def P(str):
	return f"{{urn:ebay:apis:eBLBaseComponents}}{str}"
	
def getValueString(name, item):
	itemspecifics = item.find(P('ItemSpecifics'))
	returnString = ''
	
	for each in itemspecifics:
		if (each.find(P('Name')).text == name):
			allValues = each.findall(P('Value'))
			numValues = len(allValues)
			
			if (numValues > 1):
				i = 0
				while (i < numValues):
					if (i != (numValues - 1)):
						returnString = allValues[i].text + ', ' + returnString
					else:
						returnString = returnString + allValues[i].text
					i = i + 1
			else:
				returnString = allValues[0].text
				break
	
	return returnString

def getAllItemIdsXML(pagenum):
	return f"""
<?xml version="1.0" encoding="utf-8"?>
<GetSellerListRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <RequesterCredentials>  
  <eBayAuthToken>{key}</eBayAuthToken>
  </RequesterCredentials>
  <EndTimeFrom>{today}</EndTimeFrom>
  <EndTimeTo>{future}</EndTimeTo>
  <Pagination>
    <EntriesPerPage>200</EntriesPerPage>
    <PageNumber>{pagenum}</PageNumber>
  </Pagination>
  <OutputSelector>ItemID</OutputSelector>
  <OutputSelector>PaginationResult</OutputSelector>
</GetSellerListRequest>
"""

def getAllItemsXML(itemid):
	return f"""
<?xml version="1.0" encoding="utf-8"?>
<GetItemRequest xmlns="urn:ebay:apis:eBLBaseComponents">
  <RequesterCredentials>
    <eBayAuthToken>{key}</eBayAuthToken>
  </RequesterCredentials>
  <ItemID>{itemid}</ItemID>
  <IncludeItemSpecifics>True</IncludeItemSpecifics>
  <DetailLevel>ReturnAll</DetailLevel>
</GetItemRequest>
"""

def getAllItemIds():
	cur_page = 1
	tot_pages = 1

	itemids = []
	
	while (cur_page <= tot_pages):
		r = requests.post(url, data=getAllItemIdsXML(cur_page), headers=getAllItemIdsParams)
	
		root = ET.fromstring(r.content)
	
		tot_pages = int(root.find(P('PaginationResult')).find(P('TotalNumberOfPages')).text)
	
		itemArr = root.find(P('ItemArray'))
		
		print("Looping")
	
		for eachItem in itemArr:
			itemid = eachItem.find(P('ItemID')).text
			itemids.append(itemid)
		
		# Append each half of the list of 200, we want to create 1 thread per 100 listings
		allItemIds.append(itemids[:len(itemids)//2])
		allItemIds.append(itemids[len(itemids)//2:])
		
		itemids = []
		cur_page = cur_page + 1
	

def getItems(listOfItemIds):
	for eachItemId in listOfItemIds:
		r = requests.post(url, data=getAllItemsXML(eachItemId), headers=getAllItemsParams)
		root = ET.fromstring(r.content)
		item = root.find(P('Item'))

		Action = ''
		CustomLabel = ''

		CategoryID = ''
		try:
			CategoryID = item.find(P('PrimaryCategory')).find(P('CategoryID')).text
		except:
			pass
		
		StoreCategoryID = ''
		try:
			StoreCategoryID = item.find(P('Storefront')).find(P('StoreCategoryID')).text
		except:
			pass
		
		Title = ''
		try:
			Title = item.find(P('Title')).text
		except:
			pass

		Subtitle = ''
		Relationship = ''
		RelationshipDetails = ''

		ConditionID = ''
		try:
			ConditionID = item.find(P('ConditionID')).text
		except:
			pass

		Brand = getValueString('Brand', item)	
		PartType = getValueString('Part Type', item)
		ManufacturerPartNumber = getValueString('Manufacturer Part Number', item)
		InterchangePartNumber = getValueString('Interchange Part Number', item)
		OtherPartNumber = getValueString('Other Part Number', item)
		PlacementOnVehicle = getValueString('Placement on Vehicle', item)
		Warranty = getValueString('Warranty', item)
		CustomBundle = getValueString('Custom Bundle', item)
		FitmentType = getValueString('Fitment Type', item)
		IncludedHardware = getValueString('Included Hardware', item)

		BundleDescription = ''

		Greasable = getValueString('Greasable', item)
		ModifiedItem = getValueString('Modified Item', item)
		Adjustable = getValueString('Adjustable', item)

		ModificationDescription = ''

		NonDomesticProduct = getValueString('Non-Domestic Product', item)

		ApplicableRegions = ''
		DropLength = ''
		GasChargedShock = ''
		CaliforniaProp65Warning = ''

		CountryRegionOfManufacture = getValueString('Country/Region of Manufacture', item)

		SupersededPartNumber = ''
		PicURL = ''
		try:
			PicURL = item.find(P('PictureDetails')).find(P('GalleryURL')).text
		except:
			pass

		GalleryType = ''
		try:
			GalleryType = item.find(P('PictureDetails')).find(P('GalleryType')).text
		except:
			pass
			
		Description = ''
		try:
			Description = item.find(P('Description')).text
		except:
			pass

		Format = ''

		Duration = ''
		try:
			Duration = item.find(P('ListingDuration')).text
		except:
			pass
			
		StartPrice = ''
		try:
			StartPrice = item.find(P('StartPrice')).text
		except:
			pass
			
		BuyItNowPrice = ''
		try:
			BuyItNowPrice = item.find(P('BuyItNowPrice')).text
		except:
			pass
			
		Quantity = ''
		try:
			Quantity = item.find(P('Quantity')).text
		except:
			pass

		PayPalAccepted = ''
		PayPalEmailAddress = ''
		ImmediatePayRequired = ''
		PaymentInstructions = ''
		Location = ''
		
		ShippingType = ''
		try:
			ShippingType = item.find(P('ShippingDetails')).find(P('ShippingType')).text
		except:
			pass
		
		ShippingService1Option = ''
		try:
			ShippingService1Option = item.find(P('ShippingDetails')).find(P('ShippingServiceOptions')).find(P('ShippingService')).text
		except:
			pass
			
		ShippingService1Cost = ''	
		try:
			ShippingService1Cost = item.find(P('ShippingDetails')).find(P('ShippingServiceOptions')).find(P('ShippingServiceCost')).text
		except:
			pass

		ShippingService2Option = ''
		ShippingService2Cost = ''

		DispatchTimeMax = ''
		try:
			DispatchTimeMax = item.find(P('DispatchTimeMax')).text
		except:
			pass

		PromotionalShippingDiscount = ''
		ShippingDiscountProfileID = ''
		DomesticRateTable = ''

		ReturnsAcceptedOption = ''
		try:
			ReturnsAcceptedOption = item.find(P('ReturnPolicy')).find(P('ReturnsAcceptedOption')).text
		except:
			pass
			
		ReturnsWithinOption = ''
		try:
			ReturnsWithinOption = item.find(P('ReturnPolicy')).find(P('ReturnsWithinOption')).text
		except:
			pass

		RefundOption = ''
		ShippingCostPaidByOption = ''
		AdditionalDetails = ''
		UseTaxTable = ''

		toAdd = {
			'itemid' : eachItemId,
			'*Action(SiteID=eBayMotors|Country=US|Currency=USD|Version=745|CC=UTF-8)' : Action,
			'CustomLabel' : CustomLabel,
			'*Category' : CategoryID,
			'StoreCategory' : StoreCategoryID,
			'*Title' : Title,
			'Subtitle' : Subtitle,
			'Relationship' : Relationship,
			'RelationshipDetails' : RelationshipDetails,
			'*ConditionID' : ConditionID,
			'*C:Brand' : Brand,
			'C:Part Type' : PartType,
			'*C:Manufacturer Part Number' : ManufacturerPartNumber,
			'C:Interchange Part Number' : InterchangePartNumber,
			'C:Other Part Number' : OtherPartNumber,
			'C:Placement on Vehicle' : PlacementOnVehicle,
			'C:Warranty' : Warranty,
			'C:Custom Bundle' : CustomBundle,
			'C:Fitment Type' : FitmentType,
			'C:Included Hardware' : IncludedHardware,
			'C:Bundle Description' : BundleDescription,
			'C:Greasable or Sealed' : Greasable,
			'C:Modified Item' : ModifiedItem,
			'C:Adjustable' : Adjustable,
			'C:Modification Description' : ModificationDescription,
			'C:Non-Domestic Product' : NonDomesticProduct,
			'C:Applicable Regions' : ApplicableRegions,
			'C:Drop Length' : DropLength,
			'C:Gas Charged Shock' : GasChargedShock,
			'C:California Prop 65 Warning' : CaliforniaProp65Warning,
			'C:Country/Region of Manufacture' : CountryRegionOfManufacture,
			'C:Superseded Part Number' : SupersededPartNumber,
			'' : '',
			'PicURL' : PicURL,
			'GalleryType' : GalleryType,
			'*Description' : Description,
			'*Format' : Format,
			'*Duration' : Duration,
			'*StartPrice' : StartPrice,
			'BuyItNowPrice' : BuyItNowPrice,
			'*Quantity' : Quantity,
			'PayPalAccepted' : PayPalAccepted,
			'PayPalEmailAddress' : PayPalEmailAddress,
			'ImmediatePayRequired' : ImmediatePayRequired,
			'PaymentInstructions' : PaymentInstructions,
			'*Location' : Location,
			'ShippingType' : ShippingType,
			'ShippingService-1:Option' : ShippingService1Option,
			'ShippingService-1:Cost' : ShippingService1Cost,
			'ShippingService-2:Option' : ShippingService2Option,
			'ShippingService-2:Cost' : ShippingService2Cost,
			'*DispatchTimeMax' : DispatchTimeMax,
			'PromotionalShippingDiscount' : PromotionalShippingDiscount,
			'ShippingDiscountProfileID' : ShippingDiscountProfileID,
			'DomesticRateTable' : DomesticRateTable,
			'*ReturnsAcceptedOption' : ReturnsAcceptedOption,
			'ReturnsWithinOption' : ReturnsWithinOption,
			'RefundOption' : RefundOption,
			'ShippingCostPaidByOption' : ShippingCostPaidByOption,
			'AdditionalDetails' : AdditionalDetails,
			'UseTaxTable' : UseTaxTable}
			
		wbLock.acquire()
		
		try:
			print('writing to file')
			outws.append([value for key, value in toAdd.items()])
		finally:
			wbLock.release()
			
outwb = XL.Workbook()
outws = outwb.active

wbLock = RLock()

# Write headers to excel file
outws.append([
	'itemid',
	'*Action(SiteID=eBayMotors|Country=US|Currency=USD|Version=745|CC=UTF-8)',
	'CustomLabel',
	'*Category',
	'StoreCategory',
	'*Title',
	'Subtitle',
	'Relationship',
	'RelationshipDetails',
	'*ConditionID',
	'*C:Brand',
	'C:Part Type',
	'*C:Manufacturer Part Number',
	'C:Interchange Part Number',
	'C:Other Part Number',
	'C:Placement on Vehicle',
	'C:Warranty',
	'C:Custom Bundle',
	'C:Fitment Type',
	'C:Included Hardware',
	'C:Bundle Description',
	'C:Greasable or Sealed',
	'C:Modified Item',
	'C:Adjustable',
	'C:Modification Description',
	'C:Non-Domestic Product',
	'C:Applicable Regions',
	'C:Drop Length',
	'C:Gas Charged Shock',
	'C:California Prop 65 Warning',
	'C:Country/Region of Manufacture',
	'C:Superseded Part Number',
	'',
	'PicURL',
	'GalleryType',
	'*Description',
	'*Format',
	'*Duration',
	'*StartPrice',
	'BuyItNowPrice',
	'*Quantity',
	'PayPalAccepted',
	'PayPalEmailAddress',
	'ImmediatePayRequired',
	'PaymentInstructions',
	'*Location',
	'ShippingType',
	'ShippingService-1:Option',
	'ShippingService-1:Cost',
	'ShippingService-2:Option',
	'ShippingService-2:Cost',
	'*DispatchTimeMax',
	'PromotionalShippingDiscount',
	'ShippingDiscountProfileID',
	'DomesticRateTable',
	'*ReturnsAcceptedOption',
	'ReturnsWithinOption',
	'RefundOption',
	'ShippingCostPaidByOption',
	'AdditionalDetails',
	'UseTaxTable'])

threads = []
allItemIds = []

def main(event, context):
	getAllItemIds()
		
	for listOfItemIds in allItemIds:
		print("creating thread")
		t = Thread(target=getItems, args=(listOfItemIds,))
		threads.append(t)
		
	for t in threads:
		t.start()
		
	for t in threads:
		t.join()
		
	outwb.save('/tmp/out.xlsx')
	bucket.Object(f"{userid} - Full Listing Details - {today}.xlsx").put(Body=open("/tmp/out.xlsx", 'rb'))
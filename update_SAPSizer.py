## Include those before using: pip3 install openpyxl xlsxwriter xlrd
from openpyxl import load_workbook
import urllib3, json
from datetime import date
  
# creating the date object of today's date
todays_date = date.today()
in_fname = "SAPs.xlsx"
out_fname = "SAP Sizer BRA - "+str(todays_date.year)+str(todays_date.month)+str(todays_date.day)+".xlsx"

http = urllib3.PoolManager()
run = 1
# url = "https://prices.azure.com/api/retail/prices?$filter=serviceName eq 'Virtual Machines' and endswith(productName,'Windows') eq false and meterName eq 'E32as v5' and ((currencyCode eq 'USD') or (currencyCode eq 'BRL')) and ((location eq 'Brazil South') or (location eq 'US East')) "
alldata = []

wb = load_workbook(in_fname)
sheet = wb.worksheets[0]
sheetprice = wb.worksheets[1]

cell = sheet['A1']
count=1
print ('Reading from SAPs.xlsx columns :'+cell.value)

while cell.value != None:
    count += 1
    cell = sheet['A'+str(count)]
    if cell.value != None:
        print ('Processing line: '+str(count)+' for VM: '+ cell.value)
        run = 1
        url = "https://prices.azure.com/api/retail/prices?$filter=serviceName eq 'Virtual Machines' and endswith(productName,'Windows') eq false and meterName eq '"+cell.value+"' and ((currencyCode eq 'USD') or (currencyCode eq 'BRL')) and ((location eq 'BR South') or (location eq 'US East'))"
        while run: 
            ##current url
            # print (url)
            r = http.request('GET', url)
            data = json.loads(r.data)
            for data_item in data["Items"]: 
                # print (data_item)
                sheetprice['A'+str(count)].value = sheet['A'+str(count)].value
                sheetprice['B'+str(count)].value = sheet['B'+str(count)].value
                sheetprice['C'+str(count)].value = sheet['C'+str(count)].value
                sheetprice['D'+str(count)].value = sheet['D'+str(count)].value
                sheetprice['E'+str(count)].value = sheet['E'+str(count)].value
                if data_item["currencyCode"] == "USD":
                    if "reservationTerm" in data_item: ##HAS RI
                        if data_item["reservationTerm"] == "3 Years": ## is 3 Yr
                            if data_item["armRegionName"] == "eastus":
                                sheetprice['G'+str(count)].value = data_item["retailPrice"]/36
                            else:
                                sheetprice['I'+str(count)].value = data_item["retailPrice"]/36
                    else: ##PAYG
                        if data_item["armRegionName"] == "eastus":
                            sheetprice['F'+str(count)].value = data_item["retailPrice"]
                        else: 
                            sheetprice['H'+str(count)].value = data_item["retailPrice"]
            #if paging is required, set url to nextpagelink
            if data['NextPageLink'] != None:
                url = data['NextPageLink']
            else:
                #otherwise finish loop 
                run = 0


wb.save(out_fname)
print ("SAVED FILE ", out_fname)
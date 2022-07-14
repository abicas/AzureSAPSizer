## Include those before using: pip3 install openpyxl xlsxwriter xlrd
from openpyxl import load_workbook
import urllib3, json
from datetime import date
  
# creating the date object of today's date
todays_date = date.today()

#setting up the filenames
in_fname = "SAPs.xlsx"
out_fname = "SAP Sizer BRA - "+str(todays_date.year)+str(todays_date.month)+str(todays_date.day)+".xlsx"

http = urllib3.PoolManager()
run = 1

#opening SAPs.xlsx
wb = load_workbook(in_fname)
sheet = wb.worksheets[0]
sheetprice = wb.worksheets[1]

cell = sheet['A1']
print ('Reading from SAPs.xlsx columns :'+cell.value)
count = 0 

##for each line with a certified VM name 
while cell.value != None:
    count = count + 1
    cell = sheet['A'+str(count)]
    ## if cell is not the last one 
    if cell.value != None:
        print ('Processing line: '+str(count)+' for VM: '+ cell.value)
        run = 1
        ## build odata URL for pricelist with VM name
        url = "https://prices.azure.com/api/retail/prices?$filter=serviceName eq 'Virtual Machines' and endswith(productName,'Windows') eq false and meterName eq '"+cell.value+"' and ((currencyCode eq 'USD') or (currencyCode eq 'BRL')) and ((location eq 'BR South') or (location eq 'US East'))"
        while run: 
            ## get the json
            r = http.request('GET', url)
            data = json.loads(r.data)
            ## for every item in the json query
            for data_item in data["Items"]: 
                ## copy data from SAPs sheet
                sheetprice['A'+str(count)].value = sheet['A'+str(count)].value
                sheetprice['B'+str(count)].value = sheet['B'+str(count)].value
                sheetprice['C'+str(count)].value = sheet['C'+str(count)].value
                sheetprice['D'+str(count)].value = sheet['D'+str(count)].value
                sheetprice['E'+str(count)].value = sheet['E'+str(count)].value
                ## for now just USD is supported, making sure we are not getting additional currencies
                if data_item["currencyCode"] == "USD":
                    ## If it is an RI Item
                    if "reservationTerm" in data_item: 
                        ## If this RI is for 3 Years 
                        if data_item["reservationTerm"] == "3 Years": ## is 3 Yr
                            ## If in US East populate the column G  (RI3YR price in JSON is for 36 months, so dividing by 36 to get monthly prices)
                            if data_item["armRegionName"] == "eastus":
                                sheetprice['G'+str(count)].value = data_item["retailPrice"]/36
                            else:
                                ## Otherwise populate column I (RI3YR price in JSON is for 36 months, so dividing by 36 to get monthly prices)
                                sheetprice['I'+str(count)].value = data_item["retailPrice"]/36
                    else: ## Not RI, PAYG Then
                        ## If in US, PAYG prices are hourly based in JSON
                        if data_item["armRegionName"] == "eastus":
                            sheetprice['F'+str(count)].value = data_item["retailPrice"]
                        else: 
                            ## If not in US, PAYG prices are hourly based in JSON
                            sheetprice['H'+str(count)].value = data_item["retailPrice"]
            #if paging is required, there will be content for Next PageLink. set url to nextpagelink and loop
            if data['NextPageLink'] != None:
                url = data['NextPageLink']
            else:
                #otherwise finish loop 
                run = 0

## Save the file with DATE on the name
wb.save(out_fname)
print ("SAVED FILE ", out_fname)
## Alright, go configure some SAP VMs ! 
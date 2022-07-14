# AzureSAPSizer

This is a quick Sizer for SAP on Azure VMs based on the SAPs and memory requirements. 
It aims to select the cheapest option for a given set of requirements.

It doesn't check against specific limitations like OS versions or SAP modules compatibility. 

## Instructions 

SAPs.xlsx is the calculator template - DO NOT CHANGE IT

run python3 update_SAPSizer.py and the script will go thru the SAPs.xlsx and gather the prices for RI and PAYG for all the VMs listed with SAPs ratings on the "SAPs" sheet, populating the Pricelist sheet. 

At the end the script will create a calculator with the updated prices calles "SAP Sizer BRA - YYYYMMDD.xlsx". 

This is the one that should be used. 

## Usage

1. Open the "SAP Sizer BRA - YYYYMMDD.xlsx" generated file. 

2. Go to the "Sizer" sheet. 

3. Fill the columns under CURRENT ENVIRONMENT (it required at least SAPs and RAM in GB)

4. Check on the right under SUGGESTED AZURE VMs for certified VMs. 

5. Check prices for PAYG and RI for quick estimate. 

![alt text](https://github.com/abicas/AzureSAPSizer/blob/[branch]/image.jpg?raw=true)

#! python3
# Filename:			Stryker_repairs.py
# Author:			P. Rivet
# Date:				17/07/31
# Description:		Combines individual repair files into a master Excel spreadsheet

import sys, re, os, openpyxl, datetime, csv
#import pandas as pd
#import numpy as np

def repair_aggregator():
	#create database folder
	os.makedirs('repair_database',exist_ok=True)
	os.chdir('repair_database')
	date_now=str(datetime.date.today())
	print('Today\'s date is: '+date_now)
	
	#loop through directory tree to check for db file
	dbExists=False
	for xlsxFilename in os.listdir('.'):
		if os.path.isfile('Stryker_Repairs_Master.xlsx'):
			print('DATABASE FILE DETECTED')
			dbExists=True
			break
		else:
			continue
	#create database file if it was not detected previously
	if dbExists==False:
		print('CREATING DATABASE FILE...')
		db1=openpyxl.Workbook()
		sheet1=db1.get_active_sheet()
		sheet1.title='PCB INFORMATION'
		db1.save('Stryker_Repairs_Master.xlsx')
	
	#scan for column titles
	wb1=openpyxl.load_workbook('Stryker_Repairs_Master.xlsx')
	print(type(wb1))
	sheet1=wb1.get_sheet_by_name('PCB INFORMATION')
	print('Loading sheet '+str(sheet1))
	#write appropriate labels if blank
	if (sheet1['A1'].value)==None:
		print('Writing column titles now! Yahoo')
		sheet1['A1'].value='Item No.'
		sheet1['B1'].value='QTY'
		sheet1['C1'].value='Part/Component Name'
		sheet1['D1'].value='Part No.'
		sheet1['E1'].value='Part ID/Traceability'
		sheet1['F1'].value='Received Observations'
		sheet1['G1'].value='Supplier Test Results'
		sheet1['H1'].value='Quick Fault'
		sheet1['I1'].value='No. Issues'
		sheet1['J1'].value='Resp. Code'
		sheet1['K1'].value='Reas. Code'
		sheet1['L1'].value='Warranty'
		sheet1['M1'].value='Report Date'
		sheet1['N1'].value='RMA No.'
		sheet1['O1'].value='SML RMAR No.'
		sheet1['P1'].value='Date of Manufacturing'
		sheet1['Q1'].value='ECN#'
		sheet1['R1'].value='OES Fault Classification Code'
		sheet1['S1'].value='Filename'
		sheet1['T1'].value='Date Recorded'
		print('Titles updated.')	
	wb1.save('Stryker_Repairs_Master.xlsx')
	
	fileNames=[]
	#loop through directory tree to check for other any other spreadsheets
	for xlsxFilename in os.listdir('.'):
		if (not xlsxFilename.endswith('.xlsx')):
			continue #skip other file types
		#print (xlsxFilename+' found. Establishing parameters')
		if (str(xlsxFilename)!='Stryker_Repairs_Master.xlsx'):
			fileNames.append(str(xlsxFilename))
	
	#display all detected filenames
	for i in range (0,len(fileNames)):
		print(str(fileNames[i]))
	
	#obtain critical info using regex
	#https://regex101.com/r/u8km0c/2
	regex = re.compile(r"(((\d){1,7})([A-Z]{1})((\d){2})((\d){3,4})([/]{1})([A-Z]{1}\d{1,3}))")
	test_str1 = ("38023Z05512/Y165")
	test1=regex.search(test_str1)
	print('TEST REGEX: '+str(test1))
	
	
	#go through each detected spreadsheet file to check for data 
	for j in range (0,len(fileNames)):
		currentFilename=fileNames[j]
		wbj=openpyxl.load_workbook(currentFilename)
		shj=wbj.get_active_sheet()
		print(type(wbj))
		print('Checking data in '+currentFilename+' ...')
		
		#identify extra info
		reportDate=shj.cell('D4').value
		rmaNo=shj.cell('G10').value
		smlNo=shj.cell('K10').value
		print(str(reportDate)+str(rmaNo)+str(smlNo))
		
		#loop through each row starting from A13 and store the values in a list 
		#this is problematic for sheets that jump past the 13th row --> possible looping solution
		data=[]
		
		try:
			for k in range (13,shj.max_row+1):
				subdata=[]
				for l in range (1,shj.max_column+1):
					subdata.append(shj.cell(row=k,column=l).value)
				print(subdata)
				partID=subdata[4]
				#print(str(partID))
				try:
					analysis=regex.search(partID)
					#print(str(analysis))
				except TypeError:
					print('D.N.E.')
				
				#=IF('http://oes-autodeskx64/Quality/Non-Conforming Cage/Non-Conforming Product 2012-2013/[Non-Conforming Product.xlsx]In-House (2012-2013)'!$P6="YES",'http://oes-autodeskx64/Quality/Non-Conforming Cage/Non-Conforming Product 2012-2013/[Non-Conforming Product.xlsx]In-House (2012-2013)'!$H6,"-")
				
				#determine the exact year and month of manufacture
				try:
					mth1=str(analysis.group(5))
					if str(analysis.group(4))=='Z':
						yr1='15'
					elif str(analysis.group(4))=='A':
						yr1='16'
					elif str(analysis.group(4))=='B':
						yr1='17'
					else:
						yr1='XX'
					ecn=str(analysis.group(10))
					dom=yr1+'-'+mth1
					print('Date of Manufacturing: '+dom)
					print('ECN #: '+ecn)
					oesFault=' '
					
					#record extra info
					subdata.append(str(reportDate))
					subdata.append(str(rmaNo))
					subdata.append(str(smlNo))
					subdata.append(str(dom))
					subdata.append(str(ecn))
					subdata.append(str(oesFault))
					subdata.append(str(currentFilename))
					subdata.append(date_now)
					
					data.append(subdata)
				except TypeError:
					print('D.N.E.')
		except TypeError:
			print('D.N.E.')
		#print(data)
	
		#write to MASTER
		#potentially add comparison feature to avoid duplication
		count1=0
		count2=0
		match=False
		for m in range (0,len(data)):
			rowNow=sheet1.max_row+1
			for n in range (0,20):
				currentVal=data[m][n]
				count1+=1
				print(currentVal)
				sheet1.cell(row=rowNow,column=n+1).value=currentVal
				#recognize additional columns here??
				
			#print(match)
			print (count1)
		#print(count2)
			wb1.save('Stryker_Repairs_Master.xlsx')
		wb1.save('Stryker_Repairs_Master.xlsx')
		
repair_aggregator()

#pandas?
def repair_aggregator_2():
	os.makedirs('repair_database_2',exist_ok=True)
	
	
#! python3
#Filename:			AOI_reformatter.py
#Author:			P. Rivet
#Date:				17/08/16
#Description:		Program to re-format an ORCAD txt file to a CSV
#					Output columns: X,Y,angle,linefeed,Ref,Stamp

import sys, re, os, csv

#nested list length function
def recursive_len(item):
    if type(item) == list:
        return sum(recursive_len(subitem) for subitem in item)
    else:
        return 1
		
#column swap function
def swap_columns(your_list, pos1, pos2):
    your_list[pos1], your_list[pos2] = your_list[pos2], your_list[pos1]
	
#function to reformat all CSV files in the folder to be useable by the AOI machine
#the nLines argument specifies the number of lines to be removed (top-down)
def csvReformatter(nLines):
	#creater folder for output files
	os.makedirs('csvReformatted',exist_ok=True)
	#loop throughout directory tree
	for csvFilename in os.listdir('.'):
		#skip other file types
		if not csvFilename.endswith('.csv'):
			continue
		print('Now removing '+str(nLines)+' lines from ' +csvFilename+'...')
		#read in file
		csvRows=[]
		csvFileObj=open(csvFilename,encoding='utf8')
		readerObj=csv.reader(csvFileObj)
		#skip the specified number of lines
		for row in readerObj:
			if readerObj.line_num<=nLines:
				continue#skip nLines rows
			csvRows.append(row)
		
		gogogo=True
		tick=0
		tock=0
		clock=0
		
		#count number of m's
		for i in range (0,len(csvRows)):
			tick=csvRows[i].count('m')
			print(str(tick)+' space(s) counted in row '+str(i))
			clock+=tick
		#print('Total occurences of m: '+str(clock))
		
		#loop through to remove m's
		#as long as the part names are all capitalized this should function smoothly once imported
		while (gogogo):
			for i in range(0,len(csvRows)):
				#print(csvRows[i])
				#remove m's from coordiante columns
				if ('m' in csvRows[i]):
					csvRows[i].remove('m')
					tock+=1
				#for j in range (0,len(csvRows[i])):
					#print(csvRows[i][j])
					#print(csvRows[i][j]=='m')
			if((clock-tock)==0):
				gogogo=False
			
		#remove unnecessary spaces & tabs 
		onward=True
		count1=0
		count2=0
		total=0
		#count number of spaces
		for j in range (0,len(csvRows)):
			count1=csvRows[j].count(' ')
			print(str(count1)+' space(s) counted in row '+str(j))
			total+=count1
		print('Total number of spaces: '+str(total))
		
		#continuously loop through to verify space elimination
		while (onward):
			for i in range (0,len(csvRows)):
				if (' ' in csvRows[i]):
					csvRows[i].remove(' ')
					count2+=1
			print(csvRows[i])
			print('Spaces removed: '+str(count2))	
			print("Spaces remaining: "+str(total-count2))
			if ((total-count2)==0):
				onward=False
		
		#terminate redundant commas using regexes
		#(([A-Z],){1,2}(\d,){1,3})(.{1,101})((\d,){3,4}\.,\d,\d,)((\d,){3,4}\.,\d,\d,)((\d,\d,\d)|(\d,\d)|(\d))
		#https://regex101.com/r/ImpQL1/1
		#n_c=(n_t/2)-1
		#VERSION WITH NO COMMAS PERFORMS BETTER
		#(([A-Z]){1,2}(\d){1,2})(.{1,101}?)((\d){3,4}\.\d\d)((\d){3,4}\.\d\d)((\d\d\d)|(\d\d)|(\d))
		#https://regex101.com/r/BVCpiZ/4
		#Troubleshooting:
		#https://regex101.com/r/H9d0Rp/1
		space_regex=re.compile(r"(([A-Z],){1,2}(\d,){1,2})(.{1,101})((\d,){3,4}\.,\d,\d,)((\d,){3,4}\.,\d,\d,)((\d,\d,\d)|(\d,\d)|(\d))")
		#rewrite without commas :]
		time_regex=re.compile(r"(([A-Z]){1,2}(\d){1,2})(.{1,101}?)((\d){1,3}?\.\d\d)((\d){1,3}?\.\d\d)((\d\d\d)|(\d\d)|(\d))")
		#change setting in ORCAD to force 3 digits & 2 decimal places for coordinates, and dimensions in millimetres
		#testing
		str1="Q,9,M,M,B,T,3,9,0,6,L,T,1,M,M,B,T,3,9,0,6,L,S,M,/,S,O,T,2,3,_,1,2,3,1,9,7,9,.,5,0,1,9,4,8,.,0,0,2,7,0"
		test1=space_regex.search(str1)
		test2=time_regex.search(str(csvRows[5]))
		print(str(csvRows[5]))
		print('Test regex:' +str(test2))
		
		
		#NEVER GIVE UP
		#go through the entire list of rows & re-arrange columns 
		for k in range (0,len(csvRows)):
			temp1=[]
			temp2=[]
			#parsing the list of lists values into proper strings
			print(str(csvRows[k]))
			parse1=''.join(csvRows[k])#join the list entries into one string
			print(parse1)
			test1=time_regex.search(parse1)
			print('Test regex:' +str(test1))
			#loop through the regex groupings and attach each to a new list
			try:
				for i in range (0,13):
					temp1.append(test1.group(i))
			except TypeError:
				print('0')
			print(temp1)
			temp2.append(temp1[1])
			temp2.append(temp1[4])
			temp2.append(temp1[5])
			temp2.append(temp1[7])
			temp2.append(temp1[9])
			temp2.append('1')
			#swap columns to correct order (x,y,theta,z=1,ref,stamp)
			swap_columns(temp2,0,2)
			swap_columns(temp2,1,3)
			swap_columns(temp2,2,4)
			swap_columns(temp2,3,5)
			print(temp2)
			csvRows[k]=temp2
			
		#section to update stamps w/ BIN #s
		
		csvFileObj.close()   
		#write output
		csvFileObj=open(os.path.join('csvReformatted',csvFilename),'w',newline='')
		csvWriter=csv.writer(csvFileObj)
		for row in csvRows:
			csvWriter.writerow(row)
		csvFileObj.close()
		print (csvFilename+' has been written successfully.')
		
#function to convert a txt file from an ORCAD schematic to a CSV file 
def txtToCsv():
	contentList=[]
	print('Hello, welcome to the txt to csv converter.')
	#loop while looking for txt files in directory
	i=0
	for txtFilename in os.listdir('.'):
		if not (txtFilename.endswith('.txt') or txtFilename.endswith('.TXT')):
			continue
		print(txtFilename+' detected for conversion')
		with open(txtFilename,'r') as f:
			contentList=[line.strip() for line in f]
			
		i+=1
		
		txtFilename = re.sub('(\.txt$)|(\.TXT$)', '', txtFilename)#obliterate incorrect extension for output
		csvFile=open(txtFilename+'.csv','w',newline='')
		csvFileWriter=csv.writer(csvFile)
		#loop and write each row
		for k in range (0,len(contentList)):
			csvFileWriter.writerow(contentList[k])
		csvFile.close()
	
txtToCsv()
csvReformatter(11)
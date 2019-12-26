########################################################################
# Program name: PythonRegularExpressions
# Author: Kevin Hill
# Date:  12/25/2019
# Description: Main script file for the PythonRegularExpression project 
# to read in an excel document with running log stats into a dataframe
# with pandas and capture the mileage out of the english using regular
# expressions. The output is then written to a test.txt file in a 
# format as described:
# 
# mileage
# details (all characters in cell)
# date
#
# After parsing, the highest and lowest mileage is printed to the 
# console as well as length of data, total number of runs parsed
# and first row of mileages.
#
# Note: the excel document has a max of 20 columns with odd rows as 
# dates and even rows as cells with running data input in English.
# For example to capture a 5 mile run enter in the cell:
# run 5mi
# 
# However, to guarantee the mileage is captured for a given cell
# include the mileage in parentheses somewhere. For example:
# (8mi)
#
# Using the above markup the cell can contain additional characters 
# and still capture the toal mileage.
########################################################################
import re
import pandas as pd

#arrays to store data and index and row_val variables
row_val = 0
index=0
dates = []
detail = []
data = []

df = pd.read_excel('FitnessPlan_V3.xlsx', header=None)
for idx, row in df.iterrows():
    row_val=row_val+1
    if row_val%2==0:
        for x in range(0,20):
            str_val = str(row[x])
            #put detail values in array from respective cell's characters
            detail.append(str_val)
            #format: half
            matchObjHalf = re.match( r'half', str_val, re.X)
            #format: mara
            matchObjMara = re.match( r'mara', str_val, re.X)
            #format: <additional characters> 5/8mi <additional characters>
            patternSlashDouble = re.compile( r'.(\d*\.?\d+)/(\d*\.?\d+)mi\Z', re.IGNORECASE)
            searchObjSlashDouble = re.search(patternSlashDouble, str_val)
            #format:  <additional characters> (8mi) <additional characters>
            patternRunPar = re.compile(r'\((\d*\.?\d+)mi\)', re.IGNORECASE)
            searchObjRunPar = re.search(patternRunPar, str_val)
            #format: <additional characters> run 5mi
            patternRunStd = re.compile(r'[Rr][au]n\s+(\d*\.?\d+)mi\Z', re.IGNORECASE)
            searchObjRunStd = re.search(patternRunStd, str_val)
            if searchObjRunPar:
                matchingGroupRunPar = searchObjRunPar.groups()
                data.append(float(matchingGroupRunPar[0]))
                #print('par: ',str_val)
                index=index+1
            elif searchObjSlashDouble:
                matchingGroupSlashDouble = searchObjSlashDouble.groups()
                data.append(float(matchingGroupSlashDouble[0]) + float(matchingGroupSlashDouble[1]))
                #print('slash.double: ',str_val)
                index=index+1
            elif matchObjMara:
                data.append(26)
                index=index+1
                #print('mara: ',str_val)
            elif matchObjHalf:
                data.append(13)
                index=index+1
                #print('half: ',str_val)
            else:#if not in common cases then need to refine search to capture miles
                if searchObjRunStd:
                    matchingGroupRunStd = searchObjRunStd.groups()
                    data.append(float(matchingGroupRunStd[0]))
                    #print('std0: ',str_val)
                    index=index+1
                else: #if not standard further refine search for other cases:
                    #regex in control structure for run \d+mi\/ and run \d+mi\w+
                    #format: run 5mi/<additional characters>
                    patternRunSlash = re.compile(r'[Rr][au]n\s+(\d*\.?\d+)mi\/', re.IGNORECASE)
                    searchObjRunSlash = re.search(patternRunSlash, str_val)
                    if searchObjRunSlash:
                        matchingGroupSlash = searchObjRunSlash.groups()
                        data.append(float(matchingGroupSlash[0]))
                        #print('slash.plain: ',str_val)
                        index=index+1
                    else:
                        #format: run 5mi <additional characters>
                        patternRunWhiteSpace = re.compile(r'[Rr][au]n\s+(\d*\.?\d+)mi\s+', re.IGNORECASE)
                        searchObjRunWhiteSpace = re.search(patternRunWhiteSpace, str_val)
                        if searchObjRunWhiteSpace:
                            matchingGroupWhiteSpace = searchObjRunWhiteSpace.groups()
                            data.append(float(matchingGroupWhiteSpace[0]))
                            #print('whitespace: ',str_val)
                            index=index+1
                        #assume no run on this date and put zero mileage
                        else:
                            data.append(0)
                            #print(str_val)

    #put date values in array
    else:
        for x in range(0,20):
            str_val = row[x]
            dates.append(str_val)

 #write to test.txt file and print out statistic values to console                 
idx = 0            
with open('test.txt', 'w') as outfile:
    for date in dates:
        outfile.write("%d\n" % data[idx])
        outfile.write("%s\n" % detail[idx])
        outfile.write("%s\n" % date)
        idx+=1
print('Maximum is:', max(data),' ','Minimum is: ', min(data)) 
print('\ntotal length of data: ',len(data),' total number of runs parsed: ',index)
print('\nfirst row of mileage: ')
print(data[0:20])
 
  

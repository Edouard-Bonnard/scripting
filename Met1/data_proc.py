#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# written in Python 3.8.2


#####

# ANSYS Scripting Assignment
# Please place this script in the root folder (ex: Met1)
# This script get each file named 'DAQ- Crosshead; … - (Timed).txt' 
# in sub directories, then it creates .xls files with the formated datas.
# Please note that xlwt library is required to execute the script

#####

import os
from os import listdir
from os.path import isfile, join, isdir, basename
from datetime import datetime
import xlwt

print('#####')
print('DATA PROCESSING')

#Function to convert string data with comma to float 
def str_to_float(var): 
    var = var.replace(',','.')
    var = float(var)
    return var

#Function to adapt date format
def date_format(var):
    var = datetime.strptime(var, '%d/%m/%Y %H:%M:%S')
    var = var.strftime('%Y-%m-%d %H:%M:%S')
    return var

dir = os.getcwd() #get current directory
folder_name = basename(dir) 

#Harcoded informations    
Material = 'Cast iron'
Grade = 'FC200'
Supplier = folder_name #should come from folder name
Test = 'Tensile'
file_name = 'DAQ- Crosshead; … - (Timed).txt' #name of the files to get

#Get the sub directories list
sub_directories = [d for d in listdir(dir) if isdir(join(dir, d))]

cpt_files = 0 #number of files converted

#Sub directories loop
for a in range(len(sub_directories)):

    #file opening and loading
    file_path = sub_directories[a] + '/' + file_name #path creation
    
    if os.path.exists(file_path): #check for existing file

        cpt_files +=1

        file = open(file_path, "r", encoding = 'utf8') #file loading
        lines = file.readlines() #whole file is loaded in lines

        #Properties to fill in Info sheet
        SAP_ID = sub_directories[a]
        Test_Machine_ID = lines[1][10:-1] #removal of \n
        Date = lines[2][6:-1] #removal of \n
        Date = date_format(Date)
        Original_file = os.path.abspath(file_path) #absolute path
        Test_ID = SAP_ID + '-' + Test_Machine_ID #concatenation of labels

        ##Writing in Info sheet of Excel File
        #Excel file creation
        xls_file = xlwt.Workbook()
        info = xls_file.add_sheet("Info")
        data = xls_file.add_sheet("Data")

        #Style of xls headers
        header_style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold = True
        header_style.font = font

        #Filling of values in Info sheet
        info.write(0,0,'Property',style=header_style)
        info.write(0,1,'Value',style=header_style)
        info.write(1,0,'Material')
        info.write(1,1,Material)
        info.write(2,0,'Grade')
        info.write(2,1,Grade)
        info.write(3,0,'Supplier')
        info.write(3,1,Supplier)
        info.write(4,0,'SAP ID')
        info.write(4,1,SAP_ID)
        info.write(5,0,'Test')
        info.write(5,1,Test)
        info.write(6,0,'Test Machine ID')
        info.write(6,1,Test_Machine_ID)
        info.write(7,0,'Date')
        info.write(7,1,Date)
        info.write(8,0,'Original File')
        info.write(8,1,Original_file)
        info.write(9,0,'Test ID')
        info.write(9,1,Test_ID)

        ##Writing values in Data sheet
        #Filling of header values in Data sheet
        data.write(0,0,'Crosshead (mm)',style=header_style)
        data.write(0,1,'Load (kN)',style=header_style)
        data.write(0,2,'Time (s)',style=header_style)
        data.write(0,3,'Video Time (s)',style=header_style)
        data.write(0,4,'Axial Strain (mm/mm)',style=header_style)
        data.write(0,5,'Transverse (mm/mm)',style=header_style)

        #Processing and writing of datas
        xls_line = 1 #line initialization in Excel file 

        for i in range(7, len(lines)):
            line = lines[i].split() #space separation of datas
            xls_row = 0 #row initialization in Excel file
            for str_data in line:
                float_data = str_to_float(str_data) #processing
                data.write(xls_line, xls_row, float_data) #writing
                xls_row+=1
            xls_line+=1

        #Saving file
        xls_file_name = Test_Machine_ID + '.xls' #name of the file
        xls_file.save(xls_file_name) #save

        file.close()
    

print(cpt_files, ' .txt file(s) processed to .xls')
print('#####')
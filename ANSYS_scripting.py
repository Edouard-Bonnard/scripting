# -*- Python 3.8.2 -*-
# coding UTF-8

#####

# ANSYS Scripting Assignment
# Script to put in root directory (ex: /Met1)
# This script get each file named 'DAQ- Crosshead; … - (Timed).txt' 
# in sub directories, then it creates .xls files with the formated datas

#####

import os
from os import listdir
from os.path import isfile, join, isdir
import xlwt

#function to convert string data with comma to float 
def str_to_float(var): 
    var = var.replace(',','.')
    var = float(var)
    return var

#harcoded informations    
Material = 'Cast Iron'
Grade = 'FC200'
Supplier = 'Met1' #should come from folder name
Test = 'Tensile'
folder_name = 'Met1' #name of the root directory
file_name = 'DAQ- Crosshead; … - (Timed).txt' #name of the files to get

dir = os.getcwd() #get current directory
dir = dir+'/'+folder_name #create new dir name
os.chdir(dir) #change directory to /Met1

#get the sub directories list
sub_directories = [d for d in listdir(dir) if isdir(join(dir, d))] 

#sub directories loops
for a in range(len(sub_directories)):

    #file opening and loading
    file_path = sub_directories[a] + '/' + file_name #path creation
    file = open(file_path, "r", encoding = 'utf8') #file loading
    lines = file.readlines() #whole file is loaded in lines

    #Properties to fill in Info sheet
    SAP_ID = sub_directories[a]
    Test_Machine_ID = lines[1][10:-1] #removal of \n
    Date = lines[2][6:-1] #removal of \n
    Original_file = os.path.abspath(file_path) #absolute path
    Test_ID = SAP_ID + '-' + Test_Machine_ID #concatenation of labels

    ##Writing in Info sheet of Excel File
    #Excel file creation
    xls_file = xlwt.Workbook()
    info = xls_file.add_sheet("Info")
    data = xls_file.add_sheet("Data")

    #Filling of values in Info sheet
    info.write(0,0,'Property')
    info.write(0,1,'Value')
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
    data.write(0,0,'Crosshead (mm)')
    data.write(0,1,'Load (kN)')
    data.write(0,2,'Time (s)')
    data.write(0,3,'Video Time (s)')
    data.write(0,4,'Axial Strain (mm/mm)')
    data.write(0,5,'Transverse (mm/mm)')

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

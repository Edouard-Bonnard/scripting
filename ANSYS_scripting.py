# -*- Python 3.8.2 -*-
# coding UTF-8

# ANSYS Scripting Assignment
# script to execute in Met1 directory

import os
from os import listdir
from os.path import isfile, join, isdir
import xlwt

Material = 'Cast Iron'
Grade = 'FC200'
Supplier = 'Met1' #should come from folder name
Test = 'Tensile'

folder_name = 'Met1' #name of the root directory
file_name = 'DAQ- Crosshead; … - (Timed).txt' #name of the files to get

dir = os.getcwd() #get current directory

dir = dir+'/'+folder_name #create new dir name
os.chdir(dir) #change directory to /Met1

#files= [f for f in listdir(dir) if isfile(join(dir, f))]

#get the sub directories list
sub_directories = [d for d in listdir(dir) if isdir(join(dir, d))] 

print(sub_directories)

#sub directories loops
for a in range(len(sub_directories)):

    SAP_ID = sub_directories[a]

    print('SAP ID : '+ SAP_ID) #to delete

    file_path = sub_directories[a] + '/' + file_name #path creation
    file = open(file_path, "r", encoding = 'utf8') #file loading

    lines = file.readlines() #whole file is loaded in lines

    Test_Machine_ID = lines[1][10:-1] #removal of \n
    Date = lines[2][6:-1] #removal of \n
    Original_file = os.path.abspath(file_path) #absolute path
    Test_ID = SAP_ID + '-' + Test_Machine_ID #concatenation of labels

    # Writing in Info sheet of Excel File

    #Excel file creation
    xls_file = xlwt.Workbook()
    info = xls_file.add_sheet("Info")
    data = xls_file.add_sheet("Data")

    #Filling of info values
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

    #Saving file
    xls_file_name = Test_Machine_ID + '.xls' #name of the file
    xls_file.save(xls_file_name) #save
    







    file.close()
    A = 1




    







a = 1


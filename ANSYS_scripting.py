# -*- Python 3.8.2 -*-
# coding UTF-8

# ANSYS Scripting Assignment
# script to execute in Met1 directory

import os
from os import listdir
from os.path import isfile, join, isdir
import xlrd

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

for a in range(len(sub_directories)):

    SAP_ID = sub_directories[a]

    print('SAP ID : '+ SAP_ID) #to delete

    file_path = sub_directories[a] + '/' + file_name #path creation
    file = open(file_path, "r", encoding = 'utf8') #file loading

    #newfile = []

    lines = file.readlines()

    Test_Machine_ID = lines[1][10:-1] #removal of \n
    Date = lines[2][:-1] #removal of \n
    Original_file = ' '
    Test_ID = SAP_ID + '-' + Test_Machine_ID

    #for ligne in file:
    #line = file.readline()






    file.close()
    A = 1




    







a = 1


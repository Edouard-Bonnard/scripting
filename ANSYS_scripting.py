# -*- Python 3.8.2 -*-
# coding UTF-8

# ANSYS Scripting Assignment
# script to execute in Met1 directory

import os
from os import listdir
from os.path import isfile, join, isdir
import xlrd

folder_name = 'Met1' #name of the root directory
file_name = 'DAQ- Crosshead; … - (Timed).txt' #name of the files to get

dir = os.getcwd() #get current directory

dir = dir+'/'+folder_name #create new dir name
os.chdir(dir) #change directory to /Met1

#files= [f for f in listdir(dir) if isfile(join(dir, f))]

#get the sub directories list
sub_directories = [d for d in listdir(dir) if isdir(join(dir, d))] 

for a in sub_directories:

    file_path = sub_directories[a] + '/' + file_name

    file = open("")

    







a = 1
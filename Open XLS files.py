# -*- coding: utf-8 -*-
"""
Created on Fri Jul  1 18:53:00 2022

@author: mails
"""

import os
 
# Get the list of all files and directories
Dir = "D:/2022-2023/"
path =[]

for x in os.listdir(Dir):
    if x.endswith(".xls"):
        full_path=Dir+x
        path.append(full_path)
        #print(x)

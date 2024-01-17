#!/usr/bin/python3
##
## Extraxt zipped files
## http://github.com/Thaer-Maddah
##
## Copyright (C) 2023 Thaer Maddah. All rights reserved.

import os
import zipfile
import patoolib # for extraxt all type of comressed files

# print('current dir with getcwd:', os.getcwd())

# print('list files and dirs with listdir:\n', os.listdir(os.getcwd()))

# list sub directories
counter = 0
path = 'test_assign/F22'
for folder_path, folders, files in os.walk(path):
    print(f"Folder path: {folder_path}")
    print('---------------------------------')
    for file in files:
        cur_dir = folder_path # current directory
        print(f"Current dir: {cur_dir}")
        fname = os.path.splitext(file) # get file extension
        # print(fname[0])
        # print(fname[1])

        #check file extension
        if fname[1] == '.rar' or  fname[1] == '.zip' or  fname[1] == '.7z':
            filepath = folder_path + '/' + str(file)
            
            #print(filepath)
            # with interactive = False we can getrid confirmation messages
            patoolib.extract_archive(filepath, outdir=cur_dir, verbosity=-1, interactive=False)
            counter += 1
            #with zipfile.ZipFile(folder_path + '/' + file, 'r') as zip:
            #    zip.extractall('./c25')
    print('---------------------------------')
print(counter, 'Files extracted')


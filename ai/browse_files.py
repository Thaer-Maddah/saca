import os
import sys

def browse(file_extention, directory = ''):
    file_name = []
    folder_name = []
    path = os.getcwd()
    if directory != '':
        path = path + '/' + directory
    else: 
        pass

    #print(path)
    counter = 0
    for folder_path, folders, files in os.walk(path):
        for file in files:
            ext = os.path.splitext(file)[1] # get file extension
            #check file extension
            if ext == file_extention:
                folder_name.append(folder_path)
                file_name.append(file)
                counter += 1
            else:
                pass

    print(counter, 'Files browsed in folder', path)
    return file_name, folder_name


def getFile(filename = '', folder =''):
    if filename == '':
        dir = ""
        file = ""
        if len(sys.argv) > 2 :
            dir = sys.argv[1] + "/"
            file = sys.argv[2]
            path = dir + file
        else:
            print ("You must give a directory and file name!.")
            sys.exit(2)
    else:
        path = folder + '/' +  filename
    return path



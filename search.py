############
# This script recursively searches through a directory of over 2,000 files and sub folders 
# in order to isolate relevant files for version tracking purposes:
# data collected about files in directory is stored in a dictionary containing list and then written to a csv file
# for further analysis

# This program prevented our DOT contract team from searching for 300 files through a directory of 
# 2000 plus files and then manually writing down data associated with those files contained in the file name

import glob
import pandas as pd
import csv


# create list of all files in directory
paths = glob.glob('**/*.pdf', recursive=True)

# list of folders to ignore
ignore_folders = ['5000','5020','5030','5031','5035','5099']

# list of values file path must contain
must_have = ['MOD', 'BASE']

# upper case all path names to filter files correctly 
upper_paths = []
for path in paths:
    upper_paths.append(path.upper())

# declare empty filtered file list 
filtered = []

# filter paths down to relevant files
for path in upper_paths: 
    # if path is in folder to ignore list, pass
    if any(folder in path for folder in ignore_folders):
        pass
    # if path string has required vale, add to filtered list
    elif any(string in path for string in must_have):
        filtered.append(path)

# empty dict for storing file data
file_dicitonary = {}

# empty list for storing file dictionary 
mod_list = []

# iterate through files 
for file_path in filtered:
    # set file path as dicitonary key
    file_dicitonary[file_path] = {}
    # split file path by slashes to isolate the file name
    split_path = file_path.split('\\')
    # assign file name to variable
    file_name = split_path[-1]
    # split file name by spaces to isolate data within filename
    split_name = file_name.split()

    # set file path as key 
    file_dicitonary[file_path]['file path'] = file_path
    file_dicitonary[file_path]['pmoc'] = split_name[0]

    # if files are denoted as task order or contract store values as follows
    standard_file_types = ['TO','CO']
    if any(val in split_name[1] for val in standard_file_types):
        file_dicitonary[file_path]['file type'] = split_name[1][:2] 
        file_dicitonary[file_path]['number'] =split_name[1][2:]
    # if file_type not denoted in file name, store differently 
    else: 
        file_dicitonary[file_path]['file type'] = 'Other'
        file_dicitonary[file_path]['number'] = split_name[1]

    file_dicitonary[file_path]['mod #'] = split_name[2].split('.')[0]
    
    # use try to isolate last value in list and avoid index error 
    # possibly not necesary when referencing postition [-1] instaid of position [3]....
    try:
        file_dicitonary[file_path]['pmoc type'] = split_name[-1].split('.')[0]
    except IndexError:
        file_dicitonary[file_path]['pmoc type'] = 'NA'

    # append dicitonary record to list 
    mod_list.append(file_dicitonary[file_path])

# write data stored in dictionary containing list to csv
keys = mod_list[0].keys()
with open('mods.csv', 'w', newline='')  as output_file:
    dict_writer = csv.DictWriter(output_file, keys)
    dict_writer.writeheader()
    dict_writer.writerows(mod_list)
    dict_writer.writerows(base_list)

# for x in mod_list:
#     print(x)

# print(len(mod_list))
# print(len(base_list))
       


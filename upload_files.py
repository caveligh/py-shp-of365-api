from office365_api import SharePoint

import re 
import sys, os
from pathlib import PurePath

# 1 args = Root Directory Path of files to upload_file
ROOT_DIR = sys.argv[1]
# 2 args = Sharepoint folder name. May include subfolders to upload to
SHAREPOINT_FOLDER_NAME = sys.argv[2]
# 3 args = Sharepoint fIle name PATTERN. only upload files with this pattern
FILE_NAME_PATTER = sys.argv[3]

def upload_files(folder, keyword=None):
    file_list = get_list_of_files(folder)
    for file in file_list:
        if keyword is None or keyword == 'None' or re.search(keyword, file[0]):
            file_content = get_file_Content(file[1])
            SharePoint().upload_file(file[0], SHAREPOINT_FOLDER_NAME, file_content)               
    
def get_list_of_files(folder):
    file_list = []
    folder_item_list = os.listdir(folder)
    for item in folder_item_list:
        item_full_path = PurePath(folder, item)
        if os.path.isfile(item_full_path):
            file_list.append([item, item_full_path])
    return file_list    

# read files and return the content files
def get_file_Content(file_path):
    with open(file_path, 'rb') as file:
        return file.read()
    
if __name__ == '__main__':
    upload_files(ROOT_DIR, FILE_NAME_PATTER)

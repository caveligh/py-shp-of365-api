from pathlib import PurePath

import re 
import sys, os
import logging

sys.path.append('../')
from office365_api import SharePoint

# 1 arg = Sharepoint folder name. May include subfolders
FOLDER_NAME = sys.argv[1]
# 2 args = locate or remote folder_dest
FOLDER_DEST = sys.argv[2]
# 3 args = Sharepoint file name. This is used when only one file is being downloaded
# If all files will be downloaded, then set this value as "None"
FILE_NAME = sys.argv[3]
# 4 args = Sharepoint file name pattern
# If no pattern match files are required to be downloaded, then set this value as "None"
FILE_NAME_PATTERN = sys.argv[4]

# Configuración de logs
logging.basicConfig(filename='log.txt', format='%(asctime)s , %(name)s , %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

def save_file(file_n, file_obj):
    file_dir_path = PurePath(FOLDER_DEST, file_n)
    with open(file_dir_path, 'wb') as f:
        f.write(file_obj)

def get_file(file_n, folder):
    file_obj = SharePoint().download_file(file_n, folder)
    save_file(file_n, file_obj)
    # Generamos mensaje de log
    logger.info(f"Descargado archivo {file_n} de la carpeta {folder} y almacenado en {FOLDER_DEST}")

def get_files(folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        get_file(file.name, folder)

def get_files_by_pattern(keyword, folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        if re.search(keyword, file.name):
            get_file(file.name, folder)

if __name__ == '__main__':
    try:
        if FILE_NAME != 'None':
            get_file(FILE_NAME, FOLDER_NAME)
        elif FILE_NAME_PATTERN != 'None':
            get_files_by_pattern(FILE_NAME_PATTERN, FOLDER_NAME)
        else:
            get_files(FOLDER_NAME)
    except Exception as e:
        # Generamos mensaje de log si ocurre un error durante la ejecución del script
        logger.exception("Ha ocurrido un error durante la ejecución del método get_file(s).")
    finally:
        # Cerramos el archivo de logs
        logging.shutdown()

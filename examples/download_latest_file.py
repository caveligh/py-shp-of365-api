from pathlib import PurePath

import sys
import logging

sys.path.append('../')
from office365_api import SharePoint

# 1 args  = SharePoint Folder name
FOLDER_NAME = sys.argv[1]
# 2 args = location or remote folder destintion 
FOLDER_DEST = sys.argv[2]

# Configuración de logs
logging.basicConfig(filename='log.txt', format='%(asctime)s , %(name)s , %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

def save_file(file_name, file_obj, folder_dest):
    file_dir_path = PurePath(folder_dest, file_name)
    with open(file_dir_path, 'wb') as f:
        f.write(file_obj)
        
def get_latest_file(folder, folder_dest):
    file_name, content = SharePoint().download_latest_file(folder)
    save_file(file_name, content, folder_dest)
    # Generamos mensaje de log
    logger.info(f"Descargado archivo {file_name} de la carpeta {folder} y almacenado en {folder_dest}")
    
if __name__ == '__main__':
    try:
        get_latest_file(FOLDER_NAME, FOLDER_DEST)
    except Exception as e:
        # Generamos mensaje de log si ocurre un error durante la ejecución del script
        logger.exception("Ha ocurrido un error durante la ejecución del método get_latest_file.")
    finally:
        # Cerramos el archivo de logs
        logging.shutdown()    
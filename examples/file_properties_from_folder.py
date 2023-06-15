import sys
import logging

sys.path.append('../')
from office365_api import SharePoint

# 1 args = SharePoint Folder name
FOLDER_NAME = sys.argv[1]

# Configuración de logs
logging.basicConfig(filename='log.txt', format='%(asctime)s , %(name)s , %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

def get_properties_by_folder(folder):
    files_properties = SharePoint().get_file_properties_from_folder(folder)
    logger.info(f"File count: {len(files_properties)}.")
    print('File count:', len(files_properties))
    for file in files_properties:
        print(file)
        
if __name__ == '__main__':
    try:
        get_properties_by_folder(FOLDER_NAME)
    except Exception as e:
        # Generamos mensaje de log si ocurre un error durante la ejecución del script
        logger.exception("Ha ocurrido un error durante la ejecución del método get_properties_by_folder.")
    finally:
        # Cerramos el archivo de logs
        logging.shutdown()
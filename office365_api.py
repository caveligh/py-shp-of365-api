from office365.sharepoint.client_context import ClientContext 
from office365.sharepoint.files.file import File
from office365.runtime.auth.client_credential import ClientCredential

import logging
import datetime
import environ

env = environ.Env()
# Busca en .env las variables de entorno
environ.Env.read_env()

# USERNAME = env("sharepoint_email")
# PASSWORD = env("sharepoint_password")
SHAREPOINT_SITE = env("sharepoint_url_site")
SHAREPOINT_SITE_NAME = env("sharepoint_site_name")
SHAREPOINT_DOC_LIBRARY = env("sharepoint_doc_library")
SHAREPOINT_CLIENT_ID = env("sharepoint_client_id")
SHAREPOINT_CLIENT_SECRET = env("sharepoint_client_secret")

# Configuración de logs
logging.basicConfig(filename='log.txt', format='%(asctime)s , %(name)s , %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

class SharePoint:
    def _auth(self):
        """Authenticate the user in Sharepoint using MFA and return the connection.

        :type self: T
        """
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(
            ClientCredential(
                SHAREPOINT_CLIENT_ID,
                SHAREPOINT_CLIENT_SECRET
            )
        )
        return conn
    
    def _get_files_list(self, folder_name):
        """Get files list from an existing Sharepoint document library.

        :param str folder_name: Folder name
        """
        conn = self._auth()
        # Gets the list of files in the Sharepoint library
        target_folder_url = f'{SHAREPOINT_DOC_LIBRARY}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files
    
    def get_folder_list(self, folder_name):
        """Get folder list from an existing Sharepoint document library.

        :param str folder_name: Folder name
        """
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC_LIBRARY}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Folders"]).get().execute_query()
        return root_folder.folders
    
    def download_file(self, file_name, folder_name):
        """Get a file from an existing Sharepoint document library and download it.

        :param str file_name: File name
        :param str folder_name: Folder name
        """ 
        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC_LIBRARY}/{folder_name}/{file_name}'
        # Generamos mensaje de log
        logger.info(f"Descargando archivo {file_name} de la carpeta {folder_name}.")
        file = File.open_binary(conn, file_url)
        return file.content
    
    def download_latest_file(self, folder_name):
        """Get a file from an existing document library and download it.

        :param str folder_name: Folder name
        """         
        date_format = "%Y-%m-%dT%H:%M:%SZ"
        files_list = self._get_files_list(folder_name)
        file_dict = {}
        for file in files_list:
            dt_obj = datetime.datetime.strptime(file.time_last_modified, date_format)
            file_dict[file.name] = dt_obj
        # sort dict object to get the latest file
        file_dict_sorted = {key:value for key, value in sorted(file_dict.items(), key=lambda item:item[1], reverse=True)}    
        latest_file_name = next(iter(file_dict_sorted))
        content = self.download_file(latest_file_name, folder_name)
        # Generamos mensaje de log
        logger.info(f"Devolviendo el último archivo {latest_file_name} de la carpeta {folder_name}.")
        return latest_file_name, content
    
    def upload_file(self, file_name, folder_name, content):
        """Get a file from an existing Sharepoint document library and upload it.

        :param str file_name: File name
        :param str folder_name: Folder name
        :param content: bytes or str
        """         
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC_LIBRARY}/{folder_name}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.upload_file(file_name, content).execute_query()
        # Generamos mensaje de log
        logger.info(f"Archivo {file_name} subido a la carpeta {folder_name}. Respuesta {response}.")
        return response

    def upload_file_in_chunks(self, file_path, folder_name, chunk_size, chunk_uploaded=None, **kwargs):
        """Upload a file as multiple chunks to an existing Sharepoint document library

        :param str or typing.IO file_path: path where file to upload resides or file handle
        :param int chunk_size: upload chunk size (in bytes)
        :param (long)->None or None chunk_uploaded: uploaded event
        :param kwargs: arguments to pass to chunk_uploaded function
        """  
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC_LIBRARY}/{folder_name}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.files.create_upload_session(
            source_path=file_path,
            chunk_size=chunk_size,
            chunk_uploaded=chunk_uploaded,
            **kwargs
        ).execute_query()
        # Generamos mensaje de log
        logger.info(f"Archivo {file_path} subido en trozos a la carpeta {folder_name}. Respuesta {response}.")
        return response
 
    def get_list(self, list_name):
        """Retrieve a list items from an existing Sharepoint List.

        :param str list_name: List name
        """         
        conn = self._auth()
        target_list = conn.web.lists.get_by_title(list_name)
        items = target_list.items.get().execute_query()
        return items
    
            
    def get_file_properties_from_folder(self, folder_name):
        """Retrieve a list properties from an existing folder name.

        :param str folder_name: Folder name
        """  
        files_list = self._get_files_list(folder_name)
        properties_list = []
        for file in files_list:
            file_dict = {
                'file_id': file.unique_id,
                'file_name': file.name,
                'major_version': file.major_version,
                'minor_version': file.minor_version,
                'file_size': file.length,
                'time_created': file.time_created,
                'time_last_modified': file.time_last_modified
            }
            properties_list.append(file_dict)
            file_dict = {}
        return properties_list        

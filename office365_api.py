import environ
from office365.sharepoint.client_context import ClientContext 
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File

env = environ.Env()
# Busca en .env las variables de entorno
environ.Env.read_env()

USERNAME = env("sharepoint_email")
PASSWORD = env("sharepoint_password")
SHAREPOINT_SITE = env("sharepoint_url_site")
SHAREPOINT_SITE_NAME = env("sharepoint_site_name")
SHAREPOINT_DOC_LIBRARY = env("sharepoint_doc_library")

class SharePoint:
    def _auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(
                USERNAME,
                PASSWORD
            )
        )
        return conn
    
    def _get_files_list(self, folder_name):
        conn = self._auth()
        # Obtiene la lista de archivos de la biblioteca
        target_folder_url = f'{SHAREPOINT_DOC_LIBRARY}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files
    
    def download_file(self, file_name, folder_name):
        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC_LIBRARY}/{folder_name}/{file_name}'
        file = File.open_binary(conn, file_url)
        return file.content
    
    def download_latest_file(self, folder_name):
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
        return latest_file_name, content
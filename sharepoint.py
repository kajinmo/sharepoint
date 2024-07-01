import datetime
import os
import re
from io import BytesIO

import pandas as pd
from dotenv import load_dotenv
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

load_dotenv()

USERNAME = os.getenv("SHAREPOINT_EMAIL")
PASSWORD = os.getenv("SHAREPOINT_PASSWORD")
SHAREPOINT_SITE = os.getenv("SHAREPOINT_URL_SITE")
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
SHAREPOINT_DOC = os.getenv("SHAREPOINT_DOC_LIBRARY")


class SharePoint:
    def _auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(UserCredential(USERNAME, PASSWORD))
        return conn

    def get_files_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f"{SHAREPOINT_DOC}/{folder_name}"
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files

    def download_file(self, file_name, folder_name):
        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}/{file_name}'
        file = File.open_binary(conn, file_url)
        return file.content

    # def download_files(self, folder_name):
    # 	files_list = self.get_files_list(folder_name)
    # 	files_dict = {}
    # 	for file in files_list:
    # 		content = self.download_file(file.name, folder_name)
    # 		files_dict[file.name] = content
    # 	return files_dict

    def download_files(self, folder_name):
        files_list = self.get_files_list(folder_name)
        files_list_with_content = []

        for file in files_list:
            content = self.download_file(file.name, folder_name)
            files_list_with_content.append((file.name, content))

        return files_list_with_content

    def download_latest_file(self, folder_name):
        date_format = "%Y-%m-%dT%H:%M:%SZ"
        files_list = self.get_files_list(folder_name)
        file_dict = {}
        for file in files_list:
            dt_obj = datetime.datetime.strptime(file.time_last_modified, date_format)
            file_dict[file.name] = dt_obj
        # sort dict object to get the latest file
        file_dict_sorted = {key: value for key, value in
                            sorted(file_dict.items(), key=lambda item: item[1], reverse=True)}
        latest_file_name = next(iter(file_dict_sorted))
        content = self.download_file(latest_file_name, folder_name)
        return latest_file_name, content

    def upload_file(self, file_name, folder_name, content):
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.upload_file(file_name, content).execute_query()
        return response

    def upload_file_in_chunks(self, file_path, folder_name, chunk_size, chunk_uploaded=None, **kwargs):
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.files.create_upload_session(
            source_path=file_path,
            chunk_size=chunk_size,
            chunk_uploaded=chunk_uploaded,
            **kwargs
        ).execute_query()
        return response

    def get_list(self, list_name):
        conn = self._auth()
        target_list = conn.web.lists.get_by_title(list_name)
        items = target_list.items.get().execute_query()
        return items

    def get_file_properties_from_folder(self, folder_name):
        files_list = self.get_files_list(folder_name)
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

    def update_fund_files(self, folder_name='Gestão/Dashboard'):
        """
        Update the FUND_FILES dictionary with the newest XML file names from the SharePoint list.
        """
        pattern = re.compile(r'^FD\d+_(\d+)_(\d+)_(\w+)_(FIM|FIA).xml$')

        files = self.get_files_list(folder_name)
        xml_files = [file.name for file in files]

        new_fund_files = {}
        for xml_file in xml_files:
            match = pattern.match(xml_file)
            if match:
                fund_key = match.group(3).upper()
                file_date = int(match.group(1))

                # Check if fund_key already exists and if the current file is newer than the previous one.
                if fund_key not in new_fund_files or file_date > int(pattern.match(new_fund_files[fund_key]).group(1)):
                    new_fund_files[fund_key] = xml_file

        print(f"New fund files: {new_fund_files}")
        return new_fund_files

    def download_and_read_excel(self, file_name, folder_name):
        file_content = self.download_file(file_name, folder_name)
        df = pd.read_excel(BytesIO(file_content))
        return df
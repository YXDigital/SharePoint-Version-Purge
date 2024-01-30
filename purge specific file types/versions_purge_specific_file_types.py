"""
THIS SCRIPT WILL ONLY BE EXECUTED BY THE IT DEPARTMENT.
WITH THE SCRIPT WE CAN SCAN A PROJECT YEAR AND PURGE RFEM FILES TO THE LAST 2 VERSIONS
"""
from office365.sharepoint.client_context import ClientContext
from pathlib import Path
from functools import partial

client_id = "Microsoft Azure App registration Client ID"
tenant_id = "Microsoft azure Tenant ID"
site_url = "SharePoint hostname"
root_folder = "Sharepoint root documents folder"

ctx = ClientContext(site_url).with_device_flow(tenant_id, client_id)
web = ctx.web.get().execute_query()


def clean_versions(file_to_clean, max_versions):
    """
    this function loops to the versions of the file exept for the max_versions and deletes them
    """
    print(f'This file has {len(file_to_clean.versions)} versions')
    versions_to_delete = file_to_clean.versions[:-max_versions]
    for version in versions_to_delete:  # type: FileVersion
        print(f'Removing version number: {version.version_label}')
        file_to_clean.versions.delete_by_label(version.version_label).execute_query()


def get_file_versions(file_url):
    """
    THis functions requests the file versions object
    :return FileVersionObject
    """
    file_versions = ctx.web.get_file_by_server_relative_path(file_url).expand(["Versions"]).get().execute_query()
    return file_versions


def get_files_in_folder(folder):
    folder_path = f'{root_folder}/{folder}'
    root_site_folder = ctx.web.get_folder_by_server_relative_path(folder_path).get().execute_query()
    project_files = root_site_folder.get_files(recursive=True).get().execute_query()

    return project_files


def clean_file_versions(folder_to_scan, max_versions, extension_lists):
    for folder in folder_to_scan:
        projects_folder = get_files_in_folder(folder)  # type: File
        for file in projects_folder:
            if Path(file.serverRelativeUrl).suffix in extension_lists:
                print(f'Name: {file.name}, Versions: {len(get_file_versions(file.serverRelativeUrl).versions)} ')
                if len(get_file_versions(file.serverRelativeUrl).versions) > max_versions:
                    clean_versions(get_file_versions(file.serverRelativeUrl), max_versions)


if __name__ == "__main__":
    file_type_list = [".docx", ".xlsx", ".pptx"]

    # INPUT FOR THE PROJECT YEAR THAT NEEDS TO BE SCANNED OR PROJECT FOLDERS
    folders_to_scan = ["IBA_HD_log/BM001/dat_continuous/24/01/15"]
    clean_file_versions([folder for folder in folders_to_scan], 1, file_type_list)

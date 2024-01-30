"""
THIS SCRIPT WILL ONLY BE EXECUTED BY THE IT DEPARTMENT.
WITH THE SCRIPT WE CAN SCAN A PROJECT YEAR AND PURGE RFEM FILES TO THE LAST 2 VERSIONS
"""
from office365.sharepoint.client_context import ClientContext

client_id = "Microsoft Azure App registration Client ID"
tenant_id = "Microsoft azure Tenant ID"
site_url = "SharePoint hostname"
root_folder = "Sharepoint root documents folder"

ctx = ClientContext(site_url).with_device_flow(tenant_id, client_id)
web = ctx.web.get().execute_query()


def get_file_versions(file_url):
    """
    THis functions requests the file versions object
    :return FileVersionObject
    """
    file_versions = ctx.web.get_file_by_server_relative_path(file_url).expand(["Versions"]).get().execute_query()
    return file_versions


def clean_single_file_version(file_to_clean, max_versions):
    generated_file_to_clean = f'{root_folder}/{file_to_clean}'
    if len(get_file_versions(generated_file_to_clean).versions) > max_versions:
        print(f'This file has {len(get_file_versions(generated_file_to_clean).versions)} versions')
        versions_to_delete = get_file_versions(generated_file_to_clean).versions[:-max_versions]
        print(versions_to_delete)
        for version in versions_to_delete:  # type: FileVersion
            print(version.version_label)
            get_file_versions(generated_file_to_clean).versions.delete_by_label(version.version_label).execute_query()


if __name__ == "__main__":
    # INPUT FOR ONE SINGLE FILE THAT NEEDS TO BE CLEANED
    single_file_path = ''
    clean_single_file_version(single_file_path, 8)

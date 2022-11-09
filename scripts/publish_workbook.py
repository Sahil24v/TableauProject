"""
Neccessory Module imports
"""
import os
import json
import argparse
import xml.etree.ElementTree as ET
import logging
import tableauserverclient as TSC
import requests
xmlns = {'t': 'http://tableau.com/api'}


class ApiCallError(Exception):
    """
    Class Description
    """
    pass


def _encode_for_display(text):
    return text.encode('ascii', errors="backslashreplace").decode('utf-8')


def _check_status(server_response, success_code):
    if server_response.status_code != success_code:
        parsed_response = ET.fromstring(server_response.text)
        error_element = parsed_response.find('t:error', namespaces=xmlns)
        summary_element = parsed_response.find(
            './/t:summary', namespaces=xmlns)
        detail_element = parsed_response.find('.//t:detail', namespaces=xmlns)
        code = error_element.get(
            'code', 'unknown') if error_element is not None else 'unknown code'
        summary = summary_element.text if summary_element is not None else 'unknown summary'
        detail = detail_element.text if detail_element is not None else 'unknown detail'
        error_message = f'{code}: {summary} - {detail}'
        raise ApiCallError(error_message)
    return


def sign_in(data):
    """
    Funcrion Description
    """
    tableau_auth = TSC.TableauAuth(
        args.username, args.password, None if data['is_site_default'] else data['site_name'])
    server = TSC.Server(data['server_url'], use_server_version=True)
    server.auth.sign_in(tableau_auth)
    server_response = vars(server)
    auth_token = server_response.get('_auth_token')
    version = server_response.get('version')
    return server, auth_token, version


def get_project(server, data):
    """
    Funcrion Description
    """
    all_projects, pagination_item = server.projects.get()
    project = next(
        (project for project in all_projects if project.name == data['project_path']), None)
    if project.id is not None:
        return project.id
    else:
        raise LookupError(
            f"The project for {data['file_path']} workbook could not be found.")


def publish_workbook(server, data):
    """
    Funcrion Description
    """
    project_id = get_project(server, data)
    wb_path = os.path.dirname(os.path.realpath(__file__)).rsplit(
        '/', 1)[0] + "/workbooks/" + data['file_path']
    new_workbook = TSC.WorkbookItem(
        name=data['name'], project_id=project_id, show_tabs=data['show_tabs'])
    new_workbook = server.workbooks.publish(
        new_workbook, wb_path, 'Overwrite', hidden_views=data['hidden_views'])

    print(
        f"\nSuccessfully published {data['file_path']} Workbook in {data['project_path']} project in {data['site_name']} site.")

    # Update Workbook and set tags
    if len(data['tags']) > 0:
        new_workbook.tags = set(data['tags'])
        new_workbook = server.workbooks.update(
            new_workbook)
        print("\nUpdate Workbook Successfully and set Tags.")


def get_workbook_id(server, data):
    """
    Function Description
    """
    all_workbooks_items, pagination_item = server.workbooks.get()
    workbook_id_list = [
        workbook.id for workbook in all_workbooks_items if workbook.name == data['name']]
    return workbook_id_list


def get_group_id(server, permission_group_name):
    """
    Function Description
    """
    all_groups, pagination_item = server.groups.get()
    group_id_list = [
        group.id for group in all_groups if group.name == permission_group_name]
    return group_id_list


def get_user_id(server, permission_user_name):
    """
    Funcrion Description
    """
    all_users, pagination_item = server.users.get()
    user_id_list = [
        user.id for user in all_users if user.name == permission_user_name]
    return user_id_list


def query_permission(data, wb_id, permission_user_or_group_id, version, auth_token, is_group):
    """
    Funcrion Description
    """
    url = f"https://tableau.devinvh.com/api/{version}/sites/{data['site_id']}/workbooks/{wb_id}/permissions"

    server_response = requests.get(
        url, headers={'x-tableau-auth': auth_token}, timeout=5000)
    _check_status(server_response, 200)
    server_response = _encode_for_display(server_response.text)
    parsed_response = ET.fromstring(server_response)
    capabilities = parsed_response.findall(
        './/t:granteeCapabilities', namespaces=xmlns)

    if is_group:
        for capability in capabilities:
            group = capability.find('.//t:group', namespaces=xmlns)
            if group is not None and group.get('id') == permission_user_or_group_id:
                return capability.findall('.//t:capability', namespaces=xmlns)
    else:
        for capability in capabilities:
            user = capability.find('.//t:user', namespaces=xmlns)
            if user is not None and user.get('id') == permission_user_or_group_id:
                return capability.findall('.//t:capability', namespaces=xmlns)


def add_permission(data, wb_id, permission_user_or_group_id, version, auth_token, permission_name, permission_mode, is_group):
    """
    Funcrion Description
    """
    url = f"https://tableau.devinvh.com/api/{version}/sites/{data['site_id']}/workbooks/{wb_id}/permissions"

    xml_request = ET.Element('tsRequest')
    permissions_element = ET.SubElement(xml_request, 'permissions')
    ET.SubElement(permissions_element, 'workbook', id=wb_id)
    grantee_element = ET.SubElement(permissions_element, 'granteeCapabilities')

    if is_group:
        ET.SubElement(grantee_element, 'group', id=permission_user_or_group_id)
    else:
        ET.SubElement(grantee_element, 'user', id=permission_user_or_group_id)

    capabilities_element = ET.SubElement(grantee_element, 'capabilities')
    ET.SubElement(capabilities_element, 'capability',
                  name=permission_name, mode=permission_mode)
    xml_request = ET.tostring(xml_request)

    server_request = requests.put(
        url, data=xml_request,  headers={'x-tableau-auth': auth_token}, timeout=5000)
    _check_status(server_request, 200)


def delete_permission(data, auth_token, wb_id, permission_user_or_group_id, permission_name, existing_mode, version, is_group):
    """
    Funcrion Description
    """
    if is_group:
        url = f"https://tableau.devinvh.com/api/{version}/sites/{data['site_id']}/workbooks/{wb_id}/permissions/groups/{permission_user_or_group_id}/{permission_name}/{existing_mode}"
    else:
        url = f"https://tableau.devinvh.com/api/{version}/sites/{data['site_id']}/workbooks/{wb_id}/permissions/users/{permission_user_or_group_id}/{permission_name}/{existing_mode}"

    server_response = requests.delete(
        url, headers={'x-tableau-auth': auth_token},
        timeout=5000)
    _check_status(server_response, 204)


def main(arguments):
    """
    Funcrion Description
    """
    project_data_json = json.loads(arguments.project_data)

    try:
        for data in project_data_json:

            # Step: Sign in to Tableau server.
            server, auth_token, version = sign_in(data)

            if data['project_path'] is None:
                raise LookupError(
                    "The project_path field is Null in JSON Template.")
            else:
                # Step: Form a new workbook item and publish.
                publish_workbook(server, data)

                if data['permissions']:
                    for permission_data in data['permissions']:
                        if permission_data['permission_template']:
                            print("---------------------------------------------------------------------------------------------------------------------------------")

                            is_group = None

                            # Step: Get the Workbook ID from the Workbook Name
                            wb_id = get_workbook_id(server, data)[0]

                            # Step: Get the User or Group ID of permission assigned
                            if permission_data['permission_group_name'] and not permission_data['permission_user_name']:
                                permission_user_or_group_id = get_group_id(
                                    server, permission_data['permission_group_name'])[0]
                                is_group = True
                            elif not permission_data['permission_group_name'] and permission_data['permission_user_name']:
                                permission_user_or_group_id = get_user_id(
                                    server, permission_data['permission_user_name'])[0]
                                is_group = False
                            else:
                                logging.info(
                                    "permission_group_name and permission_user_name are both null, Please provide anyone of that.")

                            # get permissions of specific workbook
                            user_permissions = query_permission(
                                data, wb_id, permission_user_or_group_id, version, auth_token, is_group)

                            for permission_name, permission_mode in permission_data['permission_template'].items():
                                update_permission_flag = True
                                if user_permissions is None:
                                    add_permission(
                                        data, wb_id, permission_user_or_group_id, version, auth_token, permission_name, permission_mode, is_group)
                                    print(
                                        f"\tPermission {permission_name} is set to {permission_mode} Successfully in {wb_id}\n")
                                    update_permission_flag = False
                                else:
                                    for permission in user_permissions:
                                        if permission.get('name') == permission_name:
                                            if permission.get('mode') != permission_mode:
                                                existing_mode = permission.get(
                                                    'mode')
                                                delete_permission(
                                                    data, auth_token, wb_id, permission_user_or_group_id, permission_name, existing_mode, version, is_group)
                                                update_permission_flag = True
                                                print(
                                                    f"\tPermission {permission_name} : {existing_mode} is deleted Successfully in {wb_id}\n")
                                            else:
                                                update_permission_flag = False

                                if update_permission_flag:
                                    add_permission(
                                        data, wb_id, permission_user_or_group_id, version, auth_token, permission_name, permission_mode, is_group)
                                    print(
                                        f"\tPermission {permission_name} is set to {permission_mode} Successfully in {wb_id}\n")
                                else:
                                    print(
                                        f"\tPermission {permission_name} is already set to {permission_mode} on {data['name']}\n")

                        else:
                            logging.info(
                                "Something went wrong, Error occured.\n User Name or List of Permissions in template are null")
                else:
                    logging.info(
                        "Something went wrong, Error occured.\n Permissions in template are null")

            # Step: Sign Out to the Tableau Server
            server.auth.sign_out()

    except Exception as tableu_exception:
        logging.error(
            "Something went wrong, Exception Fired.\n %s", tableu_exception)
        exit(1)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(allow_abbrev=False)
    parser.add_argument('--username', action='store',
                        type=str, required=True)
    parser.add_argument('--password', action='store',
                        type=str, required=True)
    parser.add_argument('--project_data', action='store',
                        type=str, required=True)
    args = parser.parse_args()
    main(args)

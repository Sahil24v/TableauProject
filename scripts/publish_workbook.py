import os
import json
import argparse
import tableauserverclient as TSC
import requests
import xml.etree.ElementTree as ET
from requests.auth import HTTPBasicAuth
from pprint import pprint



xmlns = {'t': 'http://tableau.com/api'}


class ApiCallError(Exception):
    pass


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


def signin(data):
    tableau_auth = TSC.TableauAuth(
        args.username, args.password, None if data['is_site_default'] else data['site_name'])
    server = TSC.Server(data['server_url'], use_server_version=True)
    server.auth.sign_in(tableau_auth)

    server_response = vars(server)
    auth_token = server_response.get('_auth_token')
    version = server_response.get('version')
    user_id = server_response.get('_user_id')

    return server, auth_token, version, user_id


def getProject(server, data):
    all_projects, pagination_item = server.projects.get()
    project = next(
        (project for project in all_projects if project.name == data['project_path']), None)

    if project.id is not None:
        return project.id
    else:
        raise LookupError(
            f"The project for {data['file_path']} workbook could not be found.")


def publishWB(server, data):
    project_id = getProject(server, data)

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
        print(
            f"\nUpdate Workbook Successfully and set Tags.")


def createSchedule(server):
    # Create an interval to run every 2 hours between 2:30AM and 11:00PM
    hourly_interval = TSC.HourlyInterval(start_time=time(2, 30),
                                         end_time=time(23, 0),
                                         interval_value=2)
    # Create schedule item
    hourly_schedule = TSC.ScheduleItem(
        "Hourly-Schedule", 50, TSC.ScheduleItem.Type.Extract, TSC.ScheduleItem.ExecutionOrder.Parallel, hourly_interval)
    # Create schedule
    hourly_schedule = server.schedules.create(hourly_schedule)


def getWBID(server, data):
    all_workbooks_items, pagination_item = server.workbooks.get()
    return [workbook.id for workbook in all_workbooks_items if workbook.name == data['name']]


def _encode_for_display(text):
    return text.encode('ascii', errors="backslashreplace").decode('utf-8')


def query_permission(data, wb_id, user_id, version, auth_token):
    url = f"https://tableau.devinvh.com/api/{version}/sites/{data['site_id']}/workbooks/{wb_id}/permissions"

    server_response = requests.get(url, headers={'x-tableau-auth': auth_token})
    _check_status(server_response, 200)
    server_response = _encode_for_display(server_response.text)

    # Reads and parses the response
    parsed_response = ET.fromstring(server_response)

    # Find all the capabilities for a specific user
    capabilities = parsed_response.findall(
        './/t:granteeCapabilities', namespaces=xmlns)
    for capability in capabilities:
        user = capability.find('.//t:user', namespaces=xmlns)
        if user is not None and user.get('id') == user_id:
            return capability.findall('.//t:capability', namespaces=xmlns)


def add_permission(data, wb_id, user_id, version, auth_token):
    url = f"https://tableau.devinvh.com/api/{version}/sites/{data['site_id']}/workbooks/{wb_id}/permissions"

    xml_request = ET.Element('tsRequest')
    permissions_element = ET.SubElement(xml_request, 'permissions')
    ET.SubElement(permissions_element, 'workbook', id=wb_id)
    grantee_element = ET.SubElement(permissions_element, 'granteeCapabilities')
    ET.SubElement(grantee_element, 'user', id=user_id)
    capabilities_element = ET.SubElement(grantee_element, 'capabilities')
    ET.SubElement(capabilities_element, 'capability',
                  name=data['permission_name'], mode=data['permission_mode'])
    xml_request = ET.tostring(xml_request)

    server_request = requests.put(
        url, data=xml_request,  headers={'x-tableau-auth': auth_token})
    _check_status(server_request, 200)


def delete_permission(data, auth_token, wb_id, user_id, permission_name, existing_mode):
    url = f"https://tableau.devinvh.com/api/{version}/sites/{data['site_id']}/workbooks/{wb_id}/permissions/users/{user_id}/{permission_name}/{existing_mode}"

    print("\tDeleting existing permission")

    server_response = requests.delete(
        url, headers={'x-tableau-auth': auth_token})
    _check_status(server_response, 204)
    return


def main(args):
    project_data_json = json.loads(args.project_data)
    try:
        for data in project_data_json:
            # Step: Sign in to Tableau server.
            server, auth_token, version, user_id = signin(data)

            if data['project_path'] is None:
                raise LookupError(
                    f"The project project_path field is Null in JSON Template.", data['file_path'])
            else:
                # Step: Form a new workbook item and publish.
                # publishWB(server, data)

                # Step: Get the Workbook ID from the Workbook Name
                wb_id = getWBID(server, data)

                # get permissions of specific workbook
                user_permissions = query_permission(
                    data, wb_id[0], user_id, version, auth_token)

                print("type of user_permissions ::", type(user_permissions))
                for i in user_permissions:
                    print(i)

                # Step: Update Project permissions
                # add_permission(data, wb_id[0], user_id, version, auth_token)

                # Step: Create New Schedule
                # createSchedule(server)

            # Step: Sign Out to the Tableau Server
            server.auth.sign_out()

    except Exception as e:
        print("Something went wrong, Error occured.\n", e)
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

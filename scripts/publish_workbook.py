import os
import json
import argparse
import tableauserverclient as TSC
import requests
import xml.etree.ElementTree as ET

# permissions = {"Read", "Write", "Filter", "AddComment", "ViewComments", "ShareView", "ExportData", "ViewUnderlyingData",
#                "ExportImage", "Delete", "ChangeHierarchy", "ChangePermissions", "WebAuthoring", "ExportXml"}


def raiseError(e, file_path):
    print(f"{file_path} workbook is not published.")
    raise LookupError(e)
    exit(1)


def signin(data):
    tableau_auth = TSC.TableauAuth(
        args.username, args.password, None if data['is_site_default'] else data['site_name'])
    server = TSC.Server(data['server_url'], use_server_version=True)
    server.auth.sign_in(tableau_auth)
    return server


def getProject(server, data):
    all_projects, pagination_item = server.projects.get()
    project = next(
        (project for project in all_projects if project.name == data['project_path']), None)

    if project.id is not None:
        return project.id
    else:
        raiseError(
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


def updateProjectPermissions(server, project_path):

    all_projects, pagination_item = server.projects.get()
    project = next(
        (project for project in all_projects if project.name == project_path), None)
    print(f"project name:{project.name} and id: {project.id}")

    # Query for existing workbook default-permissions
    server.projects.populate_workbook_default_permissions(project)

    for default_permissions in project.default_workbook_permissions:
        # Update permisssion
        new_capabilities = {
            TSC.Permission.Capability.AddComment: TSC.Permission.Mode.Deny,
        }

        new_rules = [TSC.PermissionsRule(
            grantee=default_permissions.grantee, capabilities=new_capabilities)]

        new_default_permissions = server.projects.update_workbook_default_permissions(
            project, new_rules)

    # Print result from adding a new default permission
    for permission in new_default_permissions:
        grantee = permission.grantee
        capabilities = permission.capabilities
        print(f"\nCapabilities for {grantee.tag_name} {grantee.id}:")

        for capability in capabilities:
            print(f"\t{capability} - {capabilities[capability]}")


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


def getUserID(server, data):
    all_users, pagination_item = server.users.get()
    print([user.name for user in all_users])
    return [user.id for user in all_users if user.name == data['user_name']]


def add_permission(data, workbook_id, user_id):
    url = f"https://tableau.devinvh.com/api/3.15/sites/{data['site_id']}/workbooks/{workbook_id}/permissions"

    # Build the request
    xml_request = ET.Element('tsRequest')
    permissions_element = ET.SubElement(xml_request, 'permissions')
    ET.SubElement(permissions_element, 'workbook', id=workbook_id)
    grantee_element = ET.SubElement(permissions_element, 'granteeCapabilities')
    ET.SubElement(grantee_element, 'user', id=user_id)
    capabilities_element = ET.SubElement(grantee_element, 'capabilities')
    ET.SubElement(capabilities_element, 'capability',
                  name=data['permission_name'], mode=data['permission_mode'])
    xml_request = ET.tostring(xml_request)

    print(xml_request)
    server_request = requests.put(url, data=xml_request, headers={
                                  'x-tableau-auth': data['auth_token']})
    print(server_request)


def main(args):
    project_data_json = json.loads(args.project_data)
    try:
        for data in project_data_json:
            # Step: Sign in to Tableau server.
            server = signin(data)

            if data['project_path'] is None:
                raiseError(
                    f"The project project_path field is Null in JSON Template.", data['file_path'])
            else:
                # Step: Form a new workbook item and publish.
                # publishWB(server, data)

                # Step: Get the Workbook ID from the Workbook Name
                wb_id = getWBID(server, data)
                print(wb_id)

                # Step: Get the User ID from the User Name
                user_id = getUserID(server, data)
                print(user_id)

                add_permission(data, wb_id, user_id)
                # Step: Update Project permissions
                # updateProjectPermissions(server, data['project_path'])

                # Step: Create New Schedule
                # createSchedule(server)

            # Step: Sign Out to the Tableau Server
            server.auth.sign_out()

    except Exception as e:
        print("Workbook not published.\n", e)
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

import sys
import requests
import getopt
import json
import urllib.parse
from veracode_api_signing.plugin_requests import RequestsAuthPluginVeracodeHMAC
import openpyxl
import time
import xml.etree.ElementTree as ET  # for parsing XML

from veracode_api_signing.credentials import get_credentials

class NoExactMatchFoundException(Exception):
    message=""
    def __init__(self, message_to_set):
        self.message = message_to_set

    def get_message(self):
        return self.message

NULL = "NULL"

json_headers = {
    "User-Agent": "Bulk application creation - python script",
    "Content-Type": "application/json"
}

xml_headers = {
    "User-Agent": "Bulk application creation - python script",
    "Content-Type": "application/xml"
}

failed_attempts = 0
max_attempts_per_request = 10
sleep_time = 10

last_column = 0
non_custom_field_headers={"Application Name",
                            "Business Criticality",
                            "Policy",
                            "Description",
                            "Tags",
                            "Business Unit",
                            "Business Owner",
                            "Owner Email",
                            "Teams",
                            "Dynamic Scan Approval",
                            "Archer Application Name"}

def print_help():
    """Prints command line options and exits"""
    print("""bulk-update-applications.py -f <excel_file_with_application_definitions> [-r <header_row>] [-d]"
        Reads all lines in <excel_file_with_application_definitions>, for each line, it will update the profile
        <header_row> defines which row contains your table headers, which will be read to determine where each field goes (default 2).
""")
    sys.exit()

def request_encode(value_to_encode):
    return urllib.parse.quote(value_to_encode, safe='')

def find_exact_match(list, to_find, field_name, list_name2):
    if list_name2:
        for index in range(len(list)):
            if (list_name2 and list[index][list_name2][field_name].lower() == to_find.lower()):
                return list[index]
    
        print(f"Unable to find a member of list with {field_name}+{list_name2} equal to {to_find}")
        raise NoExactMatchFoundException(f"Unable to find a member of list with {field_name}+{list_name2} equal to {to_find}")
    else:
        for index in range(len(list)):
            if list[index][field_name].lower() == to_find.lower():
                return list[index]

        print(f"Unable to find a member of list with {field_name} equal to {to_find}")
        raise NoExactMatchFoundException(f"Unable to find a member of list with {field_name} equal to {to_find}")

def get_field_value(excel_headers, excel_sheet, row, field_header):
    field_to_get = field_header.strip()
    if field_to_get in excel_headers:
        field_value = excel_sheet.cell(row = row, column = excel_headers[field_to_get]).value
        if field_value:
            return field_value
    return ""

def get_business_owners(excel_headers, excel_sheet, row):
    name=get_field_value(excel_headers, excel_sheet, row, "Business Owner")
    email=get_field_value(excel_headers, excel_sheet, row, "Owner Email")
    if not name or not email:
        return ""
    elif name == NULL or email == NULL:
        return '''
            "business_owners": [
                {}
            ]'''
    else:
        return f'''
            "business_owners": [
                {{
                    "email": "{email}",
                    "name": "{name}"
                }}
            ]'''

def get_field_for_json(field_value, field_name):
    if not field_value:
        return None
    if field_value == NULL:
        field_value = ""
    return f'"{field_name}": "{field_value}"'

def get_item_from_api_call(api_base, api_to_call, item_to_find, list_name, list_name2, field_to_check, field_to_get, is_exact_match, verbose):
    global failed_attempts
    global sleep_time
    global max_attempts_per_request
    path = f"{api_base}{api_to_call}"
    if verbose:
        print(f"Calling: {path}")

    response = requests.get(path, auth=RequestsAuthPluginVeracodeHMAC(), headers=json_headers)
    data = response.json()

    if response.status_code == 200:
        if verbose:
            print(data)
        if "_embedded" in data and len(data["_embedded"][list_name]) > 0:
            if list_name2:
                return (find_exact_match(data["_embedded"][list_name], item_to_find, field_to_check, list_name2) if is_exact_match else data["_embedded"][list_name][list_name2][0])[field_to_get]
            else:
                return (find_exact_match(data["_embedded"][list_name], item_to_find, field_to_check, list_name2) if is_exact_match else data["_embedded"][list_name][0])[field_to_get]
        else:
            print(f"ERROR: No {list_name}+{list_name2} named '{item_to_find}' found")
            return f"ERROR: No {list_name}+{list_name2} named '{item_to_find}' found"
    else:
        print(f"ERROR: trying to get {list_name}+{list_name2} named {item_to_find}")
        print(f"ERROR: code: {response.status_code}")
        print(f"ERROR: value: {data}")
        failed_attempts+=1
        if (failed_attempts < max_attempts_per_request):
            time.sleep(sleep_time)
            return get_item_from_api_call(api_base, api_to_call, item_to_find, list_name, list_name2, field_to_check, field_to_get, verbose)
        else:
            return f"ERROR: trying to get {list_name}+{list_name2} named {item_to_find}"

def get_business_unit(api_base, excel_headers, excel_sheet, row, verbose):
    business_unit_name = get_field_value(excel_headers, excel_sheet, row, "Business Unit")
    if not business_unit_name:
        return ""
    elif business_unit_name == NULL:
        return '''"business_unit": null'''
    else:
        return f'''
    "business_unit": {{
      "guid": "{get_item_from_api_call(api_base, "api/authn/v2/business_units?bu_name="+ request_encode(business_unit_name), business_unit_name, "business_units", None, "bu_name", "bu_id", True, verbose)}"
    }}'''

def get_policy(api_base, excel_headers, excel_sheet, row, verbose):
    policy_name = get_field_value(excel_headers, excel_sheet, row, "Policy")
    if not policy_name:
        return ""
    elif policy_name == NULL:
        return '"policies": []'
    else:
        return f'''
    "policies": [{{
      "guid": "{get_item_from_api_call(api_base, "appsec/v1/policies?category=APPLICATION&name_exact=true&public_policy=true&name="+ request_encode(policy_name), policy_name, "policy_versions", None, "name", "guid", False, verbose)}"
    }}]'''

def get_team_value(api_base, team_name, verbose):
    team_guid = get_item_from_api_call(api_base, "api/authn/v2/teams?all_for_org=true&team_name="+ request_encode(team_name), team_name, "teams", None, "team_name", "team_id", True, verbose)
    if team_guid:
        return f'''{{
        "guid": "{team_guid}"
      }}'''
    else:
        return ""

def get_teams(api_base, excel_headers, excel_sheet, row, verbose):
    all_teams_base = get_field_value(excel_headers, excel_sheet, row, "Teams")
    if not all_teams_base:
        return ""
    if all_teams_base == NULL:
        return f'''
            "teams": []'''
    all_teams = all_teams_base.split(",")
    inner_team_list = ""
    for team_name in all_teams:
        team_value = get_team_value(api_base, team_name.strip(), verbose)
        if team_name: 
            inner_team_list = inner_team_list + (""",
            """ if inner_team_list else "") + team_value
    if inner_team_list:
        return f'''
            "teams": [
                {inner_team_list}
            ]'''
    else:
        return ""
        
def get_application_settings(excel_headers, excel_sheet, row):
    base_value = get_field_value(excel_headers, excel_sheet, row, "Dynamic Scan Approval")
    value = False
    if base_value:
        value = str(base_value).strip().lower() == "false"
    inner_field = get_field_for_json(str(value).lower(), "dynamic_scan_approval_not_required")
    return f'''
    "settings": {{
      {inner_field}
    }}''' if inner_field else None

def get_custom_fields(excel_headers, excel_sheet, row):
    global non_custom_field_headers
    inner_custom_fields_list = ""
    for field in excel_headers:
        if not field in non_custom_field_headers:
            value = excel_sheet.cell(row = row, column=excel_headers[field]).value
            if value:
                found_field = f'''{{
                    "name": "{field}",
                    {get_field_for_json(value, "value")}
                }}'''
                inner_custom_fields_list = inner_custom_fields_list + (""",
                """ if inner_custom_fields_list else "") + found_field
    if inner_custom_fields_list:
        return f'''
                "custom_fields": [
                    {inner_custom_fields_list}
                ]'''
    else:
        return ""
        
def get_archer_application_name(excel_headers, excel_sheet, row):
    return get_field_for_json(get_field_value(excel_headers, excel_sheet, row, "Archer Application Name"), "archer_app_name")

def url_encode_with_plus(a_string):
    return urllib.parse.quote_plus(a_string, safe='').replace("&", "%26")

def get_error_node_value(body):
    inner_node = ET.XML(body)
    if inner_node.tag == "error" and not inner_node == None:
        return inner_node.text
    else:
        return ""

def get_application_guid(api_base, application_name, verbose):
    return get_item_from_api_call(api_base, "appsec/v1/applications?name="+ request_encode(application_name.strip()), application_name.strip(), "applications", "profile", "name", "guid", True, verbose)

def get_description(excel_headers, excel_sheet, row):
    return get_field_for_json(get_field_value(excel_headers, excel_sheet, row, "Description"), "description")

def get_tags(excel_headers, excel_sheet, row):
    return get_field_for_json(get_field_value(excel_headers, excel_sheet, row, "Tags"), "tags")

def get_business_criticality(excel_headers, excel_sheet, row):
    return get_field_for_json(get_field_value(excel_headers, excel_sheet, row, "Business Criticality").replace(" ", "_").upper(), "business_criticality")

def get_inner_profile_info(api_base, application_name, excel_headers, excel_sheet, row, verbose):
    inner_profile_info = ""
    business_criticality_json = get_business_criticality(excel_headers, excel_sheet, row)
    archer_application_name_json = get_archer_application_name(excel_headers, excel_sheet, row)
    business_owners_json = get_business_owners(excel_headers, excel_sheet, row)
    business_unit_json = get_business_unit(api_base, excel_headers, excel_sheet, row, verbose)
    description_json = get_description(excel_headers, excel_sheet, row)
    policy_json = get_policy(api_base, excel_headers, excel_sheet, row, verbose)
    tags_json = get_tags(excel_headers, excel_sheet, row)
    teams_json = get_teams(api_base, excel_headers, excel_sheet, row, verbose)
    application_settings_json = get_application_settings(excel_headers, excel_sheet, row)
    custom_fields_json=get_custom_fields(excel_headers, excel_sheet, row)

    inner_profile_info = f'''"name": "{application_name}"'''
    if business_criticality_json:
        inner_profile_info = inner_profile_info + ", " + business_criticality_json
    if archer_application_name_json:
        inner_profile_info = inner_profile_info + ", " + archer_application_name_json
    if business_owners_json:
        inner_profile_info = inner_profile_info + ", " + business_owners_json
    if business_unit_json:
        inner_profile_info = inner_profile_info + ", " + business_unit_json
    if description_json:
        inner_profile_info = inner_profile_info + ", " + description_json
    if policy_json:
        inner_profile_info = inner_profile_info + ", " + policy_json
    if tags_json:
        inner_profile_info = inner_profile_info + ", " + tags_json
    if teams_json:
        inner_profile_info = inner_profile_info + ", " + teams_json
    if application_settings_json:
        inner_profile_info = inner_profile_info + ", " + application_settings_json
    if custom_fields_json:
        inner_profile_info = inner_profile_info + ", " + custom_fields_json

    return inner_profile_info

def update_application(api_base, excel_headers, excel_sheet, row, verbose):
    application_name = get_field_value(excel_headers, excel_sheet, row, "Application Name")
    application_guid = get_application_guid(api_base, application_name, verbose)
    if not application_guid:
        error_message = f"Application not found: {application_name}"
        print (error_message)
        return error_message
    
    path = f"{api_base}appsec/v1/applications/{application_guid}?method=partial"
    request_content=f'''{{
        "profile": {{
            {get_inner_profile_info(api_base, application_name, excel_headers, excel_sheet, row, verbose)}
        }}
    }}'''
    
    if verbose:
        print(f"Calling API at: {path}")
        print(request_content)

    response = requests.patch(path, auth=RequestsAuthPluginVeracodeHMAC(), headers=json_headers, json=json.loads(request_content))

    if verbose:
        print(f"status code {response.status_code}")
        body = response.json()
        if body:
            print(body)
    if response.status_code == 200:
        print(f"Successfully updated profile {application_name}.")
        return "success"
    else:
        body = response.json()
        if (body):
            return f"Unable to update application profile: {response.status_code} - {body}"
        else:
            return f"Unable to update application profile: {response.status_code}"
    

def setup_excel_headers(excel_sheet, header_row, verbose):
    excel_headers = {}
    global last_column
    for column in range(1, excel_sheet.max_column+1):
        cell = excel_sheet.cell(row = header_row, column = column)
        if not cell or cell is None or not cell.value or cell.value.strip() == "":
            break
        to_add = cell.value
        if to_add:
            to_add = str(to_add).strip()
        if verbose:
            print(f"Adding column {column} for value {to_add}")
        excel_headers[to_add] = column
        last_column += 1
    return excel_headers

def update_all_applications(api_base, file_name, header_row, verbose):
    global failed_attempts
    excel_file = openpyxl.load_workbook(file_name)
    excel_sheet = excel_file.active
    try:
        excel_headers = setup_excel_headers(excel_sheet, header_row, verbose)
        print("Finished reading excel headers")
        if verbose:
            print("Values found are:")
            print(excel_headers)

        max_column=len(excel_headers)
        for row in range(header_row+1, excel_sheet.max_row+1):      
            failed_attempts = 0
            if verbose:
                for field in excel_headers:
                    print(f"Found column with values:")
                    print(f"{field} -> {excel_sheet.cell(row = row, column=excel_headers[field]).value}")
            status=excel_sheet.cell(row = row, column = max_column+1).value
            if (status == 'success'):
                print("Skipping row as it was already done")
            else:
                try:
                    print(f"Importing row {row-header_row}/{excel_sheet.max_row-header_row}:")
                    status = update_application(api_base, excel_headers, excel_sheet, row, verbose)
                    print(f"Finished importing row {row-header_row}/{excel_sheet.max_row-header_row}")
                    print("---------------------------------------------------------------------------")
                except NoExactMatchFoundException:
                    status= NoExactMatchFoundException.get_message()
                excel_sheet.cell(row = row, column = max_column+1).value=status
    finally:
        excel_file.save(filename=file_name)

def get_api_base():
    api_key_id, api_key_secret = get_credentials()
    api_base = "https://api.veracode.{instance}/"
    if api_key_id.startswith("vera01"):
        return api_base.replace("{instance}", "eu", 1)
    else:
        return api_base.replace("{instance}", "com", 1)

def main(argv):
    """Allows for bulk updating application profiles"""
    global failed_attempts
    global last_column
    excel_file = None
    try:
        verbose = False
        file_name = ''
        header_row = -1

        opts, args = getopt.getopt(argv, "hdf:r:", ["file_name=", "header_row="])
        for opt, arg in opts:
            if opt == '-h':
                print_help()
            if opt == '-d':
                verbose = True
            if opt in ('-f', '--file_name'):
                file_name=arg
            if opt in ('-r', '--header_row'):
                header_row=int(arg)

        api_base = get_api_base()

        if header_row < 0:
            print("INFO: No header row set, using default of 2")
            header_row = 2
        if file_name > 0:
            update_all_applications(api_base, file_name, header_row, verbose)
        else:
            print_help()
    except requests.RequestException as e:
        print("An error occurred!")
        print(e)
        sys.exit(1)
    finally:
        if excel_file:
            excel_file.save(filename=file_name)


if __name__ == "__main__":
    main(sys.argv[1:])

# Veracode Bulk Application Update

## Overview

This script allows for bulk updating application profiles in the Veracode platform

## Installation

Clone this repository:

    git clone https://github.com/cadonuno/Veracode-Bulk-Application-Update.git

Install dependencies:

    cd Veracode-Bulk-Application-Creator
    pip install -r requirements.txt

### Getting Started

It is highly recommended that you store veracode API credentials on disk, in a secure file that has 
appropriate file protections in place.

(Optional) Save Veracode API credentials in `~/.veracode/credentials`

    [default]
    veracode_api_key_id = <YOUR_API_KEY_ID>
    veracode_api_key_secret = <YOUR_API_KEY_SECRET>

### Preparing the Excel Template
    The Excel template present in the repository can be used to prepare the metadata.

### After execution:
**A new column will be added to the right containing the status of each line. Successfully updated rows will be set to 'success' and will be skipped on future runs.**
    
### Running the script
    py bulk-update-applications.py -f <excel_file_with_application_definitions> [-r <header_row>] [-d]"
        Reads all lines in <excel_file_with_application_definitions>, for each line, it will update the profile
        <header_row> defines which row contains your table headers, which will be read to determine where each field goes (default 2).

If a credentials file is not created, you can export the following environment variables:

    export VERACODE_API_KEY_ID=<YOUR_API_KEY_ID>
    export VERACODE_API_KEY_SECRET=<YOUR_API_KEY_SECRET>
    python bulk-create-applications.py -f <excel_file_with_application_definitions> -r <header_row> [-d]

## License

[![MIT license](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

See the [LICENSE](LICENSE) file for details

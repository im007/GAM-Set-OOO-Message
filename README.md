# GAM-Set-OOO-Message

A GUI application that allows you to set customized Out of Office (OOO) messages for multiple users using GAM (Google Apps Manager).

## Features

- User-friendly graphical interface
- CSV file support for batch processing
- Customizable message templates with variables
- Live message preview
- Status logging
- CSV data preview

## Prerequisites

- GAM (Google Apps Manager) installed and configured
- Python 3.x with the following packages:
  - tkinter
  - pandas
  - csv

## CSV File Format

The CSV file must contain the following columns:
- `email`: The user's email address
- `name`: The user's full name

Example:
```csv
email,name
johnsmith@company.com,John Smith
```

## Usage

1. Launch the application by running `GAM_set_OOO.py`
2. Select your CSV file containing user information
3. Verify or select the GAM installation directory
4. Enter your company settings:
   - Company Name
   - Contact Email
5. Customize the OOO message template:
   - Subject line
   - Message body
6. Review the message preview
7. Click "Set OOO Messages" to process all users

## Message Template Variables

The following variables can be used in both subject and message:
- `{name}`: User's full name
- `{company_name}`: Company name
- `{contact_email}`: Contact email address (automatically formatted as HTML link)

## Notes

- The application automatically attempts to detect your GAM installation
- All actions are logged in the status window
- HTML formatting is supported in the message body
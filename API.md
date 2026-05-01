## API Documentation - Circuit Building Application


## Base URL
Development: http://localhost:5000
Production: (http://jbdrawing.cellapps.com:5000/)

## Authentication
All API endpoints (except /login) require authentication via Flask-Login session cookies.

## Login
http
POST /login
Content-Type: application/x-www-form-urlencoded

username=admin&password=yourpassword
Response: Redirects to dashboard or specified next page

## Logout
http
GET /logout
Response: Redirects to login page

## User Management (Admin Only - Role 4)
## List All Users
http
GET /admin/users
Response: HTML page with user management interface

## Add New User
http
POST /admin/users/add
Content-Type: application/x-www-form-urlencoded

username=newuser&mobile_number=1234567890&email=user@example.com&password=password123&role=1&designation=level1&project_ids[]=1&project_ids[]=2
Required Fields: username, password, role, designation

## Edit User
http
POST /admin/users/edit/<user_id>
Content-Type: application/x-www-form-urlencoded

username=updateduser&mobile_number=0987654321&email=updated@example.com&password=newpassword&role=2&designation=level2&is_active=1&remarks=Updated&project_ids[]=1

## Delete User
http
POST /admin/users/delete/<user_id>

## Project Management
## List Projects
http
GET /
Response: Dashboard with all accessible projects

## Create New Project
http
POST /new_project
Content-Type: application/x-www-form-urlencoded

name=NewStation&description=Station description

## Edit Project
http
POST /project/<project_id>/edit
Content-Type: application/x-www-form-urlencoded

name=UpdatedStation&description=Updated description

## Delete Project
http
POST /project/<project_id>/delete

## Switch Project
http
GET /project/<project_id>/switch
Response: Sets session project and redirects to dashboard

## Check Station Name Availability
http
GET /api/check_station_name?name=StationName
Response:

json
{"exists": true}

## Excel Upload & PDF Generation
## Upload Excel File
http
POST /excel_to_pdf
Content-Type: multipart/form-data

file=@drawing.xlsx&remarks=Initial drawing
Query Parameters:

new=1 - Force create new project from Excel

Response: Redirects to PDF view page

## Download Project Data as Excel
http
GET /download/
Response: Downloads XLSX file with all project data

## Workflow Management
## Workflow Steps Navigation
http
GET /workflow/step/<step_number>
Steps:

## Station Information (StationDrawing)

## Location Box or CTR Setup (junction_box)

## Cable Information (cable)

## Terminal Details (terminal)

## Terminal Headers (terminal_header)

## Group Information (group)

## Choke Table (choketable)

## Resistor Table (resistortable)

## Relay Box (cable_box)

## Upload Excel Sheet
http
POST /upload/<sheet_name>
Content-Type: multipart/form-data

file=@data.xlsx
Available sheets: StationDrawing, junction_box, cable, terminal, terminal_header, group, choketable, resistortable

## AJAX Data Operations
## Add Cable
http
POST /add_cable_ajax
Content-Type: application/json

{
  "cable_id": "1",
  "cable_name": "Cable A",
  "junction_name": "Junction 1",
  "row": "A",
  "position": "1",
  "junction_box": "JB1",
  "terminal": "12",
  "start_no": "1",
  "cable_type": "cable"
}
## Add Terminal
http
POST /add_terminal_ajax
Content-Type: application/json

{
  "cable_id": "1",
  "terminal_id": "1",
  "terminal_no": "01",
  "symbol": "ara/wago",
  "input_left": "L1",
  "input_right": "N",
  "spare": "N",
  "input_connected": "Y",
  "output_connected": "Y",
  "input_connected_extra": "",
  "output_connected_extra": "",
  "output_left": "",
  "output_right": ""
}
## Add Group Configuration
http
POST /add_group_ajax
Content-Type: application/json

{
  "cable_id": "1",
  "group_id": "G1",
  "terminal_no": "1,2,3",
  "input_output": "input",
  "text": "Group description"
}
## Add Choke Data
http
POST /add_choke_ajax
Content-Type: application/json

{
  "cable_id": "1",
  "choke_id": "CH1",
  "input_terminal": "1",
  "output_terminal": "2",
  "terminal_name": "Choke 1"
}
## Add Resistor Data
http
POST /add_resistor_ajax
Content-Type: application/json

{
  "cable_id": "1",
  "resistor_id": "R1",
  "input_terminal": "1",
  "output_terminal": "2",
  "resistor_name": "Resistor 1"
}
## Add Terminal Header
http
POST /add_header_ajax
Content-Type: application/json

{
  "cable_id": "1",
  "header_type": "header1",
  "terminal_start": "1",
  "terminal_end": "12",
  "input_output": "input",
  "text": "Header description"
}
## Add Junction Boxes (Multiple)
http
POST /add_junctions_ajax
Content-Type: application/json

{
  "junctions": [
    {
      "station_id": "1",
      "junction_id": "1",
      "junction_name": "Location A",
      "junction_size": "Full",
      "latitude": "23.0225",
      "longitude": "72.5714",
      "row": "A"
    }
  ]
}
## Data Query Endpoints
## Get Existing Cables
http
GET /get_existing_cables
Response:

json
[
  {
    "id": 1,
    "cable_id": "1",
    "cable_name": "Cable A",
    "junction_name": "Junction 1",
    "row": "A",
    "junction_box": "JB1",
    "position": "1",
    "terminal": "12",
    "start_no": "1",
    "cable_type": "cable"
  }
]
## Get Terminals for Cable
http
GET /get_terminals_for_cable?cable_id=1
## Get Groups for Cable
http
GET /get_groups_for_cable?cable_id=1
## Get Chokes for Cable
http
GET /get_chokes_for_cable?cable_id=1
## Get Resistors for Cable
http
GET /get_resistors_for_cable?cable_id=1
## Check Terminal Duplicate
http
GET /check_terminal_duplicate?cable_id=1&terminal_id=1
Response:

json
{"exists": true}
## Check Cable Terminals
http
GET /check_cable_terminals/1
Response:

json
{"exists": true, "count": 12}
## PDF Management
## View PDF
http
GET /pdf/view/<filename>
Response: HTML page with PDF viewer

## Inline PDF Display
http
GET /pdf/inline/<filename>
Response: PDF file for inline display

## Download PDF
http
GET /download_pdf/<filename>
Response: PDF file download

## Continue from Previous Version
http
POST /continue_from_version/<pdf_id>
Content-Type: application/json

{}
Response: Loads XLSX data from PDF version and redirects to workflow

## Set Continue Drawing Flag
http
POST /set_continue_drawing
## Set New Drawing Flag
http
POST /set_new_drawing
## Approval System
## Approval Tracking Dashboard
http
GET /
Query Parameters:

page - Page number (default: 1)

rows_per_page - Rows per page (default: 10)

project_id - Filter by project

junction_id - Filter by junction

start_date - Start date filter (YYYY-MM-DD)

end_date - End date filter (YYYY-MM-DD)

approval_status - Filter by status (all, approved, rejected, pending)

show_without_drawings - Show projects without drawings (on/off)

latest_only - Show only latest versions (true/false)

## Approve PDF
http
POST /approve_pdf/<pdf_id>/<level>
Content-Type: application/x-www-form-urlencoded

approval_remarks=Approved after review
Levels: 1 (Level 1), 2 (Level 2), 3 (Level 3)

## Reject PDF
http
POST /reject_pdf/<pdf_id>/<level>
Content-Type: application/x-www-form-urlencoded

rejection_reason=Incorrect terminal numbering
## Get Approval Statistics
http
GET /get_approval_stats
Response:

json
{
  "success": true,
  "stats": {
    "total": 45,
    "approved": 30,
    "rejected": 5,
    "pending": 10,
    "level1_pending": 3,
    "level2_pending": 4,
    "level3_pending": 3,
    "recent_approvals": 12,
    "user_approvals": 5,
    "user_rejections": 1,
    "user_pending_approvals": 2
  }
}
## Get Approval History
http
GET /get_approval_history/<pdf_id>
Response:

json
{
  "success": true,
  "pdf_id": 101,
  "pdf_filename": "RailwayProject_1_v1.pdf",
  "project_name": "Railway Project Alpha",
  "version": 1,
  "created_at": "2024-01-10T09:00:00+05:30",
  "current_status": {
    "level1": {
      "status": "approved",
      "approver": "admin",
      "date": "2024-01-10T10:00:00+05:30"
    }
  },
  "approval_history": [...]
}
## Notification System
## Get Notifications
http
GET /notifications
Response:

json
{
  "all_notifications": [
    {
      "id": 1,
      "pdf_id": 101,
      "project_id": 1,
      "project_name": "Railway Project Alpha",
      "level": "level1",
      "status": "pending",
      "message": "NEW DRAWING requires your approval",
      "is_read": false,
      "created_at": "2024-01-10T09:15:00+05:30"
    }
  ],
  "recent_notifications": [...]
}
## Mark Notification as Read
http
POST /mark_notification_read/<notification_id>
## Mark All Notifications as Read
http
POST /mark_all_notifications_read
## Dismiss Notification
http
POST /dismiss_notification/<notification_id>
## Dismiss All Notifications
http
POST /dismiss_all_notifications
## Clear All Notifications
http
POST /clear_all_notifications
## Get Notification Count
http
GET /notification_count
Response:

json
{"count": 5}
## Station Master (Admin Only)
## View Station Master
http
GET /admin/station_master
Query Parameters:

project_id - Filter by project

station_name - Filter by station name

station_code - Filter by station code

## Sync All Stations
http
POST /admin/sync_all_stations
## Data Management
## Clear Project Data
http
POST /clear_current_project
## New Drawing (Clear and Start Fresh)
http
POST /new_drawing/<project_id>
Response:

json
{"success": true, "message": "Drawing cleared successfully!"}
## Set Project and Continue
http
POST /set_project_and_continue/<project_id>
Response:

json
{
  "success": true,
  "redirect_url": "/workflow/step/2"
}
## New Drawing Selection Page
http
GET /new_drawing_selection
## Admin Endpoints
## Get Designations by Role
http
GET /admin/designations_by_role/<role_id>
Response:

json
{
  "designations": [
    {
      "id": 1,
      "name": "level1",
      "level": 1
    }
  ]
}
## Test Approval Logic
http
GET /test_approval/<pdf_id>
Response:

json
{
  "pdf_id": 101,
  "level1_status": "pending",
  "user_designation": "level1",
  "can_level1_approve": true
}
## Template Filters
## Sort Cables
Usage in templates: {{ cables|sort_cables }}

## Sort Junction Boxes
Usage in templates: {{ junction_boxes|sort_junction_boxes }}

## Sort Terminals
Usage in templates: {{ terminals|sort_terminals }}

## Sort Terminal Headers
Usage in templates: {{ headers|sort_terminal_headers }}

## Workflow Edit/Delete Operations
## Edit Row in Workflow
http
GET /workflow/<sheet_name>/edit/<row_id>/<step>
Example: /workflow/junction_box/edit/1/2

## Delete Row in Workflow
http
POST /workflow/<sheet_name>/delete/<row_id>/<step>
## Sheet Operations
## View Sheet Form
http
GET /sheet/<name>
Available names: StationDrawing, junction_box, cable, terminal, terminal_header, group, choketable, resistortable, cable_box

## Edit Row in Sheet
http
GET /sheet/<name>/edit/<row_id>

## Delete Row in Sheet
http
POST /sheet/<name>/delete/<row_id>
## Data Preview
## Preview All Data
http
GET /preview
Response: HTML page showing all data tables for current project

## Error Responses
## 400 Bad Request
json
{
  "success": false,
  "message": "Invalid request parameters"
}
## 401 Unauthorized
Redirects to login page with flash message

## 403 Forbidden
Flash message: "Access denied. Admin privileges required."

## 404 Not Found
Flash message: "Resource not found"

## 500 Internal Server Error
Flash message with error details

## Role-Based Access Control
Role Number	Role Name	Permissions
0	Viewer	View approvals only
1	Creator	Create drawings, approve at Level 1
2	Approver L2	Approve at Level 2
3	Approver L3	Approve at Level 3
4	Admin	Full access to all features
## Session Variables
current_project_id - Currently selected project ID

project_id - Alternate key for project ID

is_continue_drawing - Flag for continue drawing mode

junction_count - Number of junction boxes to add

show_more_junction_grid - Flag for showing junction grid

## File Upload Configuration
Allowed File Types: .xlsx only

Upload Directory: ./uploads/

Max File Size: Configured in Flask app

FTP Upload: Optional, configured via FTP_ENABLED environment variable

## Database Models Used
User, Project, StationDrawing, JunctionBox, Cable, Terminal

Group, TerminalHeader, ChokeTable, ResistorTable, CableBox

GeneratedPDF, Approval, Notification

StationMaster, RoleMaster, DesignationMaster

## FTP Operations
Test FTP Connection
http
GET /test_ftp_connection
Tests FTP/SFTP connectivity and directory listing.

Response: JSON with connection status and directory contents

Success Response:

json
{
  "success": true,
  "protocol": "SFTP",
  "message": "Connected successfully to ftp.example.com:22",
  "directory": "/uploads",
  "contents": ["file1.txt", "file2.pdf"]
}
## Check FTP Status
http
GET /check_ftp_status
Checks FTP server connectivity.

Response: JSON with connection status

Success Response:

json
{"success": true, "message": "SFTP connection successful"}
## Upload File to FTP
http
POST /upload_to_ftp/<filename>
Manually upload a file to FTP server.

Parameters: filename - Name of file in uploads directory

Request Body: None

Response: JSON with upload status

Success Response:

json
{
  "success": true,
  "message": "File uploaded successfully. File archived to ...",
  "remote_path": "/uploads/filename.pdf"
}
## Sync All Project Files to FTP
http
POST /sync_all_to_ftp/<int:project_id>
Sync all XLSX and PDF files for a project to FTP.

Parameters: project_id - Project ID

Response: JSON with sync results

Success Response:

json
{
  "success": true,
  "message": "Synced 5/7 files to FTP",
  "results": [
    {"filename": "file1.pdf", "success": true, "message": "File uploaded successfully"},
    ...
  ]
}
## Debug FTP Upload
http
POST /debug_ftp_upload
Debug FTP upload with a test file.

Response: JSON with test upload details

Success Response:

json
{
  "success": true,
  "message": "FTP connection successful! File uploaded to: /uploads/test.txt",
  "config": {
    "host": "ftp.example.com",
    "port": 21,
    "username": "user",
    "upload_dir": "/uploads",
    "use_sftp": false
  }
}
## Draft Management
Save Junction Box Draft
http
POST /save_junction_box_draft
Save junction box data as draft.

Request Body: JSON with junction boxes array

json
{
  "junctions": [
    {
      "station_id": "ST001",
      "junction_id": "J1",
      "junction_name": "Junction 1",
      "junction_size": "Full",
      "junction_row": "1",
      "latitude": "23.0225",
      "longitude": "72.5714"
    }
  ]
}
Response: JSON with save status

Success Response:

json
{
  "success": true,
  "message": "Draft saved successfully (5 junction boxes)",
  "total_junction_boxes": 15,
  "saved_junction_ids": ["J1", "J2", ...]
}
## Get Cable Configuration
http
GET /get_cable_configuration
Get saved cable configuration for a junction box.

Query Parameters:

project_id - Project ID

junction_box_id - Junction box ID

Response: JSON with configuration rows

Success Response:

json
{
  "success": true,
  "config_rows": [
    {
      "row_number": "1",
      "location_row_name": "A",
      "cable_type": "cable",
      "number_of_cables": 6
    }
  ]
}
## Save Cable Configuration
http
POST /save_cable_configuration
Save final cable configuration.

Request Body: JSON with configuration

json
{
  "project_id": 1,
  "junction_box_id": "J1",
  "config_rows": [
    {
      "row_number": "1",
      "location_row_name": "A",
      "cable_type": "cable",
      "number_of_cables": 6
    }
  ]
}
Response: JSON with save status

## Manage Cable Configuration Draft
http
GET /get_cable_config_draft
POST /clear_cable_config_draft
Get or clear cable configuration draft.

Parameters: project_id, junction_box_id

Response: JSON with draft data or clear status

## Save More Junctions (Final)
http
POST /save_more_junctions
Save additional junction boxes (final save).

Request Body: Form data with multiple junctions

Response: Redirects with flash message

## Cable Table Management
## Save Cable Table Draft
http
POST /save_cable_table_draft
Save cable table configuration - UPDATE IF EXISTS, INSERT IF NEW.

Request Body: JSON with cable data

json
{
  "junction_box_id": "J1",
  "junction_box_name": "Junction 1",
  "cable_data": [
    {
      "row": "A",
      "position": "1",
      "terminal": "12",
      "start_no": "1",
      "cable_id": "J1-A-1",
      "cable_name": "Cable A1"
    }
  ]
}
Response: JSON with save status and statistics

## Finalize Cable Table
http
POST /finalize_cable_table
Finalize cable table (updates summary only).

Request Body: JSON with junction box ID

Response: JSON with finalization status

## Manage Cable Table Draft
http
GET /get_cable_table_draft
POST /clear_cable_table_draft
Load or clear cable table draft from Cable table.

Parameters: project_id, junction_box_id

Response: JSON with draft data or clear status

## Get Cable Summary/Stats
http
GET /get_cable_summary
GET /get_cable_stats
Get cable summary and detailed statistics.

Response: JSON with counts and statistics

## Terminal Management
## Save Terminal Draft
http
POST /save_terminal_draft
Save terminal configuration directly to Terminal table.

Request Body: JSON with terminal data

json
{
  "cable_id": "C1",
  "junction_box_id": "J1",
  "cable_name": "Cable 1",
  "terminal_data": [
    {
      "terminal_id": "C1-T1",
      "terminal_no": "1",
      "symbol": "ara/wago",
      "input_left": "L1",
      "input_right": "N",
      "spare": "N",
      "input_connected": "Y",
      "output_connected": "Y",
      "output_left": "",
      "output_right": ""
    }
  ]
}
Response: JSON with save status

## Finalize Terminal Configuration
http
POST /finalize_terminal_config
Finalize terminal configuration (simplified).

Request Body: JSON with cable ID

Response: JSON with success message

## Get Terminal Draft
http
GET /get_terminal_draft
Load existing terminal configuration from Terminal table.

Parameters: cable_id

Response: JSON with terminal data

## Terminal Header Management
## Save Terminal Header Draft
http
POST /save_terminal_header_draft
Save terminal header configuration - UPDATE IF EXISTS, INSERT IF NEW.

Request Body: JSON with header data

json
{
  "cable_id": "C1",
  "header_data": [
    {
      "header_type": "Header1",
      "terminal_start": "1",
      "terminal_end": "12",
      "input_output": "input",
      "text": "Input Header"
    }
  ]
}
Response: JSON with save status

## Finalize Terminal Header
http
POST /finalize_terminal_header
Convert draft headers to final (remove DRAFT- prefix).

Request Body: JSON with cable ID

Response: JSON with finalization status

## Manage Terminal Header Draft
http
GET /get_terminal_header_draft
POST /clear_terminal_header_draft
Load or clear terminal header draft.

Parameters: project_id, cable_id

Response: JSON with draft data or clear status

## Get Terminal Header Summary
http
GET /get_terminal_header_summary
Get current terminal header summary.

Response: JSON with total count

## Group Table Management
## Save Group Table Draft
http
POST /save_group_table_draft
Save group table configuration - UPDATE IF EXISTS, INSERT IF NEW.

Request Body: JSON with group data

json
{
  "cable_id": "C1",
  "group_data": [
    {
      "group_id": "GR001",
      "terminal_no": "1,2,3",
      "input_output": "input",
      "text": "Group Description"
    }
  ]
}
Response: JSON with save status

## Finalize Group Table
http
POST /finalize_group_table
Convert draft groups to final (remove DRAFT- prefix).

Request Body: JSON with cable ID

Response: JSON with finalization status

## Manage Group Table Draft
http
GET /get_group_table_draft
POST /clear_group_table_draft
Load or clear group table draft.

Parameters: project_id, cable_id

Response: JSON with draft data or clear status

## Get Group Summary
http
GET /get_group_summary
Get current group summary.

Response: JSON with total count

## Choke Table Management
## Save Choke Table Draft
http
POST /save_choke_table_draft
Save choke table configuration - UPDATE IF EXISTS, INSERT IF NEW.

Request Body: JSON with choke data

json
{
  "cable_id": "C1",
  "choke_data": [
    {
      "choke_id": "CH001",
      "input_terminal": "1",
      "output_terminal": "2",
      "terminal_name": "CHOKE",
      "output_type": "type1",
      "output_text": "Output Text",
      "output_connected": "Y"
    }
  ]
}
Response: JSON with save status

## Finalize Choke Table
http
POST /finalize_choke_table
Convert draft chokes to final (remove DRAFT- prefix).

Request Body: JSON with cable ID

Response: JSON with finalization status

## Manage Choke Table Draft
http
GET /get_choke_table_draft
POST /clear_choke_table_draft
Load or clear choke table draft.

Parameters: project_id, cable_id

Response: JSON with draft data or clear status

## Get Choke Summary
http
GET /get_choke_summary
Get current choke summary.

Response: JSON with total count

## Resistor Table Management
## Save Resistor Table Draft
http
POST /save_resistor_table_draft
Save resistor table configuration - UPDATE IF EXISTS, INSERT IF NEW.

Request Body: JSON with resistor data

json
{
  "cable_id": "C1",
  "resistor_data": [
    {
      "resistor_id": "R001",
      "input_terminal": "1",
      "output_terminal": "2",
      "resistor_name": "R"
    }
  ]
}
Response: JSON with save status

## Manage Resistor Table
http
GET /get_resistor_table_draft
POST /clear_resistor_table_draft
Load or clear resistor table data.

Parameters: project_id, cable_id

Response: JSON with resistor data or clear status

## Get Resistor Summary
http
GET /get_resistor_summary
Get current resistor summary.

Response: JSON with total count

## Comprehensive API Endpoints
## Get Project Data by Type
http
GET /api/project/<int:project_id>/cables
GET /api/project/<int:project_id>/terminals
GET /api/project/<int:project_id>/headers
GET /api/project/<int:project_id>/groups
GET /api/project/<int:project_id>/chokes
GET /api/project/<int:project_id>/resistors
GET /api/project/<int:project_id>/junctions
GET /api/project/<int:project_id>/station
Get all records of a specific type for a project.

Response: JSON array of records

## CRUD Operations for Project Data
http
POST /api/project/<int:project_id>/cables
PUT /api/project/<int:project_id>/cables/<int:record_id>
DELETE /api/project/<int:project_id>/cables/<int:record_id>

POST /api/project/<int:project_id>/terminals
PUT /api/project/<int:project_id>/terminals/<int:record_id>
DELETE /api/project/<int:project_id>/terminals/<int:record_id>

POST /api/project/<int:project_id>/headers
PUT /api/project/<int:project_id>/headers/<int:record_id>
DELETE /api/project/<int:project_id>/headers/<int:record_id>

POST /api/project/<int:project_id>/groups
PUT /api/project/<int:project_id>/groups/<int:record_id>
DELETE /api/project/<int:project_id>/groups/<int:record_id>

POST /api/project/<int:project_id>/chokes
PUT /api/project/<int:project_id>/chokes/<int:record_id>
DELETE /api/project/<int:project_id>/chokes/<int:record_id>

POST /api/project/<int:project_id>/resistors
PUT /api/project/<int:project_id>/resistors/<int:record_id>
DELETE /api/project/<int:project_id>/resistors/<int:record_id>

POST /api/project/<int:project_id>/junctions
PUT /api/project/<int:project_id>/junctions/<int:record_id>
DELETE /api/project/<int:project_id>/junctions/<int:record_id>

## CRUD operations for various data types.

Request Body: JSON with data fields

Response: JSON with operation status

## Update Station Data
http
POST /api/project/<int:project_id>/station
PUT /api/project/<int:project_id>/station
Create or update station data.

Request Body: JSON with station fields

Response: JSON with save status

## Debug Project Data
http
GET /api/debug/project/<int:project_id>
Debug endpoint showing all data counts for a project.

Response: JSON with counts for all data types

## Project and Stage Management
## Start New Drawing from Sidebar
http
POST /start_new_drawing_from_sidebar/<int:project_id>
Start a new drawing for selected project from sidebar (clears existing data).

Parameters: project_id - Project ID

Response: JSON with redirect URL

Success Response:

json
{
  "success": true,
  "message": "New drawing started for Project Name",
  "redirect_url": "/workflow/step/2"
}
## Continue Draft
http
GET /project/<int:project_id>/continue-draft
Continue working on a project draft from its current stage.

Response: Redirects to appropriate stage

## View Stage Data
http
GET /project/<int:project_id>/stage/<int:stage>/view
View and edit data for a specific stage (1-9).

Parameters:

project_id - Project ID

stage - Stage number (1-9)

Response: HTML page for stage data editing

## Get Stage API Data
http
GET /project/<int:project_id>/stage/<int:stage>/api-data
Get JSON data for a specific stage with optional filtering.

Query Parameters:

cable_id - Filter by cable ID (for stages 3-9)

junction_box - Filter by junction box ID (for stage 3)

junction_name - Filter by junction name (for stage 3)

Response: JSON with stage data and counts

## Stage-wise CRUD Operations
http
GET|POST|DELETE /project/<int:project_id>/stage/1/station-master/<item_id>
GET|POST|DELETE /project/<int:project_id>/stage/1/station-drawing/<item_id>
GET|POST|DELETE /project/<int:project_id>/stage/2/junction-box/<item_id>
GET|POST|DELETE /project/<int:project_id>/stage/3/cable/<item_id>
GET|POST|DELETE /project/<int:project_id>/stage/4/terminal/<item_id>
GET|POST|DELETE /project/<int:project_id>/stage/5/terminal-header/<item_id>
GET|POST|DELETE /project/<int:project_id>/stage/6/group/<item_id>
GET|POST|DELETE /project/<int:project_id>/stage/7/choke-table/<item_id>
GET|POST|DELETE /project/<int:project_id>/stage/8/resistor-table/<item_id>
GET|POST|DELETE /project/<int:project_id>/stage/9/cable-box/<item_id>
CRUD operations for each stage's data.

Parameters: item_id - Record ID or 'new' for creation

GET: Returns record data

POST: Creates/updates record

DELETE: Deletes record

## Helper Routes for Stage Data
http
GET /project/<int:project_id>/stage/2/junction-boxes/json
GET /project/<int:project_id>/stage/3/cables/json
GET /project/<int:project_id>/stage/3/cables-by-junction/<junction_id>/json
GET /project/<int:project_id>/stage/3/cable/<cable_id>/details/json
GET /project/<int:project_id>/all-stages-summary/json
Get dropdown data and summaries for stages.

Response: JSON with lists or summary data

## Comprehensive Project View
http
GET /project/<int:project_id>/comprehensive-view
Comprehensive view with step-by-step workflow.

Response: HTML page with all stages

## Cable Row Configuration
http
POST /project/<int:project_id>/save-cable-row-config
Save cable row configuration (update existing rows in place).

Request Body: JSON with rows

json
{
  "junction_box_id": "J1",
  "rows": [
    {
      "row_number": "1",
      "location_row_name": "A",
      "cable_type": "cable",
      "number_of_cables": 6
    }
  ]
}
Response: JSON with save status

## Get Cable Row Configuration
http
GET /project/<int:project_id>/cable-row-config
Get cable row configuration for a specific junction box.

Query Parameters: junction_box_id

Response: JSON with configuration rows

## Finalize Cable Configuration
http
POST /project/<int:project_id>/stage/3/cable-config/finalize
Finalize cable configuration (mark as non-draft).

Request Body: JSON with junction box ID

Response: JSON with finalization status

## Generate Cables from Configuration
http
POST /project/<int:project_id>/stage/3/generate-cables
Generate cables based on cable row configuration.

Request Body: JSON with junction details

json
{
  "junction_box_id": "J1",
  "junction_name": "Junction 1",
  "junction_size": "Full"
}
Response: JSON with generated cables

## View and List Management
## View List with Filtering
http
GET /view_list
View all projects with summary statistics and filtering.

Query Parameters:

page - Page number (default: 1)

rows_per_page - Rows per page (default: 20)

project_id - Filter by project ID

approval_status - Filter by status (all, approved, rejected, pending, drawing_in_progress, no_drawing)

Response: HTML page with filtered projects

## Check Designation Users
http
GET /check-designation-users/<int:id>
Check how many users have a specific designation.

Parameters: id - Designation ID

Response: JSON with user count

## Role Management (Admin Only)
## List Roles
http
GET /roles
List all roles with filtering and pagination.

Query Parameters:

page - Page number

per_page - Items per page

search - Search by role name

status - Filter by status (all, active, inactive)

Response: HTML page with roles list

## Create Role
http
GET|POST /roles/create
Create a new role.

POST Form Data:

role_name - Role name (required)

is_active - Active status checkbox

Response: Redirects to roles list

## Edit Role
http
GET|POST /roles/<int:id>/edit
Edit an existing role.

Parameters: id - Role ID

POST Form Data: Same as create

Response: Redirects to roles list

## Delete Role
http
GET /roles/<int:id>/delete
Delete a role (only if no users have it).

Parameters: id - Role ID

Response: Redirects to roles list

## Toggle Role Status
http
POST /roles/<int:id>/toggle-status
Toggle role active/inactive status.

Parameters: id - Role ID

Response: JSON with new status

## Designation Management (Admin Only)
## List Designations
http
GET /designations
List all designations with filtering and pagination.

Query Parameters: Same as roles

Response: HTML page with designations list

## Create Designation
http
GET|POST /designations/create
Create a new designation.

POST Form Data:

designation_name - Designation name (required)

approval_level - Approval level (optional number)

is_active - Active status checkbox

Response: Redirects to designations list

## Edit Designation
http
GET|POST /designations/<int:id>/edit
Edit an existing designation.

Parameters: id - Designation ID

POST Form Data: Same as create

Response: Redirects to designations list

## Delete Designation
http
GET /designations/<int:id>/delete
Delete a designation (only if no users have it).

Parameters: id - Designation ID

Response: Redirects to designations list

## Toggle Designation Status
http
POST /designations/<int:id>/toggle-status
Toggle designation active/inactive status.

Parameters: id - Designation ID

Response: JSON with new status

## Data Management
## Delete Project Data
http
POST /delete_projects/<int:project_id>
Delete all related data for a project but keep the project itself (Admin only).

Parameters: project_id - Project ID

Response: JSON with deletion status

## Get Checksum
http
GET /get_checksum/<int:pdf_id>
Get checksum for a fully approved PDF.

Parameters: pdf_id - PDF ID

Response: JSON with checksum and filename

Success Response:

json
{
  "success": true,
  "checksum": "a1b2c3d4e5f6...",
  "filename": "project_1_v1.pdf"
}
## Update Cable via AJAX
http
POST /update_cable_ajax/<int:cable_id>
Update cable information via AJAX.

Parameters: cable_id - Cable database ID

Request Body: JSON with cable fields to update

Response: JSON with update status

## Database Models Used
Additional Models:

CableRowConfig - Cable row configuration

CableSummary - Cable statistics summary

TerminalSummary - Terminal statistics summary

TerminalHeaderSummary - Terminal header statistics summary

GroupSummary - Group statistics summary

ChokeSummary - Choke statistics summary

ResistorSummary - Resistor statistics summary

JunctionBoxSummary - Junction box statistics summary

RoleMaster - Role management

DesignationMaster - Designation management


## API Version
1.0 (as of 2026-01-08)
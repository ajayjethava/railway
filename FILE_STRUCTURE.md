# File Structure - Circuit Building Application

## Overview
This document provides a detailed breakdown of all files and directories in the Circuit Building Application project.

## Root Directory

Circuitbuilding/
│
├── run.py # Main Flask application entry point
├── requirements.txt # Python dependencies
├── README.md # Project overview and setup guide
├── .env # Environment configuration
├── .gitattributes # Git line ending rules
│
├── app/ # Core Flask application
├── .git/ # Git repository data
│
└── FILE_STRUCTURE.md # This documentation file



## Application Structure

### Circuitbuilding/app/ - Core Application

app/
│
├── init.py # Flask app factory and configuration
├── models.py # SQLAlchemy database models (14 models)
├── routes.py # Main Flask routes and API endpoints
├── schemas.py # Excel sheet schema definitions
├── database.py # Database instance initialization
├── create_initial_admin.py # Initial admin user creation script
├── monitor.py # System monitoring functionality
│
├── routes/ # Modular route blueprints
│ └── new_project/
│ └── project_routes.py # Project-specific routes
│
├── static/ # Static assets
│ ├── images/
│ │ └── railway_logo.jpg # Application logo
│ └── styles.css # Main stylesheet
│
├── templates/ # Jinja2 HTML templates
│ ├── base.html # Base template with navigation
│ ├── index.html # Main dashboard
│ ├── login.html # Authentication page
│ ├── admin_users.html # User management interface
│ ├── approval_tracking.html # Approval workflow tracking
│ ├── new_project.html # Project creation form
│ ├── edit_project.html # Project editing interface
│ ├── excel_to_pdf.html # Excel to PDF conversion page
│ ├── pdf_result.html # PDF generation results
│ ├── pdf_view.html # PDF viewer interface
│ ├── preview.html # Data preview page
│ ├── project_selection.html # Project selection interface
│ ├── sheet_form.html # Excel sheet upload form
│ ├── upload_sheet.html # File upload interface
│ ├── view_list.html # Project listing view
│ ├── new_drawing_selection.html # Drawing mode selection
│ │
│ ├── designations/ # Designation management templates
│ │ ├── form.html # Designation create/edit form
│ │ └── list.html # Designations listing
│ │
│ ├── roles/ # Role management templates
│ │ ├── form.html # Role create/edit form
│ │ └── list.html # Roles listing
│ │
│ └── workflow/ # Workflow management templates
│ ├── comprehensive_view.html # Complete workflow overview
│ ├── workflow_dashboard.html # Workflow dashboard
│ ├── workflow_edit.html # Workflow editing interface
│ ├── workflow_step.html # Individual workflow step
│ ├── stage_data_view.html # Stage-specific data view
│ │
│ ├── step_1_station_info.html # Station information step
│ ├── step_2_locations/ # Location configuration
│ │ ├── main.html # Main locations interface
│ │ ├── add_more_grid.html # Additional grid configuration
│ │ ├── modals/ # Configuration modals
│ │ │ ├── cable_config.html # Cable configuration
│ │ │ ├── cable_table.html # Cable table
│ │ │ ├── choke_config.html # Choke configuration
│ │ │ ├── group_config.html # Group configuration
│ │ │ ├── header_config.html # Header configuration
│ │ │ ├── resistor_config.html # Resistor configuration
│ │ │ └── terminal_config.html # Terminal configuration
│ │ └── scripts/ # JavaScript files
│ │ ├── add_more_flow.js # Add grid functionality
│ │ └── main_flow.js # Main locations logic
│ │
│ ├── step_3_cables.html # Cable management
│ ├── step_4_terminals.html # Terminal configuration
│ ├── step_5_headers.html # Header management
│ ├── step_6_groups.html # Group configuration
│ ├── step_7_choke_table.html # Choke table management
│ ├── step_8_resistor_table.html # Resistor table management
│ ├── step_9_cable_box.html # Cable box configuration
│ │
│ └── includes/ # Reusable template components
│ ├── approval_buttons.html # Approval action buttons
│ ├── approval_status_panel.html # Approval status display
│ ├── comprehensive_modals.html # Comprehensive modal dialogs
│ ├── existing_records_table.html # Records table template
│ ├── progress_header.html # Progress indicator
│ ├── rejection_modal.html # Rejection confirmation modal
│ ├── success_error_modals.html # Success/error message modals
│ └── view_project_data_modal.html # Project data viewer
│
└── pycache/ # Python bytecode cache (multiple versions)



## Key Files Explained

### Root Level Files

#### run.py
- **Purpose**: Application entry point
- **Content**: Starts Flask development server on port 5000 with debugging enabled

#### requirements.txt
- **Purpose**: Python package dependencies
- **Key Dependencies**: Flask, SQLAlchemy, pandas, openpyxl, matplotlib

#### .env
- **Purpose**: Environment configuration
- **Contains**: Database connection strings, secret keys, FTP settings

### Application Core Files

#### app/__init__.py
- **Purpose**: Flask application factory
- **Key Functions**:
  - Application instance creation
  - Database initialization
  - Blueprint registration
  - Session configuration
  - Authentication setup

#### app/models.py
- **Purpose**: SQLAlchemy ORM models
- **Models Included**:
  - User (authentication and permissions)
  - Project (project management)
  - StationDrawing (station information)
  - GeneratedPDF (PDF document management)
  - JunctionBox (junction box definitions)
  - Cable (cable configurations)
  - CableBox (cable/relay boxes)
  - Terminal (terminal connections)
  - Group (terminal groupings)
  - TerminalHeader (connection headers)
  - ChokeTable (choke components)
  - ResistorTable (resistor components)
  - Notification (user notifications)
  - StationMaster (station access control)

#### app/routes.py
- **Purpose**: Main application routes and API endpoints
- **Key Route Categories**:
  - Authentication (`/login`, `/logout`)
  - User management (`/admin/users/*`)
  - Project CRUD operations
  - Excel upload and PDF generation
  - Approval workflow management
  - Notification system
  - Data management (cables, terminals, junction boxes)
  - Workflow step navigation
  - Admin utilities

#### app/schemas.py
- **Purpose**: Excel sheet schema definitions
- **Schemas Defined**:
  - StationDrawing (station metadata)
  - junction_box (junction configuration)
  - cable (cable specifications)
  - cable_box (relay box configuration)
  - terminal (terminal connections)
  - group (terminal groupings)
  - terminal_header (header specifications)
  - choketable (choke components)
  - resistortable (resistor components)

#### app/database.py
- **Purpose**: Database engine configuration
- **Content**: SQLAlchemy instance creation

#### app/create_initial_admin.py
- **Purpose**: Creates initial admin user
- **Usage**: Run once during setup to create default admin account

### Template Categories

#### Authentication Templates
- `login.html` - User authentication form

#### Dashboard Templates
- `index.html` - Main application dashboard
- `workflow_dashboard.html` - Workflow management interface

#### Project Management Templates
- `new_project.html`, `edit_project.html` - Project creation/editing
- `project_selection.html` - Project selection interface
- `view_list.html` - Project listing with filters

#### Workflow Templates
- `workflow_step.html` - Individual workflow step interface
- `comprehensive_view.html` - Complete workflow overview
- `stage_data_view.html` - Stage-specific data management

#### Configuration Templates
- `excel_to_pdf.html` - Excel to PDF conversion
- `sheet_form.html`, `upload_sheet.html` - Data upload interfaces
- `pdf_view.html`, `pdf_result.html` - PDF viewing and results

#### Administration Templates
- `admin_users.html` - User management
- `approval_tracking.html` - Approval workflow tracking
- `designations/` - Designation management
- `roles/` - Role management

### Static Assets

#### app/static/styles.css
- **Purpose**: Main application stylesheet
- **Features**: Responsive design, custom components, PDF viewer styling

#### app/static/images/
- **Purpose**: Application images and logos

## Database Structure

The application uses SQLAlchemy ORM with the following key relationships:

1. **User ↔ Project** (Many-to-Many through user projects)
2. **Project → StationDrawing** (One-to-Many)
3. **Project → GeneratedPDF** (One-to-Many)
4. **StationDrawing → JunctionBox** (One-to-Many)
5. **JunctionBox → Cable** (One-to-Many)
6. **Cable → Terminal** (One-to-Many)
7. **Cable → Group, TerminalHeader, ChokeTable, ResistorTable** (One-to-Many)

## Workflow Stages

The application follows a 9-step workflow:

1. **Station Information** - Basic project and station details
2. **Locations** - Junction box and location configuration
3. **Cables** - Cable definitions and routing
4. **Terminals** - Terminal connection specifications
5. **Headers** - Terminal header configurations
6. **Groups** - Terminal grouping logic
7. **Choke Table** - Choke component specifications
8. **Resistor Table** - Resistor component specifications
9. **Cable Box** - Relay/cable box finalization

## API Structure

The application provides RESTful API endpoints for:

- **Authentication**: `/login`, `/logout`
- **User Management**: `/admin/users/*`
- **Project Operations**: `/project/*`
- **Workflow Steps**: `/workflow/step/*`
- **Data Management**: `/api/project/*`
- **PDF Operations**: `/pdf/*`, `/download_pdf/*`
- **Approval Workflow**: `/approve_pdf/*`, `/reject_pdf/*`
- **Notifications**: `/notifications/*`

## Development Notes

### Python Version Compatibility
The application maintains Python bytecode for multiple versions (3.10-3.13) in `__pycache__` directories.

### Database Configuration
- **Primary Database**: PostgreSQL (configured in `.env`)
- **Session Storage**: PostgreSQL session table
- **ORM**: SQLAlchemy with Flask-SQLAlchemy extension

### Security Considerations
- Password hashing with Werkzeug
- Session-based authentication
- Role-based access control (5 levels)
- CSRF protection
- SQL injection prevention via ORM

### File Upload Handling
- Excel file upload for data import
- PDF generation from Excel data
- File validation and sanitization
- Temporary file cleanup

## Deployment Structure

For production deployment:

1. **WSGI Server**: Gunicorn or uWSGI
2. **Reverse Proxy**: Nginx or Apache
3. **Database**: PostgreSQL with connection pooling
4. **Session Storage**: PostgreSQL or Redis
5. **Static Files**: Served by Nginx/CDN
6. **Background Tasks**: Celery for FTP automation

## Maintenance Commands

### Initial Setup
```bash
python app/create_initial_admin.py


### Database Migrations

flask db init
flask db migrate
flask db upgrade


### Running the Application
#windows

python -m venv venv
venv/Scripts/activate

python run.py  # Development
gunicorn --bind 0.0.0.0:5000 wsgi:app  # Production


**Last Updated**: 2026-01-06
Project Version: 1.0
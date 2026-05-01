# DrawingMaster - Railway Terminal Drawing Management System

A comprehensive Flask-based web application for managing Railway Terminal Drawings with automated PDF generation, multi-level approval workflows, and FTP/SFTP integration.

## 🚀 Project Overview

DrawingMaster is a specialized system designed for railway engineering teams to create, manage, and approve terminal circuit diagrams. It converts Excel-based circuit designs into professional PDF drawings with integrated metadata, checksums, and approval tracking.

### Key Features

- **📊 Excel to PDF Conversion**: Automated conversion of technical Excel sheets into professional circuit PDFs
- **✅ Multi-Level Approval System**: Three-tier approval workflow (Level 1, 2, and 3)
- **🔐 Role-Based Access Control**: 5 user roles with different permissions
- **📡 FTP/SFTP Integration**: Automated file monitoring and processing
- **🔔 Real-time Notifications**: User notifications for drawing status changes
- **🏢 Project Management**: Organize drawings by projects and stations
- **🔒 Session Management**: PostgreSQL-backed secure sessions with auto-timeout
- **📈 Version Control**: Track drawing versions with MD5 checksums

## 📁 Project Structure

```
DrawingMaster/
├── Circuitbuilding/              # Main Flask application package
│   ├── app/                      # Application core
│   │   ├── __init__.py           # Flask app factory
│   │   ├── models.py             # SQLAlchemy database models
│   │   ├── routes.py             # Flask routes & API endpoints (523KB)
│   │   ├── schemas.py            # Excel sheet schemas
│   │   ├── database.py           # Database configuration
│   │   ├── create_initial_admin.py  # Admin user creation script
│   │   ├── templates/            # Jinja2 HTML templates (48 files)
│   │   └── static/               # CSS, JS, images
│   ├── requirements.txt          # Python dependencies
│   ├── run.py                    # Development server entry point
│   ├── .env                      # Environment configuration
│   └── README.md                 # Setup instructions
├── run.py                        # Main entry point
├── excel_to_pdf_converter.py    # PDF generation engine (3915 lines)
├── ftp_auto_converter.py         # FTP automation daemon (1492 lines)
├── test1.py                      # Enhanced FTP converter with notifications
├── 1.py                          # Directory tree generator utility
├── requirements.txt              # Root dependencies
├── uploads/                      # Generated PDFs and Excel files
├── downloaded_xlsx_files/        # FTP downloaded files
├── temp_xlsx/                    # Temporary Excel files
├── temp_pdf/                     # Temporary PDF files
├── flask_session/                # Session storage
└── instance/                     # Instance-specific configs
```

## 🛠️ Technology Stack

### Backend
- **Framework**: Flask 3.x
- **Database**: PostgreSQL
- **ORM**: SQLAlchemy
- **Authentication**: Flask-Login
- **Session**: Flask-Session (PostgreSQL-backed)

### Frontend
- **Templates**: Jinja2
- **Styling**: CSS3, Bootstrap
- **JavaScript**: Vanilla JS

### PDF Generation
- **Plotting**: Matplotlib
- **Data Processing**: Pandas, NumPy
- **Excel Parsing**: openpyxl
- **PDF Manipulation**: PyPDF2

### File Transfer
- **Protocol**: FTP/SFTP
- **Library**: Paramiko (SFTP), ftplib (FTP)

## 🗄️ Database Models

### Core Models

1. **User** - User accounts with role-based permissions
   - Roles: Viewer (0), Creator (1), Approver L2 (2), Approver L3 (3), Admin (4)
   
2. **Project** - Railway projects containing multiple drawings
   - Status tracking, junction data, stage management
   
3. **StationDrawing** - Station metadata and drawing configuration
   - Station code, zone, division, version info
   
4. **GeneratedPDF** - Generated PDF metadata and checksums
   - MD5 checksums, approval status, version tracking
   
5. **JunctionBox** - Junction box definitions
   
6. **Cable** - Cable configurations
   
7. **CableBox** - Cable box (relay box) definitions
   
8. **Terminal** - Terminal connections and symbols
   
9. **Group** - Terminal groupings
   
10. **TerminalHeader** - Terminal header configurations
   
11. **ChokeTable** - Choke component definitions
   
12. **ResistorTable** - Resistor component definitions
   
13. **Notification** - User notifications for drawing approvals

14. **StationMaster** - Station-specific user access control

## 🚦 Approval Workflow

```
Creator (Role 1)
    ↓ Creates Drawing
Level 1 Approval (Creator/Designation: level1)
    ↓ Approves/Rejects
Level 2 Approval (Role 2/Designation: level2)
    ↓ Approves/Rejects
Level 3 Approval (Role 3/Designation: level3)
    ↓ Approves
✅ Fully Approved
```

## 📋 Excel Sheet Schemas

The system processes Excel files with the following sheets:

1. **StationDrawing** - Station metadata (checksum, IDs, zones, etc.)
2. **junction_box** - Junction box locations and properties
3. **cable** - Cable definitions with positions and terminals
4. **cable_box** - Relay boxes (similar to cables)
5. **terminal** - Terminal symbols and connections
6. **group** - Terminal groupings by cable
7. **terminal_header** - Connection headers (WIREFROM/WIRETO/RELAY)
8. **choketable** - Choke components for filtering
9. **resistortable** - Resistor component values

## ⚙️ Installation & Setup

### Prerequisites

- Python 3.8+
- PostgreSQL 12+
- Virtual environment (recommended)

### 1. Clone Repository

```bash
git clone <repository-url>
cd DrawingMaster
```

### 2. Create Virtual Environment

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**Linux/Mac:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies

```bash
# Install from root requirements
pip install -r requirements.txt

# Install from Circuitbuilding requirements
cd Circuitbuilding
pip install -r requirements.txt
```

### 4. Configure Database

Edit `Circuitbuilding/app/__init__.py` line 52:
```python
app.config["SQLALCHEMY_DATABASE_URI"] = "postgresql://user:password@host:port/database"
```

### 5. Configure Environment (.env)

Create `Circuitbuilding/.env`:
```env
SECRET_KEY=your-secret-key-here

# FTP Configuration
FTP_ENABLED=True
FTP_HOST=your-ftp-server.com
FTP_PORT=22
FTP_USERNAME=your-username
FTP_PASSWORD=your-password
FTP_UPLOAD_DIR=/srv/railway/frontend/uploads/
FTP_XLSX_TAKE_DIR=/srv/railway/frontend/xlsx_download/
FTP_USE_SFTP=True
FTP_TIMEOUT=30

# Logging
LOG_LEVEL=INFO
LOG_FILE=ftp_converter.log
```

### 6. Initialize Database

```bash
python run.py
# Database tables will be created automatically on first run
```

### 7. Create Admin User

```bash
cd Circuitbuilding/app
python create_initial_admin.py
```

### 8. Run Application

**Development Server:**
```bash
python run.py
# Access at http://localhost:5000
```

**Production:**
```bash
gunicorn -w 4 -b 0.0.0.0:5000 run:app
```

## 🔄 FTP Automation

### Start FTP Auto-Converter

```bash
python ftp_auto_converter.py
```

This daemon:
1. Monitors FTP/SFTP directory for new Excel files
2. Downloads files matching pattern `RAILWAYPROJECT_ID{n}_*.xlsx`
3. Converts to PDF using `excel_to_pdf_converter.py`
4. Uploads both XLSX and PDF to remote `uploads/` directory
5. Moves processed files to `processed/` directory
6. Updates database with PDF metadata
7. Creates notifications for approvers

### Features

- **Loop Prevention**: Tracks processed files to avoid reprocessing
- **Error Handling**: Moves failed files to `failed/` directory
- **Checksum Generation**: MD5 checksums for file integrity
- **Metadata Extraction**: Automatically extracts station data
- **Version Management**: Auto-increments version numbers
- **Project Status Updates**: Updates project stages based on PDF generation

## 📖 Usage Guide

### Creating a Drawing

1. Login with Creator/Admin account
2. Select or create a Project
3. Upload Excel file with required sheets
4. System processes file and generates PDF

### Approval Process

**Level 1 Approver:**
1. Receives notification for new drawing
2. Reviews PDF
3. Approves → 5. Notifications sent to Level 1 approvers
sends to Level 2
4. Rejects → sends back to creator

**Level 2 & 3:** Similar process

### Admin Functions

- User management (create, edit, delete users)
- Project management
- View all drawings across projects
- Override approvals if needed
- System configuration

## 🔧 Key Scripts

### excel_to_pdf_converter.py
- 3915 lines of PDF generation logic
- Matplotlib-based circuit diagram renderer
- Features:
  - Symbol drawing (fuses, relays, terminals, etc.)
  - Bus connections (input/output)
  - Cable boxes and junction boxes
  - Terminal numbering and grouping
  - Checksum generation and metadata embedding
  - Row pagination (max 36 terminals/row, 6 cable boxes/row)

### ftp_auto_converter.py
- 1492 lines of FTP automation
- Background daemon for continuous monitoring
- Features:
  - SFTP/FTP connection handling
  - File pattern matching
  - Duplicate detection
  - Database updates
  - Notification creation
  - Path discovery and debugging

### test1.py
- Enhanced version with station access controls
- User-project permission checking
- Improved notification routing

## 🔒 Security Features

- **Password Hashing**: Werkzeug SHA-256
- **Session Management**: PostgreSQL-backed, 30-minute timeout
- **Role Validation**: SQLAlchemy validators
- **CSRF Protection**: Flask-WTF tokens
- **SQL Injection Prevention**: SQLAlchemy ORM
- **File Upload Validation**: Extension and size checks

## 📊 API Endpoints

*(Partial list - see routes.py for complete API)*

### Authentication
- `POST /login` - User login
- `GET /logout` - User logout

### Projects
- `GET /projects` - List all projects
- `POST /projects/create` - Create new project
- `GET /project/<id>` - View project details
- `POST /project/<id>/update` - Update project

### Drawings
- `POST /upload_excel` - Upload Excel file
- `GET /pdf/<filename>` - View PDF
- `POST /approve/<pdf_id>` - Approve drawing
- `POST /reject/<pdf_id>` - Reject drawing

### Notifications
- `GET /notifications` - Get user notifications
- `POST /notifications/<id>/read` - Mark as read

### Admin
- `GET /admin/users` - List users
- `POST /admin/users/create` - Create user
- `POST /admin/users/<id>/edit` - Edit user

## 🐛 Troubleshooting

### Database Connection Issues
```python
# Check connection string in Circuitbuilding/app/__init__.py
# Verify PostgreSQL is running
# Check credentials and port
```

### FTP Connection Fails
```bash
# Test connection manually
# Check firewall rules
# Verify credentials in .env
# Run path discovery: see ftp_auto_converter.py::discover_remote_paths()
```

### PDF Generation Errors
```bash
# Check Excel sheet structure matches schemas.py
# Verify Matplotlib is properly installed
# Check temp directories exist and are writable
```

### Session Timeout Issues
```python
# Check app.config['PERMANENT_SESSION_LIFETIME'] in __init__.py
# Verify PostgreSQL sessions table exists
# Clear browser cookies
```

## 📝 Development

### Adding New Sheet Type

1. Add schema to `schemas.py`:
```python
SHEETS["new_sheet"] = ["column1", "column2", ...]
```

2. Create model in `models.py`:
```python
class NewSheet(db.Model):
    __tablename__ = 'new_sheet'
    # ... fields
```

3. Add processing logic in `routes.py`

4. Update `excel_to_pdf_converter.py` for rendering

### Adding New User Role

1. Update `User.validate_role()` in `models.py`
2. Add permissions in `get_user_permissions()` helper
3. Update role_map in templates
4. Add route decorators for access control

## 📄 License

*Add your license information here*

## 👥 Contributors

*Add contributor information here*

## 📞 Support

*Add support contact information here*

## 🔄 Version History

- **v1.0** - Initial release with core features
- **Current** - Enhanced FTP automation, improved notifications, station access control

---

**Built for Railway Engineering Teams** 🚂

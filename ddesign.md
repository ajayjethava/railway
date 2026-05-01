## Detailed ER Diagram: Railway Projects Management System
## DIAGRAM LEGEND
PK = Primary Key

FK = Foreign Key

UK = Unique Key

(1:N) = One-to-Many relationship

(1:1) = One-to-One relationship

(M:N) = Many-to-Many via join table

⭤ = Bidirectional relationship

◄── = References/Points to

## COMPLETE ER DIAGRAM
## CORE MASTER TABLE


┌────────────────────────────────────────────────────────────────────────────┐
│                         railway_projects (MASTER TABLE)                    │
├─────────┬───────────────────────────────────────────────────────────┬──────┤
│ Column  │ Type & Constraints                                        │ Ref  │
├─────────┼───────────────────────────────────────────────────────────┼──────┤
│ id      │ SERIAL PK                                                 │      │
│ name    │ VARCHAR(200) NOT NULL UK                                  │      │
│ status  │ VARCHAR(50) DEFAULT 'drawing_in_progress'                 │      │
│ stage   │ INTEGER                                                   │      │
│ station_id │ VARCHAR(100)                                           │      │
└─────────┴───────────────────────────────────────────────────────────┴──────┘
                                 │
         ┌───────────────────────┴──────────────────────────────────────────────────────────────────┐
         │                                                                                          │
    (1:N)▼ (1:N)▼    (1:N)▼       (1:N)▼       (1:N)▼    (1:N)▼     (1:N)▼    (1:N)▼      (1:N)▼   │
┌──────────────────────────────────────────────────────────────────────────────────────────────────┐│
│                                                                                                  ││
│    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐ ││
│    │ station_    │    │ junction_   │    │    cable    │    │   terminal  │    │   terminal  │ ││
│    │ drawing     │    │ box         │    │             │    │             │    │   header    │ ││
│    │             │    │             │    │             │    │             │    │             │ ││
│    ├─────────────┤    ├─────────────┤    ├─────────────┤    ├─────────────┤    ├─────────────┤ ││
│    │ id: PK      │    │ id: PK      │    │ id: PK      │    │ id: PK      │    │ id: PK      │ ││
│    │ project_id: │    │ project_id: │    │ project_id: │    │ project_id: │    │ project_id: │ ││
│    │ FK◄railway_ │    │ FK◄railway_ │    │ FK◄railway_ │    │ FK◄railway_ │    │ FK◄railway_ │ ││
│    │ station_id: │    │ junction_id:│    │ cable_id:   │    │ cable_id:   │    │ header_type:│ ││
│    │ VARCHAR(100)│    │ VARCHAR(100)│    │ VARCHAR(100)│    │ VARCHAR(100)│    │ VARCHAR(100)│ ││
│    │ UNIQUE(proj,│    │             │    │ UNIQUE(proj,│    │ UNIQUE(proj,│    │ UNIQUE(proj,│ ││
│    │ station_id) │    │             │    │ cable_id)   │    │ cable_id,   │    │ cable_id,   │ ││
│    │             │    │             │    │             │    │ terminal_id)│    │ header_type)│ ││
│    └─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘ ││
│          │                   │                  │                  │                  │           ││
│          │                   │                  │                  │                  │           ││
│    (1:N)▼              (1:N)▼             (1:N)▼             (1:N)▼             (1:N)▼           ││
│┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐      ││
││  group_     │    │  choke_     │    │  resistor_  │    │  cable_     │    │ generated_  │      ││
││  table      │    │  table      │    │  table      │    │  box        │    │ pdf         │      ││
││             │    │             │    │             │    │             │    │             │      ││
│├─────────────┤    ├─────────────┤    ├─────────────┤    ├─────────────┤    ├─────────────┤      ││
││ id: PK      │    │ id: PK      │    │ id: PK      │    │ id: PK      │    │ id: PK      │      ││
││ project_id: │    │ project_id: │    │ project_id: │    │ project_id: │    │ project_id: │      ││
││ FK◄railway_ │    │ FK◄railway_ │    │ FK◄railway_ │    │ FK◄railway_ │    │ FK◄railway_ │      ││
││ cable_id:   │    │ cable_id:   │    │ cable_id:   │    │ cable_id:   │    │ pdf_filename│      ││
││ VARCHAR(100)│    │ VARCHAR(100)│    │ VARCHAR(100)│    │ VARCHAR(100)│    │ VARCHAR(255)│      ││
││ group_id:   │    │ choke_id:   │    │ resistor_id:│    │ UNIQUE(proj,│    │ level1_status│     ││
││ VARCHAR(100)│    │ VARCHAR(100)│    │ VARCHAR(100)│    │ cable_id)   │    │ level2_status│     ││
│└─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘    │ level3_status│     ││
│                                                                             └───────┬───────┘     ││
│                                                                                     │ (1:N)       ││
│                                                                                (1:N)▼       (1:N)▼││
│                                                                          ┌─────────────┐  ┌───────┴──────┐
│                                                                          │  approval   │  │ notifications│
│                                                                          │             │  │              │
│                                                                          ├─────────────┤  ├──────────────┤
│                                                                          │ id: PK      │  │ id: PK       │
│                                                                          │ pdf_id: FK◄─┼──┼─pdf_id       │
│                                                                          │◄generated_pdf│  │ user_id: FK◄─┼─┐
│                                                                          │ level: INT  │  │ project_id:FK│ │
│                                                                          │ status:     │  │◄railway_proj │ │
│                                                                          │ VARCHAR(20) │  │ level:       │ │
│                                                                          │ approver_id:│  │ VARCHAR(20)  │ │
│                                                                          │ FK◄users    │  └──────────────┘ │
│                                                                          └─────────────┘                   │
│                                                                                                            │
│                                                                    ┌──────────────────────────────────────┘
│                                                                    │
│                                                             (1:1)▼ ▼ (1:1)▼ (1:1)▼ (1:1)▼ (1:1)▼ (1:1)▼ (1:1)▼
│                                        ┌─────────────────────────────────────────────────────────────────────────┐
│                                        │                    SUMMARY TABLES (1:1)                                 │
│                                        ├─────────────────────────────────────────────────────────────────────────┤
│                                        │ junction_box_summary ◄── railway_projects                               │
│                                        │ cable_summary ◄───────── railway_projects                               │
│                                        │ terminal_summary ◄────── railway_projects                               │
│                                        │ group_summary ◄───────── railway_projects                               │
│                                        │ terminal_header_summary ◄ railway_projects                              │
│                                        │ choke_summary ◄───────── railway_projects                               │
│                                        │ resistor_summary ◄────── railway_projects                               │
│                                        └─────────────────────────────────────────────────────────────────────────┘
│                                                                                                            │
│                                         (1:N)▼                                                             │
│                                 ┌─────────────────────────────────────────────────────────────────────┐     │
│                                 │           CONFIGURATION TABLES                                       │     │
│                                 ├─────────────────────────────────────────────────────────────────────┤     │
│                                 │ cable_row_config ◄── railway_projects                               │     │
│                                 │ cable_location_addition ◄── railway_projects                        │     │
│                                 │ draft ◄── users (user_id)                                           │     │
│                                 └─────────────────────────────────────────────────────────────────────┘     │
│                                                                                                            │
│                                                (M:N) via user_projects                                   │
│                                                  ╱                  ╲                                     │
│                                                ╱                    ╲                                    │
│                                    ┌─────────────┐            ┌─────────────┐                            │
│                                    │    users    │◄───────────┤ user_projects├───► railway_projects     │
│                                    │             │            │             │                            │
│                                    ├─────────────┤            ├─────────────┤                            │
│                                    │ id: PK      │            │ user_id: PK │                            │
│                                    │ username: UK│            │ project_id: │                            │
│                                    │ email: UK   │            │ PK          │                            │
│                                    │ role:       │            │ FK◄users    │                            │
│                                    │ VARCHAR(20) │            │ FK◄railway_ │                            │
│                                    │ role_id: FK │            │ projects    │                            │
│                                    │◄role_master │            └─────────────┘                            │
│                                    │ designation │                                                        │
│                                    │_id: FK◄─────┼──────┐                                                │
│                                    │ designation │      │                                                │
│                                    │_master      │      │                                                │
│                                    └─────────────┘      │                                                │
│                                         │               │                                                │
│                                   (1:N)▼         (1:N)▼ │                                                │
│                           ┌─────────────┐  ┌────────────┴──────┐                                         │
│                           │ role_master │  │ designation_master │                                         │
│                           │             │  │                   │                                         │
│                           ├─────────────┤  ├───────────────────┤                                         │
│                           │ id: PK      │  │ id: PK            │                                         │
│                           │ role_name:UK│  │ designation_name: │                                         │
│                           │ is_active:  │  │ UK                │                                         │
│                           │ BOOLEAN     │  │ approval_level:   │                                         │
│                           └─────────────┘  │ INTEGER           │                                         │
│                                            │ is_active:        │                                         │
│                                            │ BOOLEAN           │                                         │
│                                            └───────────────────┘                                         │
│                                                                                                          │
│                                        ┌──────────────────────────────────────────────────────────────┐   │
│                                        │            LEGACY/UNUSED TABLES                              │   │
│                                        ├──────────────────────────────────────────────────────────────┤   │
│                                        │ station_master ◄── railway_projects                          │   │
│                                        │ deleted_projects_backup                                      │   │
│                                        │ flask_sessions                                               │   │
│                                        │ user_sessions                                                │   │
│                                        │ sessions                                                     │   │
│                                        └──────────────────────────────────────────────────────────────┘   │
│                                                                                                          │
└──────────────────────────────────────────────────────────────────────────────────────────────────────────┘


## RELATIONSHIP MATRIX
## railway_projects Relationships:

railway_projects (id) ════════════════════════════════════════════════════════════╗
├──► station_drawing (project_id)                       [1:N] - 239 records       ║
├──► junction_box (project_id)                          [1:N] - 419 records       ║
├──► cable (project_id)                                 [1:N] - 3,602 records     ║
├──► terminal (project_id)                              [1:N] - 21,847 records    ║
├──► terminal_header (project_id)                       [1:N] - 11,917 records    ║
├──► group_table (project_id)                           [1:N] - 531 records       ║
├──► choke_table (project_id)                           [1:N] - 576 records       ║
├──► resistor_table (project_id)                        [1:N] - 220 records       ║
├──► cable_box (project_id)                             [1:N] - 1,255 records     ║
├──► generated_pdf (project_id)                         [1:N] - 80 records        ║
│   └──► approval (generated_pdf_id)                    [1:N] - 11 records        ║
│   └──► notifications (pdf_id)                         [1:N] - 174 records       ║
├──◄► users (via user_projects)                         [M:N] - Association table ║
├──► cable_row_config (project_id)                      [1:N] - 225 records       ║
├──► cable_location_addition (project_id)               [1:N] - 0 records         ║
├──► station_master (project_id)                        [1:N] - 0 records (unused)║
├──◄ junction_box_summary (project_id)                  [1:1] - 43 records        ║
├──◄ cable_summary (project_id)                         [1:1] - 10 records        ║
├──◄ terminal_summary (project_id)                      [1:1] - 12 records        ║
├──◄ group_summary (project_id)                         [1:1] - 9 records         ║
├──◄ terminal_header_summary (project_id)               [1:1] - 10 records        ║
├──◄ choke_summary (project_id)                         [1:1] - 8 records         ║
└──◄ resistor_summary (project_id)                      [1:1] - 0 records         ║


## users Relationships:

users (id) ═══════════════════════════════════════════════════════════════════════╗
├──► approval (approver_id)                         [1:N] - 11 records            ║
├──► notifications (user_id)                        [1:N] - 174 records           ║
├──► draft (user_id)                                [1:N] - 3 records             ║
├──◄► railway_projects (via user_projects)          [M:N] - Association table     ║
├──► role_master (role_id)                          [N:1] - 4 master records      ║
└──► designation_master (designation_id)            [N:1] - 11 master records     ║


## DATA FLOW AND WORKFLOW
## Stage Progression Sequence:

┌──────────────┐    ┌──────────────┐    ┌──────────────┐    ┌──────────────┐
│  STAGE 0     │───►│  STAGE 1     │───►│  STAGE 2     │───►│  STAGE 3     │
│  drawing_    │    │  Station     │    │  Junction    │    │  Cable       │
│  not_started │    │  Drawing     │    │  Box         │    │  Data        │
└──────────────┘    └──────────────┘    └──────────────┘    └──────────────┘
                                                              │
                ┌──────────────┐    ┌──────────────┐    ┌────┴────┐
                │  STAGE 7     │◄───│  STAGE 6     │◄───│ STAGE 5 │◄───┐
                │  Choke       │    │  Group       │    │ Terminal│    │
                │  Table       │    │  Table       │    │ Header  │    │
                └──────────────┘    └──────────────┘    └─────────┘    │
                      │                                               │
                ┌─────┴──────┐    ┌──────────────┐    ┌──────────────┐│
                │  STAGE 8   │    │  STAGE 9     │    │  STAGE 10    ││
                │  Resistor  │───►│  Cable Box   │───►│  Ready for   ││
                │  Table     │    │              │    │  PDF Gen     ││
                └────────────┘    └──────────────┘    └───────┬──────┘│
                                                              │       │
                ┌──────────────┐    ┌──────────────┐    ┌─────┴──────┐│
                │  STAGE 14    │◄───│  STAGE 13    │◄───│  STAGE 12  │◄──┐
                │  Drawing     │    │  Level 2     │    │  Level 1   │   │
                │  Approved    │    │  Approved    │    │  Approved  │   │
                └──────────────┘    └──────────────┘    └────────────┘   │
                                                                         │
                                                                   ┌─────┴────┐
                                                                   │ STAGE 11  │
                                                                   │ Waiting   │
                                                                   │ for L1    │
                                                                   └──────────┘


## Approval Workflow:

┌─────────────────────────────────────────────────────────────────────────────┐
│                        PDF APPROVAL WORKFLOW                                │
├─────────────────────┬───────────────────────────────────────────────────────┤
│  Level 1 Approver   │  Level 2 Approver   │  Level 3 Approver               │
│  (designation=1)    │  (designation=2)    │  (designation=3)                │
├─────────────────────┼───────────────────────────────────────────────────────┤
│ 1. PDF generated    │ 4. Level 1 approved │ 7. Level 2 approved             │
│    → status:        │    → status:        │    → status:                    │
│    level1=pending   │    level2=pending   │    level3=pending               │
│                     │                     │                                 │
│ 2. Notification     │ 5. Notification     │ 8. Notification                 │
│    sent to L1       │    sent to L2       │    sent to L3                   │
│                     │                     │                                 │
│ 3. L1 approves/     │ 6. L2 approves/     │ 9. L3 approves/                 │
│    rejects          │    rejects          │    rejects                      │
│    → updates        │    → updates        │    → updates                    │
│    level1_status    │    level2_status    │    level3_status                │
│    & creates        │    & creates        │    & creates                    │
│    approval record  │    approval record  │    approval record              │
└─────────────────────┴───────────────────────────────────────────────────────┘

## TABLE DEPENDENCY HIERARCHY
## Level 1: Core Masters

railway_projects ─┐
users ────────────┼─┐
role_master ──────┘ │
designation_master ─┘

## Level 2: Project Data

railway_projects
  ├── station_drawing
  ├── junction_box
  ├── cable
  ├── terminal
  ├── terminal_header
  ├── group_table
  ├── choke_table
  ├── resistor_table
  └── cable_box

## Level 3: PDF & Approval

railway_projects
  └── generated_pdf
        ├── approval (also depends on users)
        └── notifications (also depends on users)

## Level 4: Configuration

railway_projects
  ├── cable_row_config
  └── cable_location_addition

users
  └── draft


## Level 5: Association

railway_projects ─┐
                  ├── user_projects
users ────────────┘

## Level 6: Redundant (Can be removed)

railway_projects
  ├── junction_box_summary
  ├── cable_summary
  ├── terminal_summary
  ├── group_summary
  ├── terminal_header_summary
  ├── choke_summary
  └── resistor_summary

## Level 7: Legacy/Unused

railway_projects
  └── station_master (UNUSED)

(Standalone unused tables)
  ├── deleted_projects_backup
  ├── flask_sessions
  ├── user_sessions
  └── sessions


## FOREIGN KEY CHAIN ANALYSIS
## Complete Chain 1: Project → PDF → Approval → User

railway_projects (id: 100)
     ↓ FK: project_id=100
generated_pdf (id: 50)
     ↓ FK: generated_pdf_id=50
approval (id: 5)
     ↓ FK: approver_id=10
users (id: 10)
     ↓ FK: role_id=2
role_master (id: 2)
     ↓ FK: designation_id=3
designation_master (id: 3)


## Complete Chain 2: User → Project Assignment

users (id: 10)
     ↓ FK: user_id=10
user_projects (project_id=100)
     ↓ FK: project_id=100
railway_projects (id: 100)


## CRITICAL PATHS FOR INTEGRITY
## Path 1: Data Entry Flow
User Input → station_drawing/junction_box/cable → 
Trigger update_project_stage_auto() → 
Update railway_projects.stage → 
When stage=10 → Generate PDF → 
Begin approval workflow

## Path 2: Approval Flow
PDF generated → level1_status='pending' → 
Notify Level 1 Approver → 
Approval/Rejection → Update level1_status → 
If approved → level2_status='pending' → 
Repeat for Level 2 & 3 → 
When all approved → stage=14 (drawing_approved)


## INDEX ANALYSIS BY TABLE
## High Traffic Tables (Need Indexes):

railway_projects: (status, stage, updated_date)
users: (role, is_active)
generated_pdf: (project_id, created_at, level1_status, level2_status, level3_status)
approval: (generated_pdf_id, created_at)
notifications: (user_id, pdf_id, is_read)

## Large Data Tables (Need Indexes):
terminal (21,847 rows): (project_id, cable_id, terminal_id)
terminal_header (11,917 rows): (project_id, cable_id, header_type)
cable (3,602 rows): (project_id, cable_id)
from .database import db
from datetime import datetime, timezone, timedelta
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy.orm import validates
from sqlalchemy import Integer, ForeignKey, DateTime, String, Boolean, Text
from passlib.hash import pbkdf2_sha256, scrypt

# IST timezone (still useful for display conversions)
IST = timezone(timedelta(hours=5, minutes=30))

def get_ist_now():
    """Return current datetime in Indian Standard Time (timezone-aware)"""
    return datetime.now(IST)

def get_utc_now():
    """Return current datetime in UTC (timezone-aware)"""
    return datetime.now(timezone.utc)

# User-Project association table
user_projects = db.Table('user_projects',
    db.Column('user_id', db.Integer, db.ForeignKey('users.id'), primary_key=True),
    db.Column('project_id', db.Integer, db.ForeignKey('railway_projects.id'), primary_key=True)
)

class User(UserMixin, db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)

    username = db.Column(db.String(15), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=True)
    password_hash = db.Column(db.String(255), nullable=False)
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    is_active = db.Column(db.Boolean, default=True)
    remarks = db.Column(db.Text, nullable=True)
    mobile_number = db.Column(db.String(15), unique=True, nullable=True)
    role = db.Column(db.String(20), nullable=False, default='user')
    designation = db.Column(db.String(100), nullable=True)

    role_id = db.Column(db.Integer, ForeignKey('role_master.id'), nullable=True)
    designation_id = db.Column(db.Integer, ForeignKey('designation_master.id'), nullable=True)

    role_rel = db.relationship('RoleMaster', backref='users', lazy='joined')
    designation_rel = db.relationship('DesignationMaster', backref='users', lazy='joined')

    projects = db.relationship(
        'Project',
        secondary=user_projects,
        backref=db.backref('assigned_users', lazy='dynamic')
    )

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)
        #self.password_hash = pbkdf2_sha256.hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)
    def check_password_new(self, password):
        """
        Verify password and upgrade hash if needed
        """
        stored_hash = self.password_hash

        # Case 1: old scrypt hash
        if stored_hash.startswith("scrypt:"):
            if scrypt.verify(password, stored_hash):
                # Rehash with pbkdf2_sha256 for future use
                #new_hash = pbkdf2_sha256.hash(password)
                #self.password_hash = new_hash
                #db.session.commit()
                return True
            else:
                return False

        # Case 2: already pbkdf2_sha256 hash
        elif stored_hash.startswith("pbkdf2:sha256:"):
            return pbkdf2_sha256.verify(password, stored_hash)

        # Optional: fallback for other hashes
        else:
            return False
    
    def get_id(self):
        return str(self.id)

    @property
    def role_name(self):
        if self.role_rel:
            return self.role_rel.role_name
        return self.role

    @property
    def designation_name(self):
        if self.designation_rel:
            return self.designation_rel.designation_name
        return self.designation

    @property
    def approval_level(self):
        if self.designation_rel:
            return self.designation_rel.approval_level
        return None

    @property
    def normalized_designation(self):
        if not self.designation:
            return None
        return self.designation.replace(' ', '').lower()

    def is_level(self, level_num):
        if not self.approval_level:
            return False
        return self.approval_level == level_num

    def can_approve_level1(self):
        if not self.role_rel or self.role_rel.role_name != 'approver':
            return False
        if not self.approval_level:
            return False
        return self.approval_level >= 1

    def can_approve_level2(self):
        if not self.role_rel or self.role_rel.role_name != 'approver':
            return False
        if not self.approval_level:
            return False
        return self.approval_level >= 2

    def can_approve_level3(self):
        if not self.role_rel or self.role_rel.role_name != 'approver':
            return False
        if not self.approval_level:
            return False
        return self.approval_level >= 3

    @validates('role')
    def validate_role(self, key, value):
        if value:
            return value.upper()
        return value

    @validates('designation')
    def validate_designation(self, key, value):
        if value:
            value = value.replace(' ', '')
            value = value.upper()
        return value

    @validates('mobile_number')
    def validate_mobile_number(self, key, value):
        if value:
            value = ''.join(filter(str.isdigit, value))
            if len(value) < 10:
                raise ValueError("Mobile number must be at least 10 digits")
        return value

    def __repr__(self):
        return f'<User {self.username} - Role: {self.role_name}>'

class CableBox(db.Model):
    __tablename__ = 'cable_box'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    cable_id = db.Column(db.String(100))
    cable_name = db.Column(db.String(200))
    junction_box = db.Column(db.String(200))
    junction_name = db.Column(db.String(200))
    row = db.Column(db.String(50))
    position = db.Column(db.String(50))
    terminal = db.Column(db.String(100))
    start_no = db.Column(db.String(100))
    cable_type = db.Column(db.String(50), default='cable_box')
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    output = db.Column(db.String(255))
    __table_args__ = (
        db.UniqueConstraint('project_id', 'cable_id', name='uq_cable_box_project_cable'),
    )

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'cable_id': self.cable_id,
            'cable_name': self.cable_name,
            'junction_box': self.junction_box,
            'junction_name': self.junction_name,
            'row': self.row,
            'position': self.position,
            'terminal': self.terminal,
            'start_no': self.start_no,
            'cable_type': self.cable_type,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_date else None,
            'output' : self.output
        }

class Project(db.Model):
    __tablename__ = 'railway_projects'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False, unique=True)
    description = db.Column(db.Text)
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    updated_date = db.Column(db.DateTime(timezone=True), default=get_utc_now, onupdate=get_utc_now)
    station_id = db.Column(db.String(100), nullable=True)
    stage = db.Column(db.Integer, nullable=True)
    status = db.Column(db.String(50), default='drawing_in_progress')
    junction_data = db.Column(db.Text, nullable=True)
    junction_boxes = db.relationship('JunctionBox', backref='project', lazy=True)

    def __repr__(self):
        return f'<Project {self.id}: {self.name}>'

    def get_latest_pdf_status(self):
        latest_pdf = GeneratedPDF.query.filter_by(project_id=self.id).order_by(GeneratedPDF.created_at.desc()).first()
        if latest_pdf:
            return latest_pdf.get_approval_status()
        return None

    def update_junction_data(self):
        import json
        junctions = JunctionBox.query.filter_by(project_id=self.id).all()
        junction_list = [
            {
                "junction_id": jb.junction_id,
                "junction_name": jb.junction_name
            }
            for jb in junctions
        ]
        self.junction_data = json.dumps(junction_list) if junction_list else None
        db.session.commit()
        return junction_list

    def update_status_from_latest_pdf(self):
        latest_pdf = GeneratedPDF.query.filter_by(project_id=self.id).order_by(GeneratedPDF.created_at.desc()).first()
        if latest_pdf:
            pdf_status = latest_pdf.get_approval_status()
            status_mapping = {
                'level1_pending': 'level1_pending',
                'level2_pending': 'level2_pending',
                'level3_pending': 'level3_pending',
                'approved': 'approved',
                'rejected': 'rejected'
            }
            if pdf_status in status_mapping:
                self.status = status_mapping[pdf_status]
            else:
                self.status = 'pdf_generated'
        else:
            has_drawing_data = (
                StationDrawing.query.filter_by(project_id=self.id).first() is not None or
                JunctionBox.query.filter_by(project_id=self.id).first() is not None or
                Cable.query.filter_by(project_id=self.id).first() is not None
            )
            if has_drawing_data:
                self.status = 'drawing_in_progress'
            else:
                self.status = 'new_project'
        return self.status

    def update_stage(self, step_number):
        if step_number is not None and step_number > 0:
            if self.stage is None or step_number > self.stage:
                self.stage = step_number
                self.updated_date = get_utc_now()
                db.session.commit()
                return True
        return False

    def set_pdf_generated_stage(self):
        self.stage = 10
        self.updated_date = get_utc_now()
        db.session.commit()
        return True

class StationDrawing(db.Model):
    __tablename__ = 'station_drawing'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    checksum = db.Column(db.String(100))
    station_id = db.Column(db.String(100))
    diagram_name = db.Column(db.String(200))
    station_name = db.Column(db.String(200))
    station_code = db.Column(db.String(50))
    version = db.Column(db.String(50))
    date = db.Column(db.String(100))
    drawn_by = db.Column(db.String(200))
    checked_by = db.Column(db.String(200))
    division = db.Column(db.String(200))
    zone = db.Column(db.String(200))
    total_sheet = db.Column(db.String(50))
    designation1 = db.Column(db.String(200))
    designation2 = db.Column(db.String(200))
    designation3 = db.Column(db.String(200))
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'diagram_name': self.diagram_name if hasattr(self, 'diagram_name') else None,
            'station_name': self.station_name if hasattr(self, 'station_name') else None,
            'station_code': self.station_code if hasattr(self, 'station_code') else None,
            'version': self.version if hasattr(self, 'version') else None,
            'drawn_by': self.drawn_by if hasattr(self, 'drawn_by') else None,
            'checked_by': self.checked_by if hasattr(self, 'checked_by') else None,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_date else None
        }

class JunctionBox(db.Model):
    __tablename__ = 'junction_box'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    station_id = db.Column(db.String(100))
    junction_id = db.Column(db.String(100))
    junction_name = db.Column(db.String(200))
    latitude = db.Column(db.String(100))
    longitude = db.Column(db.String(100))
    junction_size = db.Column(db.String(100))
    junction_row = db.Column(db.String(100))
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    status = db.Column(db.Integer, default=1)

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'station_id': self.station_id,
            'junction_id': self.junction_id,
            'junction_name': self.junction_name,
            'latitude': self.latitude,
            'longitude': self.longitude,
            'junction_size': self.junction_size,
            'junction_row': self.junction_row,
            'status': self.status,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_date else None
        }

class JunctionApproval(db.Model):
    __tablename__ = 'junction_approval'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    generated_pdf_id = db.Column(db.Integer, db.ForeignKey('generated_pdf.id'), nullable=True)
    junction_box_id = db.Column(db.Integer, db.ForeignKey('junction_box.id'), nullable=False)

    level1_status = db.Column(db.String(20), default='pending')
    level2_status = db.Column(db.String(20), default='pending')
    level3_status = db.Column(db.String(20), default='pending')

    level1_approver_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    level2_approver_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    level3_approver_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)

    level1_approval_date = db.Column(db.DateTime(timezone=True), nullable=True)
    level2_approval_date = db.Column(db.DateTime(timezone=True), nullable=True)
    level3_approval_date = db.Column(db.DateTime(timezone=True), nullable=True)

    rejection_reason = db.Column(db.Text, nullable=True)

    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, onupdate=get_utc_now)

    project = db.relationship('Project', backref='junction_approvals')
    generated_pdf = db.relationship('GeneratedPDF', backref='junction_approvals')
    junction_box = db.relationship('JunctionBox', backref='approvals')
    level1_approver = db.relationship('User', foreign_keys=[level1_approver_id])
    level2_approver = db.relationship('User', foreign_keys=[level2_approver_id])
    level3_approver = db.relationship('User', foreign_keys=[level3_approver_id])

    def __repr__(self):
        return f'<JunctionApproval {self.id} for Junction {self.junction_box_id}>'

    def get_approval_status(self):
        if self.level3_status == 'approved':
            return 'approved'
        elif self.level3_status == 'rejected' or self.level2_status == 'rejected' or self.level1_status == 'rejected':
            return 'rejected'
        elif self.level2_status == 'approved':
            return 'level3_pending'
        elif self.level1_status == 'approved':
            return 'level2_pending'
        else:
            return 'level1_pending'

    def is_fully_approved(self):
        return (
            self.level1_status == 'approved' and
            self.level2_status == 'approved' and
            self.level3_status == 'approved'
        )

class Cable(db.Model):
    __tablename__ = 'cable'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    cable_id = db.Column(db.String(100))
    cable_name = db.Column(db.String(200))
    junction_box = db.Column(db.String(200))
    junction_name = db.Column(db.String(200))
    row = db.Column(db.String(50))
    position = db.Column(db.String(50))
    terminal = db.Column(db.String(100))
    start_no = db.Column(db.String(100))
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    __table_args__ = (
        db.UniqueConstraint('project_id', 'cable_id', name='uq_cable_project_cable'),
    )

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'cable_id': self.cable_id,
            'cable_name': self.cable_name,
            'junction_box': self.junction_box,
            'junction_name': self.junction_name,
            'row': self.row,
            'position': self.position,
            'terminal': self.terminal,
            'start_no': self.start_no,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_date else None
        }

class Terminal(db.Model):
    __tablename__ = 'terminal'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    cable_id = db.Column(db.String(100))
    terminal_id = db.Column(db.String(100))
    terminal_no = db.Column(db.String(200))
    symbol = db.Column(db.String(100))
    input_left = db.Column(db.String(200))
    input_right = db.Column(db.String(200))
    spare = db.Column(db.String(50))
    input_connected = db.Column(db.String(200))
    output_connected = db.Column(db.String(200))
    input_connected_extra = db.Column(db.String(200))
    output_connected_extra = db.Column(db.String(200))
    output_left = db.Column(db.String(200))
    output_right = db.Column(db.String(200))
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    __table_args__ = (
        db.UniqueConstraint('project_id', 'cable_id', 'terminal_id', name='uq_terminal_project_cable_terminal'),
    )

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'cable_id': self.cable_id,
            'terminal_id': self.terminal_id,
            'terminal_no': self.terminal_no,
            'symbol': self.symbol,
            'input_left': self.input_left,
            'input_right': self.input_right,
            'spare': self.spare,
            'input_connected': self.input_connected,
            'output_connected': self.output_connected,
            'input_connected_extra': self.input_connected_extra,
            'output_connected_extra': self.output_connected_extra,
            'output_left': self.output_left,
            'output_right': self.output_right,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_date else None
        }

class Group(db.Model):
    __tablename__ = 'group_table'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    cable_id = db.Column(db.String(100))
    group_id = db.Column(db.String(100))
    terminal_no = db.Column(db.String(100))
    input_output = db.Column(db.String(100))
    text = db.Column(db.Text)
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'cable_id': self.cable_id,
            'group_id': self.group_id,
            'terminal_no': self.terminal_no,
            'input_output': self.input_output,
            'text': self.text,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_date else None
        }

class TerminalHeader(db.Model):
    __tablename__ = 'terminal_header'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    cable_id = db.Column(db.String(100))
    header_type = db.Column(db.String(100))
    terminal_start = db.Column(db.String(100))
    terminal_end = db.Column(db.String(100))
    input_output = db.Column(db.String(100))
    text = db.Column(db.Text)
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    __table_args__ = (
        db.UniqueConstraint('project_id', 'cable_id', 'header_type', 'terminal_start', 'terminal_end',
                          name='uq_terminal_header_project_cable_type'),
    )

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'cable_id': self.cable_id,
            'header_type': self.header_type,
            'terminal_start': self.terminal_start,
            'terminal_end': self.terminal_end,
            'input_output': self.input_output,
            'text': self.text,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_date else None
        }

class ChokeTable(db.Model):
    __tablename__ = 'choke_table'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    cable_id = db.Column(db.String(100))
    choke_id = db.Column(db.String(100))
    input_terminal = db.Column(db.String(100))
    output_terminal = db.Column(db.String(100))
    terminal_name = db.Column(db.String(200))
    output_type = db.Column(db.String(200))
    output_text = db.Column(db.String(200))
    output_connected = db.Column(db.String(200))
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'cable_id': self.cable_id,
            'choke_id': self.choke_id,
            'input_terminal': self.input_terminal,
            'output_terminal': self.output_terminal,
            'terminal_name': self.terminal_name,
            'output_type': self.output_type,
            'output_text': self.output_text,
            'output_connected': self.output_connected,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_date else None
        }

class ResistorTable(db.Model):
    __tablename__ = 'resistor_table'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    cable_id = db.Column(db.String(100))
    resistor_id = db.Column(db.String(100))
    input_terminal = db.Column(db.String(100))
    output_terminal = db.Column(db.String(100))
    resistor_name = db.Column(db.String(200))
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'cable_id': self.cable_id,
            'resistor_id': self.resistor_id,
            'input_terminal': self.input_terminal,
            'output_terminal': self.output_terminal,
            'resistor_name': self.resistor_name,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_date else None
        }

class Approval(db.Model):
    __tablename__ = "approval"
    id = db.Column(db.Integer, primary_key=True)
    generated_pdf_id = db.Column(db.Integer, db.ForeignKey("generated_pdf.id"), nullable=False)
    level = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), nullable=False)
    remarks = db.Column(db.Text, nullable=True)
    approver_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)

    generated_pdf = db.relationship('GeneratedPDF', backref='approval_history')
    approver = db.relationship('User', foreign_keys=[approver_id])

    def __repr__(self):
        return f"<Approval {self.id} for PDF {self.generated_pdf_id} Level {self.level} - {self.status}>"

    def to_dict(self):
        return {
            'id': self.id,
            'generated_pdf_id': self.generated_pdf_id,
            'level': self.level,
            'status': self.status,
            'remarks': self.remarks,
            'approver': {
                'id': self.approver.id,
                'username': self.approver.username,
                'designation': self.approver.designation
            } if self.approver else None,
            'created_at': self.created_at.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if self.created_at else None
        }

class GeneratedPDF(db.Model):
    __tablename__ = "generated_pdf"
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey("railway_projects.id"), nullable=False)
    version = db.Column(db.Integer, nullable=False, default=1)
    pdf_filename = db.Column(db.String(255), nullable=False)
    xlsx_filename = db.Column(db.String(255), nullable=True)
    file_size = db.Column(db.Integer, nullable=False)
    checksum_md5 = db.Column(db.String(32), nullable=False)
    metadata_checksum = db.Column(db.String(32))
    metadata_data = db.Column(db.Text)
    initial_size_bytes = db.Column(db.Integer)
    final_size_bytes = db.Column(db.Integer)
    metadata_ts_ist = db.Column(db.DateTime)
    station_code = db.Column(db.String(50))
    source_pdf_name = db.Column(db.String(255))
    full_file_md5 = db.Column(db.String(32))
    remarks = db.Column(db.Text)
    checksum_algo = db.Column(db.String(16), default="md5", nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)

    level1_status = db.Column(db.String(20), default='pending')
    level2_status = db.Column(db.String(20), default='pending')
    level3_status = db.Column(db.String(20), default='pending')
    level1_approver_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    level2_approver_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    level3_approver_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    level1_approval_date = db.Column(db.DateTime(timezone=True), nullable=True)
    level2_approval_date = db.Column(db.DateTime(timezone=True), nullable=True)
    level3_approval_date = db.Column(db.DateTime(timezone=True), nullable=True)

    junction_data = db.Column(db.Text, nullable=True)

    level1_approver = db.relationship('User', foreign_keys=[level1_approver_id])
    level2_approver = db.relationship('User', foreign_keys=[level2_approver_id])
    level3_approver = db.relationship('User', foreign_keys=[level3_approver_id])
    signed_status = db.Column(db.Integer)
    project = db.relationship('Project', backref='generated_pdfs')

    @property
    def approval_history_sorted(self):
        return sorted(self.approval_history, key=lambda x: (x.level, x.created_at))

    def can_level1_approve(self):
        return self.level1_status == 'pending'

    def can_level2_approve(self):
        return (self.level1_status == 'approved' and
                self.level2_status == 'pending')

    def can_level3_approve(self):
        return (self.level1_status == 'approved' and
                self.level2_status == 'approved' and
                self.level3_status == 'pending')

    def get_approval_status(self):
        if self.level3_status == 'approved':
            return 'approved'
        elif self.level3_status == 'rejected' or self.level2_status == 'rejected' or self.level1_status == 'rejected':
            return 'rejected'
        elif self.level2_status == 'approved':
            return 'level3_pending'
        elif self.level1_status == 'approved':
            return 'level2_pending'
        else:
            return 'level1_pending'

    def is_fully_approved(self):
        return (self.level1_status == 'approved' and
                self.level2_status == 'approved' and
                self.level3_status == 'approved')

    def can_user_approve(self, user):
        if not user.designation:
            return False
        if user.designation == 'level1':
            return self.can_level1_approve()
        elif user.designation == 'level2':
            return self.can_level2_approve()
        elif user.designation == 'level3':
            return self.can_level3_approve()
        return False

    def get_checksum_display(self):
        if self.is_fully_approved():
            return self.checksum_md5
        return None

    def record_approval(self, level, status, approver_id, remarks=None):
        approval = Approval(
            generated_pdf_id=self.id,
            level=level,
            status=status,
            approver_id=approver_id,
            remarks=remarks
        )
        db.session.add(approval)
        db.session.commit()
        return approval

class Notification(db.Model):
    __tablename__ = 'notifications'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    pdf_id = db.Column(db.Integer, db.ForeignKey('generated_pdf.id'), nullable=False)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    level = db.Column(db.String(20), nullable=False)
    status = db.Column(db.String(20), nullable=False)
    message = db.Column(db.String(500))
    is_read = db.Column(db.Boolean, default=False)
    is_dismissed = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, onupdate=get_utc_now)

    user = db.relationship('User', backref='notifications')
    pdf = db.relationship('GeneratedPDF', backref='notifications')
    project = db.relationship('Project', backref='notifications')

    def __repr__(self):
        return f'<Notification {self.id} - {self.status} for Level {self.level}>'

class StationMaster(db.Model):
    __tablename__ = 'station_master'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    station_id = db.Column(db.String(100))
    station_name = db.Column(db.String(200))
    station_code = db.Column(db.String(50))

    __table_args__ = (
        db.UniqueConstraint('station_id', 'project_id', name='unique_station_project'),
    )

    project = db.relationship('Project', backref='master_stations')

    def __repr__(self):
        return f'<StationMaster {self.station_name} ({self.station_code})>'

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'station_id': self.station_id,
            'station_name': self.station_name,
            'station_code': self.station_code,
            'created_date': self.created_date.astimezone(IST).strftime('%Y-%m-%d %H:%M:%S') if hasattr(self, 'created_date') and self.created_date else None
        }

class RoleMaster(db.Model):
    __tablename__ = 'role_master'
    id = db.Column(db.Integer, primary_key=True)
    role_name = db.Column(db.String(50), nullable=False, unique=True)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, onupdate=get_utc_now)

    def __repr__(self):
        return f'<RoleMaster {self.role_name}>'

class DesignationMaster(db.Model):
    __tablename__ = 'designation_master'
    id = db.Column(db.Integer, primary_key=True)
    designation_name = db.Column(db.String(100), nullable=False, unique=True)
    approval_level = db.Column(db.Integer, nullable=True)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, onupdate=get_utc_now)

    def __repr__(self):
        return f'<DesignationMaster {self.designation_name} - Level {self.approval_level}>'

class JunctionBoxSummary(db.Model):
    __tablename__ = 'junction_box_summary'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), unique=True, nullable=False)
    total_junction_boxes = db.Column(db.Integer, default=0, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)
    project = db.relationship('Project', backref=db.backref('junction_box_summary', uselist=False))

class CableSummary(db.Model):
    __tablename__ = 'cable_summary'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), unique=True, nullable=False)
    total_cables = db.Column(db.Integer, default=0, nullable=False)
    total_rows = db.Column(db.Integer, default=0, nullable=False)
    total_junctions = db.Column(db.Integer, default=0, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)
    project = db.relationship('Project', backref=db.backref('cable_summary', uselist=False))

class TerminalSummary(db.Model):
    __tablename__ = 'terminal_summary'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), unique=True, nullable=False)
    total_terminals = db.Column(db.Integer, default=0, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)
    project = db.relationship('Project', backref=db.backref('terminal_summary', uselist=False))

class GroupSummary(db.Model):
    __tablename__ = 'group_summary'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), unique=True, nullable=False)
    total_groups = db.Column(db.Integer, default=0, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)
    project = db.relationship('Project', backref=db.backref('group_summary', uselist=False))

class TerminalHeaderSummary(db.Model):
    __tablename__ = 'terminal_header_summary'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), unique=True, nullable=False)
    total_terminal_headers = db.Column(db.Integer, default=0, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)
    project = db.relationship('Project', backref=db.backref('terminal_header_summary', uselist=False))

class ChokeSummary(db.Model):
    __tablename__ = 'choke_summary'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), unique=True, nullable=False)
    total_chokes = db.Column(db.Integer, default=0, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)
    project = db.relationship('Project', backref=db.backref('choke_summary', uselist=False))

class ResistorSummary(db.Model):
    __tablename__ = 'resistor_summary'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), unique=True, nullable=False)
    total_resistors = db.Column(db.Integer, default=0, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)
    project = db.relationship('Project', backref=db.backref('resistor_summary', uselist=False))

class CableRowConfig(db.Model):
    __tablename__ = 'cable_row_config'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, nullable=False)
    junction_box_id = db.Column(db.String(100), nullable=False)
    row_number = db.Column(db.Integer, nullable=False)
    location_row_name = db.Column(db.String(50), nullable=False)
    cable_type = db.Column(db.String(50), nullable=False)
    number_of_cables = db.Column(db.Integer, nullable=False)
    is_draft = db.Column(db.Boolean, default=True)
    draft_version = db.Column(db.Integer, default=1)
    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, onupdate=get_utc_now)

    __table_args__ = (
        db.UniqueConstraint(
            'project_id', 'junction_box_id', 'row_number', 'is_draft', 'draft_version',
            name='uq_project_jb_row_draft_version'
        ),
    )

    def __repr__(self):
        return f'<CableRowConfig {self.junction_box_id} Row {self.row_number}>'

    def to_dict(self):
        return {
            'id': self.id,
            'project_id': self.project_id,
            'junction_box_id': self.junction_box_id,
            'row_number': self.row_number,
            'location_row_name': self.location_row_name,
            'cable_type': self.cable_type,
            'number_of_cables': self.number_of_cables,
            'is_draft': self.is_draft,
            'draft_version': self.draft_version,
            'created_at': self.created_at.astimezone(IST).isoformat() if self.created_at else None,
            'updated_at': self.updated_at.astimezone(IST).isoformat() if self.updated_at else None
        }

class CableLocationAddition(db.Model):
    __tablename__ = 'cable_location_addition'
    id = db.Column(db.Integer, primary_key=True)
    project_id = db.Column(db.Integer, db.ForeignKey('railway_projects.id'), nullable=False)
    junction_count = db.Column(db.Integer, default=0, nullable=False)
    is_draft = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, onupdate=get_utc_now, nullable=False)
    project = db.relationship('Project', backref=db.backref('cable_location_additions', lazy=True))

# CTR models
class CTRUpload(db.Model):
    __tablename__ = 'ctr_upload'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)

    filename = db.Column(db.String(255), nullable=False)
    stored_filename = db.Column(db.String(255), nullable=True)
    file_path = db.Column(db.String(500), nullable=True)
    file_size = db.Column(db.Integer, nullable=True)
    station_id = db.Column(db.String(100), nullable=True)
    checksum_md5 = db.Column(db.String(32), nullable=True, index=True)

    status_id = db.Column(db.Integer, db.ForeignKey('status_master.id'), default=1)
    station_name = db.Column(db.String(255))
    version = db.Column(db.Integer, default=1, nullable=False)
    is_latest_version = db.Column(db.Boolean, default=True, nullable=False)

    current_approval_level = db.Column(db.Integer, default=1, nullable=False)
    is_fully_approved = db.Column(db.Boolean, default=False, nullable=False)
    fully_approved_at = db.Column(db.DateTime(timezone=True), nullable=True)

    sent_for_approval = db.Column(db.Boolean, default=False, nullable=False)
    sent_for_approval_at = db.Column(db.DateTime(timezone=True), nullable=True)

    upload_date = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)
    parent_version_id = db.Column(db.Integer, db.ForeignKey('ctr_upload.id'), nullable=True)

    generated_pdf = db.Column(db.String(255), nullable=True)
    pdf_generated_date = db.Column(db.DateTime(timezone=True), nullable=True)

    admin_approved = db.Column(db.Boolean, default=False)
    admin_approved_at = db.Column(db.DateTime(timezone=True), nullable=True)
    admin_approved_by = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    admin_approval_notes = db.Column(db.Text, nullable=True)

    user = db.relationship('User', foreign_keys=[user_id], backref=db.backref('ctr_uploads', lazy=True))
    admin_approver = db.relationship('User', foreign_keys=[admin_approved_by], backref='admin_approved_ctr_uploads')
    status = db.relationship('StatusMaster', foreign_keys=[status_id], backref='ctr_uploads')
    parent_version = db.relationship('CTRUpload', remote_side=[id], backref='child_versions', uselist=False)
    sign_document = db.Column(db.String(255))
    name = db.Column(db.String(255))
    is_deleted = db.Column(db.Integer, default=0)
    
    def __repr__(self):
        return f'<CTRUpload {self.id}: {self.filename} (v{self.version})>'

class CTRSummary(db.Model):
    __tablename__ = 'ctr_summary'
    id = db.Column(db.Integer, primary_key=True)
    ctr_upload_id = db.Column(db.Integer, db.ForeignKey('ctr_upload.id'), nullable=False)
    station_id = db.Column(db.String(100), nullable=True)
    station = db.Column(db.String(200))
    project = db.Column(db.String(200))
    designation1 = db.Column(db.String(100))
    designation2 = db.Column(db.String(100))
    designation3 = db.Column(db.String(100))
    station_name = db.Column(db.String(200))
    junction_name = db.Column(db.String(200))
    station_code = db.Column(db.String(100))
    zone = db.Column(db.String(100))
    division = db.Column(db.String(100))
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    no_of_rows = db.Column(db.Integer, nullable=True)
    no_of_terminal_per_row = db.Column(db.Integer, nullable=True)

    ctr_upload = db.relationship('CTRUpload', backref='summary_list', lazy=True)

    def __repr__(self):
        return f'<CTRSummary {self.id}: {self.station}>'

class CTRDiagram(db.Model):
    __tablename__ = 'ctr_diagram'
    id = db.Column(db.Integer, primary_key=True)
    ctr_upload_id = db.Column(db.Integer, db.ForeignKey('ctr_upload.id'), nullable=False)

    terminal_no = db.Column(db.String(100), nullable=False)
    positive = db.Column(db.String(200))
    fuse_input_left = db.Column(db.String(200))
    fuse_input_right = db.Column(db.String(200))
    fuse_output_left = db.Column(db.String(200))
    fuse_output_right = db.Column(db.String(200))
    function = db.Column(db.String(200))
    capsule_input_left = db.Column(db.String(200))
    capsule_input_right = db.Column(db.String(200))
    capsule_output_left = db.Column(db.String(200))
    capsule_output_right = db.Column(db.String(200))
    negative = db.Column(db.String(200))
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)

    ctr_upload = db.relationship('CTRUpload', backref='diagram_list', lazy=True)

    def __repr__(self):
        return f'<CTRDiagram {self.id}: {self.terminal_no}>'

class CTRRowDetail(db.Model):
    __tablename__ = 'ctr_row_detail'
    id = db.Column(db.Integer, primary_key=True)
    ctr_upload_id = db.Column(db.Integer, db.ForeignKey('ctr_upload.id'), nullable=False)

    row_marker = db.Column(db.String(100), nullable=False)
    terminal_no = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text)
    cable_name = db.Column(db.String(200))
    cable_core_start = db.Column(db.String(100))
    cable_core_end = db.Column(db.String(100))
    block_size = db.Column(db.String(50))
    color = db.Column(db.String(100))
    created_date = db.Column(db.DateTime(timezone=True), default=get_utc_now)

    ctr_upload = db.relationship('CTRUpload', backref='rowdetail_list', lazy=True)

    def __repr__(self):
        return f'<CTRRowDetail {self.id}: {self.row_marker} - {self.terminal_no}>'

class CTRApproval(db.Model):
    __tablename__ = 'ctr_approval'
    id = db.Column(db.Integer, primary_key=True)
    ctr_upload_id = db.Column(db.Integer, db.ForeignKey('ctr_upload.id', ondelete='CASCADE'))
    approver_role_id = db.Column(db.Integer, db.ForeignKey('role_master.id'))
    approver_user_id = db.Column(db.Integer, db.ForeignKey('users.id'))
    approval_level = db.Column(db.Integer, nullable=False)
    approval_status = db.Column(db.String(20), nullable=False, default='pending')
    comments = db.Column(db.Text)
    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, onupdate=get_utc_now)

    approver_role = db.relationship('RoleMaster', backref='ctr_approvals', lazy=True)
    approver_user = db.relationship('User', backref='ctr_approvals', lazy=True)
    ctr_upload = db.relationship('CTRUpload', backref='approval_history', lazy=True)

    def __repr__(self):
        return f'<CTRApproval {self.id} - Level {self.approval_level} - {self.approval_status}>'

class CTRApprovalHistory(db.Model):
    __tablename__ = 'ctr_approval_history'
    id = db.Column(db.Integer, primary_key=True)
    ctr_upload_id = db.Column(db.Integer, db.ForeignKey('ctr_upload.id', ondelete='CASCADE'), nullable=False)

    action = db.Column(db.String(50), nullable=False)
    action_level = db.Column(db.Integer, nullable=False)
    action_details = db.Column(db.Text, nullable=True)

    action_by_user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    action_by_role_id = db.Column(db.Integer, db.ForeignKey('role_master.id'), nullable=True)

    action_date = db.Column(db.DateTime(timezone=True), default=get_utc_now, nullable=False)

    previous_status_id = db.Column(db.Integer, db.ForeignKey('status_master.id'), nullable=True)
    new_status_id = db.Column(db.Integer, db.ForeignKey('status_master.id'), nullable=True)

    version_number = db.Column(db.Integer, nullable=False)

    ctr_upload = db.relationship('CTRUpload', backref='approval_histories', lazy=True)
    action_by_user = db.relationship('User', foreign_keys=[action_by_user_id], backref='ctr_approval_histories')
    action_by_role = db.relationship('RoleMaster', foreign_keys=[action_by_role_id], backref='ctr_approval_histories')
    previous_status = db.relationship('StatusMaster', foreign_keys=[previous_status_id], backref='previous_status_histories')
    new_status = db.relationship('StatusMaster', foreign_keys=[new_status_id], backref='new_status_histories')

    def __repr__(self):
        return f'<CTRApprovalHistory {self.id}: Upload {self.ctr_upload_id} - {self.action} at L{self.action_level} by {self.action_by_user_id}>'

class StatusMaster(db.Model):
    __tablename__ = 'status_master'
    id = db.Column(db.Integer, primary_key=True)
    status_code = db.Column(db.String(50), unique=True, nullable=False, index=True)
    status_name = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text, nullable=True)
    category = db.Column(db.String(50), nullable=True)
    is_active = db.Column(db.Boolean, default=True)
    sequence = db.Column(db.Integer, default=0)

    created_at = db.Column(db.DateTime(timezone=True), default=get_utc_now)
    updated_at = db.Column(db.DateTime(timezone=True), default=get_utc_now, onupdate=get_utc_now)

    def __repr__(self):
        return f'<StatusMaster {self.id}: {self.status_code} - {self.status_name}>'

    def to_dict(self):
        return {
            'id': self.id,
            'status_code': self.status_code,
            'status_name': self.status_name,
            'description': self.description,
            'category': self.category,
            'is_active': self.is_active,
            'sequence': self.sequence
        }
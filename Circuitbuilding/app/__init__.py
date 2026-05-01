import os
import sys
import logging
import redis
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime, timedelta, timezone

from flask import Flask, session, request, current_app, jsonify
from flask_login import LoginManager, current_user
from flask_session import Session


from .database import db
from .routes import bp as main_bp
from .schemas import SHEETS, HEADER_HINTS

# ================= TIMEZONE =================
IST = timezone(timedelta(hours=5, minutes=30))

def ist_now():
    return datetime.now(IST)


# ================= LOGGER SETUP =================
def setup_logging(app):
    #log_dir = Path("logs")
    log_dir = Path(os.path.dirname(__file__)) / "logs"
    log_dir.mkdir(exist_ok=True)

    log_file = log_dir / "app.log"

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        "%Y-%m-%d %H:%M:%S"
    )

    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=5 * 1024 * 1024,  # 5 MB
        backupCount=5
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.INFO)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.INFO)

    app.logger.setLevel(logging.INFO)
    app.logger.addHandler(file_handler)
    app.logger.addHandler(console_handler)

    app.logger.info("🔥 Logging initialized")


# ================= ENV LOADER =================
def load_environment():
    try:
        from dotenv import load_dotenv
        base_dir = Path(__file__).resolve().parent.parent
        env_file = base_dir / ".env"

        if env_file.exists():
            load_dotenv(env_file)
        else:
            print("⚠️ .env not found, using system environment")

    except Exception as e:
        print("⚠️ dotenv load skipped:", e)


# ================= APP FACTORY =================
def create_app():
    app = Flask(__name__, static_folder="static", template_folder="templates")

    setup_logging(app)
    app.logger.info("🚀 create_app() started")

    load_environment()

    # ================= CORE CONFIG =================
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "Saltriver@123")
    app.config["SQLALCHEMY_DATABASE_URI"] = (
        "postgresql://postgres:railway123@localhost:5432/postgrestest"
    )
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
    app.config["SQLALCHEMY_ECHO"] = False

    app.config["ENV"] = os.environ.get("FLASK_ENV", "production")
    app.config["DEBUG"] = os.environ.get("FLASK_DEBUG", "0") == "1"

    # ================= FTP CONFIG =================
    app.config.update(
        FTP_ENABLED=os.environ.get("FTP_ENABLED", "False").lower() == "true",
        FTP_HOST=os.environ.get("FTP_HOST", ""),
        FTP_PORT=int(os.environ.get("FTP_PORT", 21)),
        FTP_USERNAME=os.environ.get("FTP_USERNAME", ""),
        FTP_PASSWORD=os.environ.get("FTP_PASSWORD", ""),
        FTP_UPLOAD_DIR=os.environ.get("FTP_UPLOAD_DIR", "/uploads/"),
        FTP_USE_SFTP=os.environ.get("FTP_USE_SFTP", "False").lower() == "true",
        FTP_TIMEOUT=int(os.environ.get("FTP_TIMEOUT", 30)),
    )

    # ================= CTR CONFIG =================
    app.config.update(
        CTR_UPLOAD_FOLDER=os.environ.get("CTR_UPLOAD_FOLDER", "uploads_ctr"),
        CTR_PDF_FOLDER=os.environ.get("CTR_PDF_FOLDER", "static/ctr_pdfs"),
        MAX_CONTENT_LENGTH=int(os.environ.get("MAX_CONTENT_LENGTH", 10 * 1024 * 1024)),
    )

    # ================= SESSION CONFIG =================
    app.permanent_session_lifetime = timedelta(minutes=30)
    
    app.config.update(
        SESSION_TYPE="filesystem",
        SESSION_FILE_DIR=r"C:\temp\flask_sessions",
        SESSION_PERMANENT=True,
        SESSION_COOKIE_NAME="railway_session",
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SECURE=False,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_REFRESH_EACH_REQUEST=True,
    )

   
    '''
    app.config.update(
        SESSION_TYPE="redis",
        SESSION_REDIS=redis.Redis(host="localhost", port=6379),

        SESSION_PERMANENT=True,
        SESSION_COOKIE_NAME="railway_session",
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SECURE=False,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_REFRESH_EACH_REQUEST=True,
    )
    '''     

    # ================= INIT EXTENSIONS =================
    db.init_app(app)
    Session(app)

    # ================= DB CLEANUP =================
    @app.teardown_appcontext
    def shutdown_session(exception=None):
        try:
            if exception:
                app.logger.error("DB Rollback due to exception", exc_info=True)
                db.session.rollback()
        except Exception:
            pass
        finally:
            db.session.remove()

    # ================= LOGIN MANAGER =================
    login_manager = LoginManager()
    login_manager.login_view = "main.login"
    login_manager.login_message = "Please log in to continue."
    login_manager.session_protection = "basic"
    login_manager.init_app(app)

    from .models import User

    @login_manager.user_loader
    def load_user(user_id):
        try:
            return User.query.get(int(user_id))
        except Exception:
            app.logger.exception("User load failed")
            return None

    # ================= REQUEST LOGGING =================
    @app.before_request
    def log_request():
        app.logger.info(
            f"Request: {request.method} {request.path} | IP: {request.remote_addr}"
        )

        if current_user.is_authenticated:
            session.modified = True

    # ================= GLOBAL ERROR HANDLER =================
    @app.errorhandler(Exception)
    def handle_exception(e):
        app.logger.exception("Unhandled Exception occurred")

        return jsonify({
            "status": "error",
            "message": "Internal Server Error"
        }), 500

    # ================= JINJA + BLUEPRINT =================
    app.jinja_env.globals["SHEETS"] = SHEETS
    app.jinja_env.globals["HEADER_HINTS"] = HEADER_HINTS
    app.register_blueprint(main_bp)

    # ================= INIT DB + FOLDERS =================
    with app.app_context():
        try:
            db.create_all()
            app.logger.info("✅ Database ready")

            for folder in (
                app.config["CTR_UPLOAD_FOLDER"],
                app.config["CTR_PDF_FOLDER"],
                os.path.join(app.static_folder, "images"),
            ):
                os.makedirs(folder, exist_ok=True)
                app.logger.info(f"📁 Ensured folder: {folder}")

        except Exception:
            app.logger.exception("❌ Startup failure")

    app.logger.info("✅ App initialized successfully")

    return app
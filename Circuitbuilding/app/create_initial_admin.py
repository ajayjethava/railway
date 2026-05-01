import sys
import os
print("🔍 DEBUG: Current directory:", os.getcwd())
print("🔍 DEBUG: Script location:", __file__)
# Add the correct paths
current_dir = os.path.dirname(os.path.abspath(__file__))
frontend_dir = os.path.dirname(os.path.dirname(current_dir)) # Goes to .../frontend
sys.path.insert(0, frontend_dir)
sys.path.insert(0, current_dir)
print("🔍 DEBUG: Python path:")
for path in sys.path:
    print(" -", path)
def create_initial_admin():
    try:
        print("🔍 DEBUG: Trying to import create_app...")
        from Circuitbuilding.app import create_app
        print("✅ DEBUG: Successfully imported create_app from Circuitbuilding.app")
    except ImportError as e:
        print(f"❌ DEBUG: Failed to import from Circuitbuilding.app: {e}")
        try:
            # Try relative import
            from app import create_app
            print("✅ DEBUG: Successfully imported create_app from app")
        except ImportError as e:
            print(f"❌ DEBUG: Failed to import from app: {e}")
            return
   
    try:
        print("🔍 DEBUG: Trying to import models...")
        from Circuitbuilding.app.models import db, User, Project
        print("✅ DEBUG: Successfully imported models from Circuitbuilding.app.models")
    except ImportError as e:
        print(f"❌ DEBUG: Failed to import from Circuitbuilding.app.models: {e}")
        try:
            # Try relative import
            from models import db, User, Project
            print("✅ DEBUG: Successfully imported models from models")
        except ImportError as e:
            print(f"❌ DEBUG: Failed to import from models: {e}")
            return
   
    app = create_app()
   
    with app.app_context():
        # FIXED: Drop and recreate all tables to ensure schema matches current model definitions
        # This resolves the VARCHAR(120) truncation error on password_hash (model uses VARCHAR(255))
        # Safe for initial setup; will wipe any existing data (re-run if needed after setup)
        print("🔍 DEBUG: Dropping all tables to sync schema...")
        db.drop_all()
        print("🔍 DEBUG: Creating all tables with current model schema...")
        db.create_all()
        print("✅ DEBUG: Tables recreated successfully")
       
        print("🔍 DEBUG: In app context, checking database...")
       
        # Check if admin user already exists
        admin = db.session.query(User).filter_by(username='8780926980').first()
        if not admin:
            print("🔍 DEBUG: Admin user not found, creating...")
           
            # Create a default project if none exists
            project = db.session.query(Project).first()
            if not project:
                print("🔍 DEBUG: No projects found, creating default project...")
                project = Project(name="Default Project", description="Initial project")
                db.session.add(project)
                db.session.commit()
                print("Created default project")
           
            # Create admin user
            admin = User(
                username='8780926980',
                role='admin'
            )
            admin.set_password('admin123')
           
            # Assign the default project to admin
            if project:
                admin.projects.append(project)
           
            db.session.add(admin)
            db.session.commit()
            print("✅ Admin user created successfully!")
            print(" Mobile: 8780926980")
            print(" Password: admin123")
        else:
            print("ℹ️ Admin user already exists")
       
        # Show all users
        users = db.session.query(User).all()
        print(f"\n📋 Total users: {len(users)}")
        for user in users:
            project_names = [p.name for p in user.projects]
            print(f" - {user.username} ({user.role}) - Projects: {', '.join(project_names) if project_names else 'None'}")
if __name__ == "__main__":
    create_initial_admin()
import sqlite3
import os
import sys

# ==============================================================================
# CONFIGURATION - READ CAREFULLY
# ==============================================================================
# This is the standard server path found in your code.
# If you are using a local database, change this path.
DB_PATH = r"H:\Public AZ\DailyReportAutomate\ktib_ops_database.db"

# ==============================================================================
# THE REPAIR SCRIPT
# ==============================================================================
def revert_to_app3_version():
    if not os.path.exists(DB_PATH):
        print(f"❌ ERROR: Database not found at: {DB_PATH}")
        print("Please check the path in the script.")
        return

    print(f"⚡ Connecting to database: {DB_PATH}")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    try:
        print("1. Checking current columns...")
        cursor.execute("PRAGMA table_info(tasks)")
        columns = [info[1] for info in cursor.fetchall()]
        
        if "sort_order" not in columns:
            print("✅ Database is already compatible with App 3. No action needed.")
            return

        print("⚠️  'sort_order' column found. Removing it now...")

        # 1. Rename the existing table (with the extra column)
        print("2. Backing up existing table...")
        cursor.execute("ALTER TABLE tasks RENAME TO tasks_backup_v4")

        # 2. Create the OLD table structure (App 3 version - 15 columns)
        print("3. Recreating App 3 table structure...")
        cursor.execute("""
            CREATE TABLE tasks (
                task_id TEXT, 
                report_date TEXT, 
                os_user TEXT,
                employee TEXT, 
                department TEXT, 
                description TEXT,
                priority TEXT, 
                planned_time TEXT, 
                dependencies TEXT,
                morning_notes TEXT, 
                status TEXT, 
                percent_complete INTEGER,
                results TEXT, 
                reason_delay TEXT, 
                next_steps TEXT,
                PRIMARY KEY (task_id, report_date, os_user)
            )
        """)

        # 3. Copy data from Backup to New Table (Ignoring sort_order)
        print("4. Migrating data (This preserves your tasks)...")
        cursor.execute("""
            INSERT INTO tasks 
            SELECT 
                task_id, report_date, os_user, employee, department, description,
                priority, planned_time, dependencies, morning_notes, status, 
                percent_complete, results, reason_delay, next_steps
            FROM tasks_backup_v4
        """)

        # 4. Verify data count
        cursor.execute("SELECT count(*) FROM tasks")
        new_count = cursor.fetchone()[0]
        cursor.execute("SELECT count(*) FROM tasks_backup_v4")
        old_count = cursor.fetchone()[0]

        if new_count == old_count:
            print(f"5. Verification Successful: {new_count} tasks restored.")
            # 5. Drop the backup table
            cursor.execute("DROP TABLE tasks_backup_v4")
            conn.commit()
            print("\n✅ SUCCESS! Database reverted.")
            print("   App 3 will now work for all employees.")
        else:
            print(f"❌ ERROR: Data count mismatch ({old_count} vs {new_count}). Rolling back.")
            conn.rollback()

    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: {e}")
        print("Changes have been rolled back.")
        conn.rollback()
    finally:
        conn.close()
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    revert_to_app3_version()
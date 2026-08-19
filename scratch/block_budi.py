import sqlite3

db_path = r'd:\new project\export_tools_app\app\database\zinlog_users.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Check if status column exists, if not add it
cursor.execute("PRAGMA table_info(users)")
columns = [col[1] for col in cursor.fetchall()]
if 'status' not in columns:
    cursor.execute("ALTER TABLE users ADD COLUMN status TEXT DEFAULT 'active'")
    print("Added 'status' column to users table.")

# Set all users to active by default
cursor.execute("UPDATE users SET status = 'active' WHERE status IS NULL")

# Block user Budi Santoso (username: budi)
cursor.execute("UPDATE users SET status = 'blocked' WHERE username = 'budi' OR name LIKE '%Budi Santoso%'")
conn.commit()

# Verify
cursor.execute("SELECT id, name, username, email, role, status FROM users")
rows = cursor.fetchall()
print("\nUpdated Users Status:")
for r in rows:
    print(f"- ID: {r[0]} | Nama: {r[1]} | Username: {r[2]} | Status: {r[5]}")

conn.close()

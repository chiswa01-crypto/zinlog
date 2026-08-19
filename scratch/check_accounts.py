import sqlite3

db_path = r'd:\new project\export_tools_app\app\database\zinlog_users.db'
conn = sqlite3.connect(db_path)
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

cursor.execute('SELECT id, name, username, email, role, created_at FROM users')
rows = cursor.fetchall()

print(f"Total akun terdaftar di database: {len(rows)}\n")
for r in rows:
    print(f"- ID: {r['id']} | Nama: {r['name']} | Username: {r['username']} | Email: {r['email']} | Role: {r['role']}")

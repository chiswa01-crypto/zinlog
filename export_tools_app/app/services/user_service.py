import os
import sqlite3
import random
from datetime import datetime, timedelta
from werkzeug.security import generate_password_hash, check_password_hash
from app.services.email_service import send_verification_otp_email

DB_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'database'))
DB_PATH = os.path.join(DB_DIR, 'zinlog_users.db')

def get_db_connection():
    """Membuka koneksi database SQLite dan mengembalikan row sebagai dictionary"""
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Inisialisasi tabel users dan password_resets serta membuat akun default admin"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Tabel Pengguna
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            username TEXT UNIQUE NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT DEFAULT 'user',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 2. Tabel Verifikasi Reset Kata Sandi (OTP)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS password_resets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            email TEXT NOT NULL,
            otp_code TEXT NOT NULL,
            expires_at TIMESTAMP NOT NULL,
            is_used INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users (id)
        )
    ''')
    
    # Akun default admin (Password: 1111)
    cursor.execute("SELECT id FROM users WHERE username = 'admin' OR email = 'admin@zinlog.com'")
    admin = cursor.fetchone()
    default_pwd_hash = generate_password_hash('1111')
    if not admin:
        cursor.execute('''
            INSERT INTO users (name, username, email, password_hash, role)
            VALUES (?, ?, ?, ?, ?)
        ''', ('Administrator Zinlog', 'admin', 'admin@zinlog.com', default_pwd_hash, 'admin'))
        print("[AUTH] Akun default admin berhasil dibuat (Username: admin, Password: 1111)")
    else:
        cursor.execute("UPDATE users SET password_hash = ? WHERE id = ?", (default_pwd_hash, admin['id']))
        
    conn.commit()
    conn.close()

def create_user(name: str, username: str, email: str, password: str):
    """Mendaftarkan pengguna baru"""
    username = username.strip().lower()
    email = email.strip().lower()
    name = name.strip()
    
    if not username or not email or not password or not name:
        return False, "Semua kolom wajib diisi.", None
        
    if len(password) < 4:
        return False, "Kata sandi minimal 4 karakter.", None

    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id FROM users WHERE username = ?", (username,))
    if cursor.fetchone():
        conn.close()
        return False, f"Username '{username}' sudah digunakan. Silakan pilih username lain.", None
        
    cursor.execute("SELECT id FROM users WHERE email = ?", (email,))
    if cursor.fetchone():
        conn.close()
        return False, f"Email '{email}' sudah terdaftar. Silakan gunakan email lain atau login.", None
        
    pwd_hash = generate_password_hash(password)
    try:
        cursor.execute('''
            INSERT INTO users (name, username, email, password_hash, role)
            VALUES (?, ?, ?, ?, ?)
        ''', (name, username, email, pwd_hash, 'user'))
        conn.commit()
        user_id = cursor.lastrowid
        conn.close()
        return True, "Akun berhasil dibuat! Silakan masuk.", {'id': user_id, 'username': username, 'name': name, 'email': email}
    except Exception as e:
        conn.close()
        return False, f"Gagal membuat akun: {str(e)}", None

def authenticate_user(identifier: str, password: str):
    """Otentikasi login pengguna"""
    if not identifier or not password:
        return False, None, "Email/Username dan Kata Sandi wajib diisi."
        
    clean_id = identifier.strip().lower()
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT id, name, username, email, password_hash, role, status
        FROM users
        WHERE username = ? OR email = ?
    ''', (clean_id, clean_id))
    
    user = cursor.fetchone()
    conn.close()
    
    if not user:
        return False, None, "Akun tidak ditemukan. Periksa kembali Email atau Username Anda."
        
    # Cek apakah akun dalam status diblokir
    if user['status'] == 'blocked':
        return False, None, "Akun ini telah dinonaktifkan / DIBLOKIR oleh Administrator. Silakan hubungi admin."
        
    if check_password_hash(user['password_hash'], password):
        return True, {
            'id': user['id'],
            'name': user['name'],
            'username': user['username'],
            'email': user['email'],
            'role': user['role']
        }, "Login berhasil!"
    else:
        return False, None, "Kata sandi salah. Silakan coba lagi atau gunakan Lupa Kata Sandi."

def get_user_by_id(user_id: int):
    """Mengambil data profil pengguna berdasarkan ID"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, username, email, role, created_at FROM users WHERE id = ?", (user_id,))
    user = cursor.fetchone()
    conn.close()
    return dict(user) if user else None

def _mask_email(email: str) -> str:
    """Sensor email untuk privasi: e.g. a***n@zinlog.com"""
    if not email or '@' not in email:
        return email
    parts = email.split('@')
    name_part = parts[0]
    domain_part = parts[1]
    if len(name_part) <= 2:
        masked_name = name_part[0] + "***"
    else:
        masked_name = name_part[0] + "***" + name_part[-1]
    return f"{masked_name}@{domain_part}"

def create_password_reset_otp(identifier: str):
    """
    Menghasilkan kode OTP 6 digit dan mengirimkannya ke alamat email akun terdaftar.
    Mengembalikan (success: bool, message: str, masked_email: str, otp_code: str)
    """
    if not identifier:
        return False, "Silakan masukkan Email atau Username Anda.", "", ""
        
    clean_id = identifier.strip().lower()
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, name, username, email, status FROM users WHERE username = ? OR email = ?", (clean_id, clean_id))
    user = cursor.fetchone()
    
    if not user:
        conn.close()
        return False, f"Akun dengan Email atau Username '{identifier}' tidak ditemukan.", "", ""
        
    if user['status'] == 'blocked':
        conn.close()
        return False, "Akun ini telah diblokir dan tidak dapat melakukan reset kata sandi.", "", ""
        
    # Generate 6-digit random OTP
    otp_code = f"{random.randint(100000, 999999)}"
    expires_at = datetime.now() + timedelta(minutes=10)
    
    cursor.execute('''
        INSERT INTO password_resets (user_id, email, otp_code, expires_at, is_used)
        VALUES (?, ?, ?, ?, 0)
    ''', (user['id'], user['email'], otp_code, expires_at.strftime('%Y-%m-%d %H:%M:%S')))
    
    conn.commit()
    conn.close()
    
    # Kirim email verifikasi
    masked_email = _mask_email(user['email'])
    send_verification_otp_email(user['email'], otp_code, user['name'])
    
    return True, f"Kode verifikasi telah dikirim ke {masked_email}", masked_email, otp_code

def verify_otp_and_reset_password(identifier: str, otp_code: str, new_password: str):
    """
    Memvalidasi kode OTP email dan mengatur ulang kata sandi pengguna.
    Mengembalikan (success: bool, message: str)
    """
    if not identifier or not otp_code or not new_password:
        return False, "Semua kolom verifikasi wajib diisi."
        
    if len(new_password) < 4:
        return False, "Kata sandi baru minimal 4 karakter."
        
    clean_id = identifier.strip().lower()
    clean_otp = otp_code.strip()
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, username, email FROM users WHERE username = ? OR email = ?", (clean_id, clean_id))
    user = cursor.fetchone()
    
    if not user:
        conn.close()
        return False, "Akun tidak ditemukan."
        
    # Cari OTP yang masih berlaku dan belum digunakan
    cursor.execute('''
        SELECT id, otp_code, expires_at, is_used
        FROM password_resets
        WHERE user_id = ? AND is_used = 0
        ORDER BY id DESC LIMIT 1
    ''', (user['id'],))
    
    reset_rec = cursor.fetchone()
    if not reset_rec:
        conn.close()
        return False, "Tidak ada permintaan reset kata sandi aktif. Silakan minta kode baru."
        
    # Cek kecocokan OTP
    if reset_rec['otp_code'] != clean_otp:
        conn.close()
        return False, "Kode verifikasi salah. Periksa kembali kode yang dikirim ke email Anda."
        
    # Cek kadaluarsa (10 menit)
    try:
        exp_time = datetime.strptime(reset_rec['expires_at'], '%Y-%m-%d %H:%M:%S')
        if datetime.now() > exp_time:
            conn.close()
            return False, "Kode verifikasi telah kadaluarsa (lebih dari 10 menit). Silakan minta kode baru."
    except Exception:
        pass
        
    # Sukses: Update kata sandi dan tandai OTP telah digunakan
    new_hash = generate_password_hash(new_password)
    cursor.execute("UPDATE users SET password_hash = ? WHERE id = ?", (new_hash, user['id']))
    cursor.execute("UPDATE password_resets SET is_used = 1 WHERE id = ?", (reset_rec['id'],))
    
    conn.commit()
    conn.close()
    
    return True, f"Kata sandi untuk akun '{user['username']}' berhasil diubah! Silakan masuk dengan kata sandi baru."

def reset_password(identifier: str, new_password: str):
    """Langsung mereset kata sandi (helper / bypass)"""
    if not identifier or not new_password:
        return False, "Kolom identifier dan kata sandi baru wajib diisi."
    if len(new_password) < 4:
        return False, "Kata sandi baru minimal 4 karakter."
    clean_id = identifier.strip().lower()
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username FROM users WHERE username = ? OR email = ?", (clean_id, clean_id))
    user = cursor.fetchone()
    if not user:
        conn.close()
        return False, "Akun tidak ditemukan."
    new_hash = generate_password_hash(new_password)
    cursor.execute("UPDATE users SET password_hash = ? WHERE id = ?", (new_hash, user['id']))
    conn.commit()
    conn.close()
    return True, f"Kata sandi untuk akun '{user['username']}' berhasil diubah!"

# Inisialisasi DB saat modul dimuat
init_db()

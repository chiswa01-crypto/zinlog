import os

class Config:
    """Konfigurasi utama aplikasi Export Tools App."""
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'export-tools-app-secret-key-2026'
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # Maksimum ukuran upload file: 16MB

import os
from app import create_app

app = create_app()

if __name__ == '__main__':
    # Memastikan folder uploads tersedia saat aplikasi pertama kali dijalankan
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True, use_reloader=False, host='127.0.0.1', port=5000)

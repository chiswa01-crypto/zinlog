import os
from flask import Flask, redirect, url_for

def create_app(config_class='config.Config'):
    """Factory function untuk inisialisasi aplikasi Flask."""
    app = Flask(__name__)
    app.config.from_object(config_class)

    # Memastikan folder uploads sudah terbentuk
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

    # Secret Key untuk Flask Session
    app.secret_key = app.config.get('SECRET_KEY', 'zinlog-smart-logistics-secret-key-2026')

    # Registrasi Flask Blueprints
    from app.routes.auth import auth_bp, login_required
    from app.routes.npe_peb import npe_bp
    from app.routes.draf_peb import draf_bp
    from app.routes.cipl import cipl_bp
    from app.routes.compare import compare_bp
    from app.routes.bc40 import bc40_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(npe_bp, url_prefix='/npe-peb')
    app.register_blueprint(draf_bp, url_prefix='/draf-peb')
    app.register_blueprint(cipl_bp, url_prefix='/cipl')
    app.register_blueprint(compare_bp, url_prefix='/compare')
    app.register_blueprint(bc40_bp, url_prefix='/bc40')

    @app.route('/zinlog')
    @app.route('/zinlog/')
    @login_required
    def zinlog():
        from flask import render_template
        return render_template('scorplog.html')

    @app.route('/scorplog')
    @app.route('/scorplog/')
    def scorplog():
        return redirect(url_for('zinlog'))

    # Default route (Homepage) me-redirect ke /login
    @app.route('/')
    def index():
        return redirect(url_for('auth.login'))

    return app

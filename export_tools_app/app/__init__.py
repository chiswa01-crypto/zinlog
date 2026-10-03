import os
from flask import Flask, redirect, url_for

def create_app(config_class='config.Config'):
    """Factory function untuk inisialisasi aplikasi Flask."""
    app = Flask(__name__)
    app.config.from_object(config_class)

    # Memastikan folder uploads sudah terbentuk
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

    # Secret Key untuk Flask Session (Di-reset untuk membersihkan cookie lama)
    app.secret_key = 'zinlog-fresh-session-key-v20261004-unlocked-all'

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

    from flask import session

    @app.before_request
    def clear_legacy_freeze_flashes():
        if '_flashes' in session:
            flashes = session.get('_flashes', [])
            cleaned = [
                f for f in flashes
                if 'freeze' not in str(f[1]).lower() and 'difreeze' not in str(f[1]).lower()
            ]
            if cleaned:
                session['_flashes'] = cleaned
            else:
                session.pop('_flashes', None)

    @app.after_request
    def add_no_cache_headers(response):
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response

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

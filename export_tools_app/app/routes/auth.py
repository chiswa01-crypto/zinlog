import functools
from flask import Blueprint, render_template, request, redirect, url_for, flash, session, g
from app.services.user_service import (
    authenticate_user, create_user, reset_password, get_user_by_id,
    create_password_reset_otp, verify_otp_and_reset_password
)

auth_bp = Blueprint('auth', __name__)

def login_required(view):
    """Decorator untuk memproteksi endpoint yang mewajibkan login"""
    @functools.wraps(view)
    def wrapped_view(**kwargs):
        if 'user_id' not in session:
            flash('Silakan masuk terlebih dahulu untuk mengakses workspace.', 'warning')
            return redirect(url_for('auth.login', next=request.url))
        return view(**kwargs)
    return wrapped_view

@auth_bp.before_app_request
def load_logged_in_user():
    """Muat data pengguna dari sesi ke variabel global context g.user"""
    user_id = session.get('user_id')
    if user_id is None:
        g.user = None
    else:
        g.user = get_user_by_id(user_id)

@auth_bp.route('/login', methods=['GET', 'POST'])
@auth_bp.route('/login/', methods=['GET', 'POST'])
def login():
    if session.get('user_id'):
        return redirect(url_for('zinlog'))
        
    if request.method == 'POST':
        identifier = request.form.get('identifier', '').strip()
        password = request.form.get('password', '')
        
        success, user, message = authenticate_user(identifier, password)
        if success:
            session.clear()
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['name'] = user['name']
            session['role'] = user['role']
            flash(f"Selamat datang kembali, {user['name']}!", 'success')
            
            next_page = request.args.get('next')
            if next_page:
                return redirect(next_page)
            return redirect(url_for('zinlog'))
        else:
            flash(message, 'danger')
            return render_template('login.html', identifier=identifier)
            
    return render_template('login.html')

@auth_bp.route('/register', methods=['GET', 'POST'])
@auth_bp.route('/register/', methods=['GET', 'POST'])
def register():
    if session.get('user_id'):
        return redirect(url_for('zinlog'))
        
    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        username = request.form.get('username', '').strip()
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        confirm_password = request.form.get('confirm_password', '')
        
        if password != confirm_password:
            flash("Konfirmasi kata sandi tidak cocok. Pastikan kedua kata sandi sama.", "danger")
            return render_template('register.html', name=name, username=username, email=email)
            
        success, message, user = create_user(name, username, email, password)
        if success:
            flash(message, "success")
            return redirect(url_for('auth.login'))
        else:
            flash(message, "danger")
            return render_template('register.html', name=name, username=username, email=email)
            
    return render_template('register.html')

@auth_bp.route('/forgot-password', methods=['GET', 'POST'])
@auth_bp.route('/forgot-password/', methods=['GET', 'POST'])
def forgot_password():
    # Step 1: Permintaan Kirim Kode Verifikasi ke Email
    # Step 2: Input Kode OTP & Set Kata Sandi Baru
    step = request.args.get('step', 'request')
    identifier = request.args.get('identifier', '')
    masked_email = request.args.get('masked_email', '')
    
    if request.method == 'POST':
        action = request.form.get('action', 'request_otp')
        
        if action == 'request_otp':
            identifier = request.form.get('identifier', '').strip()
            success, message, masked_email, otp_code = create_password_reset_otp(identifier)
            
            if success:
                flash(message, "success")
                return render_template('forgot_password.html', step='verify', identifier=identifier, masked_email=masked_email)
            else:
                flash(message, "danger")
                return render_template('forgot_password.html', step='request', identifier=identifier)
                
        elif action == 'verify_and_reset':
            identifier = request.form.get('identifier', '').strip()
            otp_code = request.form.get('otp_code', '').strip()
            new_password = request.form.get('new_password', '')
            confirm_password = request.form.get('confirm_password', '')
            masked_email = request.form.get('masked_email', '')
            
            if new_password != confirm_password:
                flash("Konfirmasi kata sandi baru tidak cocok.", "danger")
                return render_template('forgot_password.html', step='verify', identifier=identifier, masked_email=masked_email)
                
            success, message = verify_otp_and_reset_password(identifier, otp_code, new_password)
            if success:
                flash(message, "success")
                return redirect(url_for('auth.login'))
            else:
                flash(message, "danger")
                return render_template('forgot_password.html', step='verify', identifier=identifier, masked_email=masked_email)
                
    return render_template('forgot_password.html', step=step, identifier=identifier, masked_email=masked_email)

@auth_bp.route('/logout')
@auth_bp.route('/logout/')
def logout():
    session.clear()
    flash("Anda telah berhasil keluar (Logged out).", "info")
    return redirect(url_for('auth.login'))

import os
import sys
sys.path.insert(0, r'd:\new project\export_tools_app')

# Test user_service functions
from app.services.user_service import (
    init_db, create_user, authenticate_user, reset_password, get_user_by_id
)

print("=== 1. TEST DATABASE INITIALIZATION ===")
init_db()

print("\n=== 2. TEST DEFAULT ADMIN LOGIN ===")
ok, user, msg = authenticate_user('admin', 'admin123')
print(f"Login admin with username: {ok} -> {msg}")
assert ok == True
assert user['username'] == 'admin'

ok_email, user_email, msg = authenticate_user('admin@zinlog.com', 'admin123')
print(f"Login admin with email: {ok_email} -> {msg}")
assert ok_email == True

print("\n=== 3. TEST BAD PASSWORD ===")
ok_bad, _, msg_bad = authenticate_user('admin', 'wrongpassword')
print(f"Bad password test: {ok_bad} -> {msg_bad}")
assert ok_bad == False

print("\n=== 4. TEST USER REGISTRATION ===")
ok_reg, msg_reg, new_u = create_user("Budi Santoso", "budi", "budi@example.com", "secret123")
print(f"Create user Budi: {ok_reg} -> {msg_reg}")

# Duplicate username check
ok_dup, msg_dup, _ = create_user("Budi Duplikat", "budi", "other@example.com", "secret123")
print(f"Duplicate username test: {ok_dup} -> {msg_dup}")
assert ok_dup == False

# Login new user
ok_budi, u_budi, _ = authenticate_user("budi", "secret123")
print(f"Login Budi: {ok_budi} (Name: {u_budi['name']})")
assert ok_budi == True

print("\n=== 5. TEST FORGOT / RESET PASSWORD ===")
ok_reset, msg_reset = reset_password("budi", "newsecret2026")
print(f"Reset password Budi: {ok_reset} -> {msg_reset}")
assert ok_reset == True

# Login with old password should fail
ok_old, _, _ = authenticate_user("budi", "secret123")
assert ok_old == False

# Login with new password should succeed
ok_new, _, _ = authenticate_user("budi", "newsecret2026")
print(f"Login Budi with new password: {ok_new}")
assert ok_new == True

print("\n=== 6. TEST FLASK APP CLIENT & SESSION PROTECTIONS ===")
from app import create_app
app = create_app()
client = app.test_client()

# 6.1 Unauthenticated access to /zinlog should redirect to /login
r_unauth = client.get('/zinlog', follow_redirects=False)
print("Unauthenticated GET /zinlog status:", r_unauth.status_code, "->", r_unauth.headers.get('Location'))
assert r_unauth.status_code == 302
assert '/login' in r_unauth.headers.get('Location')

# 6.2 POST /login with valid admin credentials
r_login = client.post('/login', data={'identifier': 'admin', 'password': 'admin123'}, follow_redirects=True)
print("POST /login valid credentials status:", r_login.status_code)
assert r_login.status_code == 200
assert 'ZINLOG Logistics Workspace' in r_login.get_data(as_text=True)
print("Successfully logged in and viewed /zinlog workspace!")

# 6.3 GET /logout
r_logout = client.get('/logout', follow_redirects=True)
print("GET /logout status:", r_logout.status_code)
assert r_logout.status_code == 200
assert 'Production Workspace Ready' in r_logout.get_data(as_text=True)

print("\n🎉 ALL AUTHENTICATION TESTS PASSED SUCCESSFULLY! 🎉")

# CODE PROTECTION & LOCK RULE

## Status: ACTIVE LOCK

To maintain stability and prevent unintentional regression, all codebase components in this repository (`d:\new project`) are strictly **LOCKED** against modifications, with the single explicit exception of the **Draf PEB** module.

---

## 🔒 LOCKED COMPONENTS (STRICTLY READ-ONLY)

The following modules, routes, templates, and core scripts are **LOCKED** and must NOT be edited:

1. **Authentication & Session**:
   - `export_tools_app/app/routes/auth.py`
   - `export_tools_app/app/templates/login.html`
   - `export_tools_app/app/templates/register.html`
   - `export_tools_app/app/templates/forgot_password.html`

2. **CIPL Module**:
   - `export_tools_app/app/routes/cipl.py`
   - `export_tools_app/app/templates/cipl.html`
   - `cipl.py`

3. **Komparasi Module**:
   - `export_tools_app/app/routes/compare.py`
   - `export_tools_app/app/templates/compare.html`
   - `komparasi.py`
   - `modules/export/komparasi.py`

4. **BC 4.0 Module**:
   - `export_tools_app/app/routes/bc40.py`
   - `export_tools_app/app/templates/bc40.html`

5. **NPE PEB Upload Module**:
   - `export_tools_app/app/routes/npe_peb.py`
   - `export_tools_app/app/templates/npe_peb.html`

6. **Core Config & App Initialization**:
   - `export_tools_app/run.py`
   - `export_tools_app/config.py`
   - `export_tools_app/app/__init__.py`
   - `export_tools_app/app/templates/base.html`

---

## 🔓 UNLOCKED COMPONENT (PERMITTED FOR EDITING & REPAIR)

Only the **Draf PEB** module is unlocked for repairs, enhancements, and bug fixes:

- `export_tools_app/app/routes/draf_peb.py`
- `export_tools_app/app/templates/draf_peb.html`
- `npe_peb.py` (if required for Draf PEB generation logic)

---

## 🚀 WEB SERVER STATUS
The Flask web application server (`export_tools_app/run.py`) runs continuously at `http://127.0.0.1:5000` without disruption.

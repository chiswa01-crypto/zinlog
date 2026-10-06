# AGENTS & WORKSPACE INSTRUCTIONS

## 🔒 CODEBASE LOCK POLICY

All modules and scripts in this workspace are **LOCKED** (Read-Only) to preserve working features, **EXCEPT** for the **Draf PEB** module.

### Allowed Files for Editing:
- `export_tools_app/app/routes/draf_peb.py`
- `export_tools_app/app/templates/draf_peb.html`
- `npe_peb.py` (Draft PEB related functions only)

### Locked Files (Do Not Edit):
- `auth.py`, `cipl.py`, `compare.py`, `bc40.py`, `npe_peb.py` (routes & standalone)
- `base.html`, `login.html`, `register.html`, `cipl.html`, `compare.html`, `bc40.html`, `npe_peb.html`
- Core setup: `run.py`, `config.py`, `__init__.py`

### Web Server:
The Flask app must continue running at `http://127.0.0.1:5000`.
